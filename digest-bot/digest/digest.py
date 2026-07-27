"""Weekly Demand Planning digest: alert tables in, narrative out.

Flow:
  Power Automate (weekly)
    -> runs the `Demand Alerts` Office Script on DemandAlertsScripts
    -> reads the four output tables
    -> POST /api/demand-digest
    -> this module: normalise -> deterministic stats -> snapshot to Supabase
       -> trend lookup vs previous weeks -> LLM narration -> narrative
    -> Power Automate posts the narrative to the Demand Planning Teams channel

The LLM is given a factual brief and told to narrate it. Every number in that
brief was computed in `alerts.py`. The model is explicitly forbidden from
producing figures that are not in the brief, because a planning digest that
invents a percentage is worse than no digest.
"""
from __future__ import annotations

import datetime as dt
import html
import os
import re
from typing import Any

from supabase import create_client, Client

from .config import Config
from .alerts import (
    ALERT_SPECS,
    DETAIL_KIND,
    LEVEL_PRODUCT,
    LEVEL_PRODUCT_CUSTOMER,
    TEXT,
    AlertRow,
    AlertStats,
    fmt_by_kind,
    fmt_vol,
    parse_payload,
    summarise,
)
from .llm import GroqClient

# How many past weeks to consider when deciding whether an alert is recurring.
TREND_WEEKS = int(os.getenv("DEMAND_TREND_WEEKS", "8"))
# Product-level rows are always stored. Product-Customer rows are the long tail,
# so only the top N per alert are snapshotted — enough for recurrence, bounded
# growth on a free Supabase project.
SNAPSHOT_PC_LIMIT = int(os.getenv("DEMAND_SNAPSHOT_PC_LIMIT", "500"))
TOP_N = int(os.getenv("DEMAND_TOP_N", "8"))


def monday_of(day: dt.date) -> dt.date:
    return day - dt.timedelta(days=day.weekday())


def parse_week_of(value: Any) -> dt.date:
    """Week the digest covers. Defaults to the Monday of the current week."""
    if isinstance(value, str) and value.strip():
        text = value.strip()[:10]
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return monday_of(dt.datetime.strptime(text, fmt).date())
            except ValueError:
                continue
    return monday_of(dt.date.today())


class DemandStore:
    """Supabase persistence for alert snapshots and generated digests."""

    def __init__(self, cfg: Config):
        cfg.require_supabase()  # validated here, not at config load
        self.cfg = cfg
        self.client: Client = create_client(cfg.supabase_url, cfg.supabase_key)

    def save_snapshot(self, week_of: dt.date, parsed: dict[str, list[AlertRow]]) -> int:
        """Upsert this week's rows. Idempotent: a retried flow overwrites."""
        payload: list[dict[str, Any]] = []

        for alert, rows in parsed.items():
            product_rows = [r for r in rows if r.agg_level == LEVEL_PRODUCT]
            pc_rows = [r for r in rows if r.agg_level == LEVEL_PRODUCT_CUSTOMER]
            spec = ALERT_SPECS[alert]
            pc_rows.sort(
                key=lambda r: (
                    abs(r.metric) if (spec.rank_abs and r.metric is not None) else (r.metric or 0.0)
                ),
                reverse=spec.rank_desc,
            )
            for row in product_rows + pc_rows[:SNAPSHOT_PC_LIMIT]:
                payload.append(
                    {
                        "week_of": week_of.isoformat(),
                        "alert": row.alert,
                        "agg_level": row.agg_level,
                        "market": row.market,
                        "product": row.product,
                        "customer": row.customer,
                        "bucket": row.bucket,
                        "metric": row.metric,
                        "volume": row.volume,
                        "alert_type": row.alert_type,
                        "details": row.details,
                    }
                )

        conflict = "week_of,alert,agg_level,market,product,customer,bucket"
        for start in range(0, len(payload), 500):
            self.client.table("demand_alert_snapshots").upsert(
                payload[start : start + 500], on_conflict=conflict
            ).execute()
        return len(payload)

    def recurrence(
        self, alert: str, level: str, since: dt.date
    ) -> dict[tuple[str, str, str], dict[str, Any]]:
        """{(market, product, customer): {weeks_seen, first_seen, last_seen}}."""
        try:
            resp = self.client.rpc(
                "demand_alert_recurrence",
                {"p_alert": alert, "p_level": level, "p_since": since.isoformat()},
            ).execute()
        except Exception as e:  # a missing migration must not kill the digest
            print(f"recurrence lookup skipped for {alert}/{level}: {e}")
            return {}

        return {
            (r.get("market", ""), r.get("product", ""), r.get("customer", "")): r
            for r in (resp.data or [])
        }

    def save_digest(
        self, week_of: dt.date, narrative: str, stats: dict[str, Any], model: str
    ) -> None:
        self.client.table("demand_digests").upsert(
            {
                "week_of": week_of.isoformat(),
                "narrative": narrative,
                "stats": stats,
                "model": model,
            },
            on_conflict="week_of",
        ).execute()

    def previous_digest_week(self, before: dt.date) -> str | None:
        resp = (
            self.client.table("demand_digests")
            .select("week_of")
            .lt("week_of", before.isoformat())
            .order("week_of", desc=True)
            .limit(1)
            .execute()
        )
        data = resp.data or []
        return data[0]["week_of"] if data else None

    def latest_digest(self) -> dict[str, Any] | None:
        resp = (
            self.client.table("demand_digests")
            .select("*")
            .order("week_of", desc=True)
            .limit(1)
            .execute()
        )
        data = resp.data or []
        return data[0] if data else None


def _row_line(
    index: int,
    row: AlertRow,
    metric_kind: str,
    volume_label: str,
    recurring: dict[str, Any] | None,
) -> str:
    metric = fmt_by_kind(row.metric, metric_kind)
    parts = [f"  {index}. {row.label()}: {metric}, {volume_label.lower()} {fmt_vol(row.volume)}"]
    if row.alert_type:
        parts.append(f"[{row.alert_type}]")
    for label, value in row.details.items():
        parts.append(f"{label} {fmt_by_kind(value, DETAIL_KIND.get(label, TEXT))}")
    if recurring:
        parts.append(
            f"(RECURRING: alerted in {recurring['weeks_seen']} of the last {TREND_WEEKS} weeks, "
            f"since {recurring['first_seen']})"
        )
    return " ".join(parts)


def build_brief(
    week_of: dt.date,
    previous_week: str | None,
    stats_by_alert: dict[str, AlertStats],
    recurrence_by_alert: dict[str, dict[tuple[str, str, str], dict[str, Any]]],
) -> str:
    """Compact factual brief. This is the ONLY source of numbers for the LLM."""
    lines: list[str] = [
        f"WEEK OF {week_of.isoformat()} (Monday).",
        f"Previous digest: {previous_week or 'none — this is the first one'}.",
        "",
        "Notes for interpretation:",
        "- Each alert is reported at two levels: Product (customer = ALL) and Product-Customer.",
        "- Totals use Product-level rows only, so nothing is double counted.",
        "- Each volume figure is named. Alerts 1 and 3 report forecast CHANGE volume;",
        "  Alerts 2 and 4 report the ACTUAL SALES covered. They are not comparable.",
        "- Accuracy and bias differences are in PERCENTAGE POINTS, not percent change.",
        "- Positive bias = over-forecast; negative bias = under-forecast.",
        "- 'RECURRING' means the same combination alerted in earlier weeks too.",
        "",
    ]

    for key, spec in ALERT_SPECS.items():
        stats = stats_by_alert[key]
        recurrence = recurrence_by_alert.get(key, {})

        lines.append(f"### {key.upper()} — {stats.title}")
        lines.append(
            f"rows: {stats.total_rows} "
            f"(Product {stats.product_rows}, Product-Customer {stats.product_customer_rows})"
        )
        if stats.total_rows == 0:
            lines.append("no combinations met the criteria this week.")
            lines.append("")
            continue

        lines.append(
            f"total {spec.volume_label.lower()} across Product-level rows: "
            f"{fmt_vol(stats.volume_total)}"
        )
        if stats.by_type:
            lines.append(
                "classification: "
                + ", ".join(f"{name} {count}" for name, count in stats.by_type.items())
            )
        if stats.markets:
            top_markets = list(stats.markets.items())[:5]
            lines.append(
                "most affected markets (Product rows): "
                + ", ".join(f"{name} {count}" for name, count in top_markets)
            )

        recurring_count = sum(
            1 for row in stats.top_product if row.entity_key in recurrence
        )
        lines.append(
            f"of the top {len(stats.top_product)} Product-level items, "
            f"{recurring_count} also alerted in previous weeks."
        )

        lines.append(f"top Product-level items by {spec.metric_label}:")
        for i, row in enumerate(stats.top_product, 1):
            lines.append(
                _row_line(
                    i, row, spec.metric_kind, spec.volume_label,
                    recurrence.get(row.entity_key),
                )
            )

        if stats.top_product_customer:
            pc_recurrence = recurrence_by_alert.get(f"{key}:pc", {})
            lines.append("top Product-Customer items:")
            for i, row in enumerate(stats.top_product_customer, 1):
                lines.append(
                    _row_line(
                        i, row, spec.metric_kind, spec.volume_label,
                        pc_recurrence.get(row.entity_key),
                    )
                )
        lines.append("")

    return "\n".join(lines)


PROMPT = """You are a demand planning analyst writing the weekly alert digest for the \
Demand Planning team. You are given a factual brief produced by an automated \
analysis of SAP IBP data.

Write a short narrative that a planner can read in two minutes and act on.

STRICT RULES:
- Use ONLY the figures in the brief. Never invent, estimate, extrapolate or \
recompute a number. If something is not in the brief, do not mention it.
- Never average accuracy or bias figures, and never add percentages together.
- Accuracy and bias movements are in PERCENTAGE POINTS. Write "down 21 percentage \
points", never "down 21%" or "a -21% deterioration".
- Do not compare a volume from one alert with a volume from another: they measure \
different things and the brief names each one.
- If an alert has zero rows, say so in one clause and move on.
- Call out RECURRING items explicitly — a combination alerting several weeks in \
a row is the most important thing in the digest.
- Refer to products the way the brief does (market / code / description).

STRUCTURE (use these exact headings, no others):
**Headline** - two or three sentences on what matters most this week.
**Forecast changes vs last cycle** - Alert 1.
**Accuracy and bias** - Alert 2.
**Forecast vs recent sales** - Alert 3.
**Where the statistical forecast is winning** - Alert 4, and what it implies about \
manual forecast adjustments.
**Suggested focus this week** - three to five bullets, each naming a specific \
combination from the brief and the action to take.

STYLE:
- Plain business English, no jargon beyond normal demand planning terms.
- Around 400 words. Never exceed 600.
- Bold headings as shown, short paragraphs, bullets only in the final section.
- No tables, no code blocks, no closing pleasantries.

BRIEF:
{brief}
"""


def _empty_narrative(week_of: dt.date) -> str:
    return (
        f"**Headline** - No demand planning alerts were raised for the week of "
        f"{week_of.isoformat()}. Every combination stayed within the configured "
        f"thresholds for forecast change, accuracy and bias, forecast versus recent "
        f"sales, and statistical forecast value added.\n\n"
        f"**Suggested focus this week** - Nothing requires attention from the alert "
        f"run. If this is unexpected, check that the DBAlerts extract was refreshed "
        f"before the script ran."
    )


_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")


def to_teams_html(text: str) -> str:
    """Convert the narrative's limited markdown to the HTML Teams renders.

    Power Automate's "Post message in a chat or channel" treats the message body
    as HTML, not markdown: posting the raw narrative would show literal asterisks
    around every heading. Only what the prompt can produce is handled — bold
    headings, paragraphs and one bullet list.
    """
    out: list[str] = []
    in_list = False

    for raw in text.split("\n"):
        # Escape first: '**' survives escaping, so bold conversion still works
        # afterwards and any stray '<' in a product name is neutralised.
        line = _BOLD_RE.sub(r"<b>\1</b>", html.escape(raw.strip()))

        if not line:
            continue
        if line.startswith("* ") or line.startswith("- "):
            if not in_list:
                out.append("<ul>")
                in_list = True
            out.append(f"<li>{line[2:].strip()}</li>")
            continue
        if in_list:
            out.append("</ul>")
            in_list = False
        out.append(f"<p>{line}</p>")

    if in_list:
        out.append("</ul>")
    return "".join(out)


def generate_narrative(cfg: Config, brief: str) -> tuple[str, str]:
    """Returns (narrative, model_name)."""
    # No reply deadline here — this runs on a weekly schedule, so a full-quality
    # generation is affordable (unlike a Teams Outgoing Webhook, capped at ~5s).
    client = GroqClient(cfg)
    text = client.generate(PROMPT.format(brief=brief), max_output_tokens=1400)
    return text.strip(), cfg.groq_model


def run_digest(payload: dict[str, Any], cfg: Config | None = None) -> dict[str, Any]:
    """Full pipeline for one weekly POST. Returns the narrative and its stats."""
    cfg = cfg or Config.load()
    week_of = parse_week_of(payload.get("week_of"))
    parsed = parse_payload(payload)

    stats_by_alert = {key: summarise(key, rows, TOP_N) for key, rows in parsed.items()}
    total_rows = sum(s.total_rows for s in stats_by_alert.values())

    store: DemandStore | None = None
    snapshot_rows = 0
    previous_week: str | None = None
    recurrence_by_alert: dict[str, dict[tuple[str, str, str], dict[str, Any]]] = {}

    try:
        store = DemandStore(cfg)
        previous_week = store.previous_digest_week(week_of)
        since = week_of - dt.timedelta(weeks=TREND_WEEKS)
        # Recurrence is read BEFORE this week's snapshot is written, so "seen in
        # N previous weeks" never counts the run that is happening right now.
        for key in ALERT_SPECS:
            recurrence_by_alert[key] = store.recurrence(key, LEVEL_PRODUCT, since)
            recurrence_by_alert[f"{key}:pc"] = store.recurrence(
                key, LEVEL_PRODUCT_CUSTOMER, since
            )
        snapshot_rows = store.save_snapshot(week_of, parsed)
    except Exception as e:  # storage problems must not lose the digest
        print(f"snapshot/trend step failed, continuing without history: {e}")

    brief = build_brief(week_of, previous_week, stats_by_alert, recurrence_by_alert)

    if total_rows == 0:
        narrative, model = _empty_narrative(week_of), "none"
    else:
        narrative, model = generate_narrative(cfg, brief)

    stats_json = {
        key: {
            "total_rows": s.total_rows,
            "product_rows": s.product_rows,
            "product_customer_rows": s.product_customer_rows,
            "volume_total": round(s.volume_total, 2),
            "volume_meaning": ALERT_SPECS[key].volume_label,
            "by_type": s.by_type,
        }
        for key, s in stats_by_alert.items()
    }

    if store is not None:
        try:
            store.save_digest(week_of, narrative, stats_json, model)
        except Exception as e:
            print(f"digest not persisted: {e}")

    return {
        "week_of": week_of.isoformat(),
        "previous_week": previous_week,
        "narrative": narrative,
        # Use this one in the Teams action: it renders as formatted text rather
        # than showing literal ** around every heading.
        "narrative_html": to_teams_html(narrative),
        "model": model,
        "total_alert_rows": total_rows,
        "snapshot_rows_stored": snapshot_rows,
        "stats": stats_json,
        "brief": brief,
    }
