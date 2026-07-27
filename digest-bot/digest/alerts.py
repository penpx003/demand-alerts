"""Parsing and deterministic statistics for the Demand Planning alert tables.

The four alert tables are produced by the `Demand Alerts` Office Script (see
../../office-script) in the `DemandAlertsScripts` workbook, and pushed here by a
weekly Power Automate flow.

Everything numeric in the weekly digest is computed HERE, in Python. The LLM is
only ever asked to narrate a brief built from these figures — it never sees raw
rows and never derives a number. That keeps a planning digest trustworthy.

Aggregation note: the alert tables contain BOTH aggregation levels in one table
(`Product-Customer` and `Product`, the latter with Customer = "ALL"). Totals are
computed from Product-level rows only, otherwise every volume is counted twice.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Iterable

LEVEL_PRODUCT_CUSTOMER = "Product-Customer"
LEVEL_PRODUCT = "Product"

_NON_ALNUM = re.compile(r"[^a-z0-9]+")


def normalise_key(name: str) -> str:
    """Column header -> canonical lookup key.

    "Variation %" -> "variation", "W-1 Actual" -> "w1actual",
    "Baseline Accuracy W-5 to W-2" -> "baselineaccuracyw5tow2".
    """
    return _NON_ALNUM.sub("", (name or "").strip().lower())


def to_float(value: Any) -> float | None:
    """Tolerant numeric conversion.

    Power Automate's "List rows present in a table" returns the FORMATTED cell
    text, so a percentage arrives as "-42.0%" and a volume as "18,400.00" or
    "1.234,56". Blank cells arrive as "". Mirrors the Office Script's toNumber
    so both ends agree.
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).replace(" ", " ").strip()
    if text in ("", "-", "--", "n/a", "N/A"):
        return None

    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]

    percent = text.endswith("%")
    if percent:
        text = text[:-1]

    text = re.sub(r"\s", "", text)

    last_comma, last_dot = text.rfind(","), text.rfind(".")
    if last_comma >= 0 and last_dot >= 0:
        if last_comma > last_dot:            # 1.234,56
            text = text.replace(".", "").replace(",", ".")
        else:                                # 1,234.56
            text = text.replace(",", "")
    elif last_comma >= 0:
        decimals = len(text) - last_comma - 1
        if 0 < decimals <= 2 and text.count(",") == 1:
            text = text.replace(",", ".")    # 1234,56
        else:
            text = text.replace(",", "")     # 1,234,567

    try:
        number = float(text)
    except ValueError:
        return None
    if percent:
        number /= 100.0
    return -number if negative else number


# How a figure is rendered in the brief. The kind comes from the FIELD, never from
# the value's magnitude: a plan of 0 is "0", not "0.0%", and a volume of 1 is "1",
# not "100.0%".
#
# PCT is a level (accuracy of 82%). PP is a DIFFERENCE between two levels, measured
# in percentage points. Rendering a difference as "%" is exactly how "down 21
# percentage points" becomes the meaningless "down 21%", so the distinction is
# enforced in the formatting rather than left to the prompt to remember.
PCT, PP, VOL, TEXT = "pct", "pp", "vol", "text"


@dataclass(frozen=True)
class AlertSpec:
    """How to read one alert table and what its headline figures are."""

    key: str
    title: str
    metric_field: str          # canonical key of the headline metric
    metric_label: str
    metric_kind: str           # PCT (a level) | PP (a difference) | VOL
    volume_field: str          # canonical key of the headline volume
    volume_label: str
    type_field: str | None     # classification column, if any
    bucket_field: str | None   # per-week bucket column (Alert 1 only)
    rank_abs: bool             # rank on |metric| rather than the signed value
    rank_desc: bool            # highest first


ALERT_SPECS: dict[str, AlertSpec] = {
    "alert1": AlertSpec(
        key="alert1",
        title="Final Demand Plan change vs previous cycle (W+1 to W+12)",
        metric_field="variation",
        metric_label="Variation %",
        metric_kind=PCT,
        volume_field="absolutedifference",
        volume_label="Absolute difference",
        type_field="alerttype",
        bucket_field="week",
        rank_abs=True,
        rank_desc=True,
    ),
    "alert2": AlertSpec(
        key="alert2",
        title="Accuracy and Bias deterioration (W-1 vs W-5..W-2)",
        metric_field="accuracydifference",
        metric_label="Accuracy difference",
        metric_kind=PP,
        volume_field="w1actual",
        volume_label="W-1 actual",
        type_field="alertreason",
        bucket_field=None,
        rank_abs=False,
        rank_desc=False,           # most negative (worst deterioration) first
    ),
    "alert3": AlertSpec(
        key="alert3",
        title="Future forecast vs recent sales (W-4..W-1 vs W+1..W+4)",
        metric_field="variation",
        metric_label="Variation %",
        metric_kind=PCT,
        volume_field="absolutedifference",
        volume_label="Absolute difference",
        type_field="alerttype",
        bucket_field=None,
        rank_abs=True,
        rank_desc=True,
    ),
    "alert4": AlertSpec(
        key="alert4",
        title="Statistical Forecast beating the Final Demand Plan (negative FVA, W-4..W-1)",
        metric_field="accuracydifference",
        metric_label="Accuracy advantage of the statistical forecast",
        metric_kind=PP,
        volume_field="saleshistorylast4weeks",
        volume_label="Sales history last 4 weeks",
        type_field="fvaclassification",
        bucket_field=None,
        rank_abs=False,
        rank_desc=True,
    ),
}

# Extra columns worth surfacing in the brief, per alert:
# (canonical column key, label used in the brief, kind).
DETAIL_FIELDS: dict[str, list[tuple[str, str, str]]] = {
    "alert1": [
        ("finaldemandplanw1", "previous plan", VOL),
        ("currentfinaldemandplan", "current plan", VOL),
    ],
    "alert2": [
        ("accuracyw1", "accuracy W-1", PCT),
        ("baselineaccuracyw5tow2", "baseline accuracy", PCT),
        ("biasw1", "bias W-1", PCT),
        ("biasdeterioration", "bias deterioration", PP),
        ("biasdirection", "bias direction", TEXT),
    ],
    "alert3": [
        ("saleslast4weeks", "sales last 4w", VOL),
        ("forecastnext4weeks", "forecast next 4w", VOL),
    ],
    "alert4": [
        ("statisticalforecastaccuracy", "stat accuracy", PCT),
        ("finaldemandplanaccuracy", "FDP accuracy", PCT),
    ],
}

# label -> kind, so the brief can format a detail without guessing.
DETAIL_KIND: dict[str, str] = {
    label: kind for fields in DETAIL_FIELDS.values() for _, label, kind in fields
}


@dataclass
class AlertRow:
    """One normalised alert row, independent of the alert it came from."""

    alert: str
    agg_level: str
    market: str
    product: str
    description: str
    customer: str
    bucket: str                       # future week label for Alert 1, else ""
    metric: float | None
    volume: float | None
    alert_type: str
    details: dict[str, Any] = field(default_factory=dict)

    @property
    def entity_key(self) -> tuple[str, str, str]:
        return (self.market, self.product, self.customer)

    def label(self) -> str:
        who = f"{self.market} / {self.product}"
        if self.description:
            who += f" {self.description}"
        if self.agg_level == LEVEL_PRODUCT_CUSTOMER and self.customer:
            who += f" / {self.customer}"
        if self.bucket:
            who += f" / {self.bucket}"
        return who


def _get(row: dict[str, Any], canonical: dict[str, Any], key: str) -> Any:
    """Fetch a field by canonical key, falling back to the raw header."""
    if key in canonical:
        return canonical[key]
    return row.get(key)


def parse_rows(alert: str, raw_rows: Iterable[dict[str, Any]]) -> list[AlertRow]:
    """Normalise the rows of one alert table as delivered by Power Automate."""
    spec = ALERT_SPECS[alert]
    out: list[AlertRow] = []

    for row in raw_rows:
        if not isinstance(row, dict):
            continue
        canonical = {normalise_key(k): v for k, v in row.items()}

        market_raw = str(_get(row, canonical, "market") or "").strip()
        product_raw = str(_get(row, canonical, "product") or "").strip()
        if not market_raw and not product_raw:
            # An alert table that matched nothing this week still exists (so the
            # Power Automate flow finds it) and carries one placeholder row with
            # no Market and no Product. Skipping it here is what keeps a quiet
            # week reported as "no combinations met the criteria" rather than as
            # one blank alert.
            continue

        level = str(_get(row, canonical, "aggregationlevel") or "").strip()
        if level not in (LEVEL_PRODUCT_CUSTOMER, LEVEL_PRODUCT):
            # Unknown or missing level: infer from the Customer column.
            customer_probe = str(_get(row, canonical, "customer") or "").strip()
            level = LEVEL_PRODUCT if customer_probe.upper() == "ALL" else LEVEL_PRODUCT_CUSTOMER

        details: dict[str, Any] = {}
        for canonical_key, label, kind in DETAIL_FIELDS.get(alert, []):
            value = canonical.get(canonical_key)
            if value in (None, ""):
                continue
            if kind == TEXT:
                details[label] = str(value).strip()
            else:
                numeric = to_float(value)
                if numeric is not None:
                    details[label] = numeric

        out.append(
            AlertRow(
                alert=alert,
                agg_level=level,
                market=market_raw,
                product=product_raw,
                description=str(_get(row, canonical, "productdescription") or "").strip(),
                customer=str(_get(row, canonical, "customer") or "").strip(),
                bucket=(
                    str(canonical.get(spec.bucket_field) or "").strip()
                    if spec.bucket_field
                    else ""
                ),
                metric=to_float(canonical.get(spec.metric_field)),
                volume=to_float(canonical.get(spec.volume_field)),
                alert_type=(
                    str(canonical.get(spec.type_field) or "").strip()
                    if spec.type_field
                    else ""
                ),
                details=details,
            )
        )
    return out


def parse_payload(payload: dict[str, Any]) -> dict[str, list[AlertRow]]:
    """Accepts {"alerts": {"alert1": [...], ...}} or {"alert1": [...], ...}.

    Also tolerates the Power Automate shape where each table arrives wrapped as
    {"value": [...]}, which is what "List rows present in a table" returns.
    """
    container = payload.get("alerts") if isinstance(payload.get("alerts"), dict) else payload
    parsed: dict[str, list[AlertRow]] = {}

    for alert in ALERT_SPECS:
        raw = container.get(alert)
        if isinstance(raw, dict):
            raw = raw.get("value", [])
        if not isinstance(raw, list):
            raw = []
        parsed[alert] = parse_rows(alert, raw)

    return parsed


@dataclass
class AlertStats:
    """Deterministic figures for one alert. The only numbers the LLM may use."""

    key: str
    title: str
    total_rows: int
    product_customer_rows: int
    product_rows: int
    # Sum of the alert's headline volume over Product-level rows only (no double
    # counting). What it MEANS differs per alert — absolute forecast change for
    # Alerts 1 and 3, actual sales covered for Alerts 2 and 4 — so it is always
    # rendered with the alert's own volume_label, never as a generic total.
    volume_total: float
    by_type: dict[str, int]
    top_product: list[AlertRow]
    top_product_customer: list[AlertRow]
    markets: dict[str, int]


def _rank(rows: list[AlertRow], spec: AlertSpec, limit: int) -> list[AlertRow]:
    scored = [r for r in rows if r.metric is not None]
    unscored = [r for r in rows if r.metric is None]

    def score(row: AlertRow) -> float:
        assert row.metric is not None
        return abs(row.metric) if spec.rank_abs else row.metric

    scored.sort(key=score, reverse=spec.rank_desc)
    # Rows with no percentage (new forecast / no recent sales) rank on volume.
    unscored.sort(key=lambda r: abs(r.volume or 0.0), reverse=True)
    return (scored + unscored)[:limit]


def summarise(alert: str, rows: list[AlertRow], top_n: int = 8) -> AlertStats:
    spec = ALERT_SPECS[alert]
    product_rows = [r for r in rows if r.agg_level == LEVEL_PRODUCT]
    pc_rows = [r for r in rows if r.agg_level == LEVEL_PRODUCT_CUSTOMER]

    by_type: dict[str, int] = {}
    for row in rows:
        if row.alert_type:
            by_type[row.alert_type] = by_type.get(row.alert_type, 0) + 1

    markets: dict[str, int] = {}
    for row in product_rows:
        if row.market:
            markets[row.market] = markets.get(row.market, 0) + 1

    return AlertStats(
        key=alert,
        title=spec.title,
        total_rows=len(rows),
        product_customer_rows=len(pc_rows),
        product_rows=len(product_rows),
        volume_total=sum(abs(r.volume or 0.0) for r in product_rows),
        by_type=dict(sorted(by_type.items(), key=lambda kv: kv[1], reverse=True)),
        top_product=_rank(product_rows, spec, top_n),
        top_product_customer=_rank(pc_rows, spec, top_n),
        markets=dict(sorted(markets.items(), key=lambda kv: kv[1], reverse=True)),
    )


def fmt_pct(value: float | None) -> str:
    """A level: 0.82 -> '82.0%'."""
    return "n/a" if value is None else f"{value * 100:.1f}%"


def fmt_pp(value: float | None) -> str:
    """A difference: -0.21 -> '-21.0 pp'. Signed, so direction is unambiguous."""
    return "n/a" if value is None else f"{value * 100:+.1f} pp"


def fmt_vol(value: float | None) -> str:
    return "n/a" if value is None else f"{value:,.0f}"


def fmt_by_kind(value: Any, kind: str) -> str:
    """Single formatting entry point, so the brief never guesses a unit."""
    if kind == PCT:
        return fmt_pct(value)
    if kind == PP:
        return fmt_pp(value)
    if kind == VOL:
        return fmt_vol(value)
    return "n/a" if value is None else str(value)
