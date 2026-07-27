r"""Self-test for the digest service. No pytest, no network, no LLM.

Run from the digest-bot folder:
    .\.venv\Scripts\python.exe tests\test_digest.py

Covers the things that would silently corrupt a planning digest: numeric parsing
of Power Automate's formatted cells, double counting across aggregation levels,
and percent vs percentage-point rendering. Then checks endpoint auth and the
empty-payload short circuit.
"""
from __future__ import annotations

import datetime as dt
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from digest.alerts import parse_payload, summarise, to_float  # noqa: E402
from digest.digest import build_brief, parse_week_of  # noqa: E402

failures: list[str] = []


def check(label: str, condition: bool, detail: str = "") -> None:
    if condition:
        print(f"  ok   {label}")
    else:
        print(f"  FAIL {label} {detail}")
        failures.append(label)


print("\nnumeric parsing (Power Automate returns formatted cell text)")
for raw, expected in [
    ("-42.0%", -0.42),          # percentage cell
    ("18,400.00", 18400.0),     # thousand separator
    ("1.234,56", 1234.56),      # European convention
    ("(1,200)", -1200.0),       # accounting negative
    ("", None),                 # blank cell
    ("n/a", None),
    ("ALL", None),              # text in a numeric column
    (0.155, 0.155),             # already numeric
]:
    check(f"to_float({raw!r}) == {expected!r}", to_float(raw) == expected, f"got {to_float(raw)!r}")


print("\nweek handling")
check("any weekday snaps back to Monday", parse_week_of("2026-07-29") == dt.date(2026, 7, 27))
check("dd/mm/yyyy accepted", parse_week_of("29/07/2026") == dt.date(2026, 7, 27))
check("missing week_of falls back to this Monday", parse_week_of(None).weekday() == 0)


PAYLOAD = {
    "week_of": "2026-07-29",
    "alerts": {
        # Dict-wrapped, as "List rows present in a table" returns it.
        "alert1": {"value": [
            {"Aggregation Level": "Product", "Market": "IBERIA", "Product": "100234",
             "Product Description": "Orangina 1.5L", "Customer": "ALL", "Week": "W31 26 (27/07)",
             "Final Demand Plan W-1": "44,000", "Current Final Demand Plan": "25,600",
             "Difference": "-18,400", "Absolute Difference": "18,400",
             "Variation %": "-41.8%", "Alert Type": "Forecast Decrease"},
            {"Aggregation Level": "Product-Customer", "Market": "IBERIA", "Product": "100234",
             "Product Description": "Orangina 1.5L", "Customer": "CARREFOUR",
             "Week": "W31 26 (27/07)", "Final Demand Plan W-1": "20,000",
             "Current Final Demand Plan": "0", "Difference": "-20,000",
             "Absolute Difference": "20,000", "Variation %": "-100.0%",
             "Alert Type": "Forecast Removed"},
            {"Aggregation Level": "Product", "Market": "FRANCE", "Product": "100999",
             "Product Description": "Schweppes 33cl", "Customer": "ALL",
             "Week": "W32 26 (03/08)", "Final Demand Plan W-1": "0",
             "Current Final Demand Plan": "9,500", "Difference": "9,500",
             "Absolute Difference": "9,500", "Variation %": "",
             "Alert Type": "New Forecast"},
        ]},
        # Bare list, as it arrives without ?['value'].
        "alert2": [
            {"Aggregation Level": "Product", "Market": "IBERIA", "Product": "100234",
             "Product Description": "Orangina 1.5L", "Customer": "ALL",
             "Accuracy W-1": "61.0%", "Baseline Accuracy W-5 to W-2": "82.0%",
             "Accuracy Difference": "-21.0%", "Bias W-1": "18.0%",
             "Baseline Bias W-5 to W-2": "4.0%", "Bias Signed Difference": "14.0%",
             "Bias Deterioration": "14.0%", "Bias Direction": "Further from zero",
             "Alert Reason": "Accuracy and Bias deterioration",
             "W-1 Actual": "12,300", "Baseline Actual W-5 to W-2": "51,800"},
        ],
        "alert3": [],
        "alert4": [
            {"Aggregation Level": "Product", "Market": "FRANCE", "Product": "100999",
             "Product Description": "Schweppes 33cl", "Customer": "ALL",
             "Sales History Last 4 Weeks": "40,000",
             "Statistical Forecast Snapshot Last 4 Weeks": "41,000",
             "Final Demand Plan Snapshot Last 4 Weeks": "58,000",
             "Statistical Forecast Absolute Error": "3,600",
             "Final Demand Plan Absolute Error": "18,400",
             "Statistical Forecast Accuracy": "91.0%",
             "Final Demand Plan Accuracy": "54.0%",
             "Accuracy Difference": "37.0%", "FVA Classification": "Negative FVA"},
        ],
    },
}

print("\npayload parsing")
parsed = parse_payload(PAYLOAD)
counts = [len(v) for v in parsed.values()]
check("both dict-wrapped and bare list shapes parse", counts == [3, 1, 0, 1], f"got {counts}")

# A quiet week still delivers a table, carrying one placeholder row. If that row
# were parsed as an alert, the digest would report a blank combination as news.
placeholder = {"Aggregation Level": "No alerts this week", "Market": "", "Product": "",
               "Product Description": "", "Customer": "", "Variation %": "",
               "Absolute Difference": "", "Alert Type": ""}
check(
    "placeholder row from an empty alert table is skipped",
    len(parse_payload({"alerts": {"alert1": [placeholder]}})["alert1"]) == 0,
)
check(
    "a real row is still parsed alongside it",
    len(parse_payload({"alerts": {"alert1": [placeholder, PAYLOAD["alerts"]["alert1"]["value"][0]]}})["alert1"]) == 1,
)

stats = {k: summarise(k, rows) for k, rows in parsed.items()}
a1 = stats["alert1"]

print("\naggregation (the double-counting trap)")
check("levels split correctly", a1.product_rows == 2 and a1.product_customer_rows == 1)
check(
    "total uses Product rows only (27,900, not 47,900)",
    a1.volume_total == 27900.0,
    f"got {a1.volume_total}",
)
check("empty alert summarises to zero", stats["alert3"].total_rows == 0)

print("\nranking")
check("row with blank Variation % is kept, ranked on volume", len(a1.top_product) == 2)
check(
    "New Forecast row survives ranking",
    any(r.alert_type == "New Forecast" for r in a1.top_product),
)
check("worst deterioration ranks first", stats["alert2"].top_product[0].metric == -0.21)

print("\nbrief rendering (percent vs percentage points)")
brief = build_brief(dt.date(2026, 7, 27), "2026-07-20", stats, {})
check("empty alert is stated, not omitted", "no combinations met the criteria" in brief)
check("accuracy difference renders as pp", "-21.0 pp" in brief, "expected '-21.0 pp'")
check("bias deterioration renders as pp", "+14.0 pp" in brief)
check("FVA advantage renders as pp", "+37.0 pp" in brief)
check("accuracy LEVEL still renders as percent", "accuracy W-1 61.0%" in brief)
check("volume of 0 is not rendered as 0.0%", "previous plan 0.0%" not in brief)
check("volume of 0 renders as 0", "previous plan 0" in brief)
check(
    "alert 2 total is labelled 'w-1 actual', not 'volume at risk'",
    "total w-1 actual across Product-level rows" in brief,
)

print("\nTeams HTML rendering")
from digest.digest import to_teams_html  # noqa: E402

sample = "**Headline**\nSales fell 21 percentage points.\n\n**Focus**\n* Review A & B\n* Check C"
rendered = to_teams_html(sample)
check("bold heading becomes <b>", "<b>Headline</b>" in rendered)
check("no literal asterisks survive", "**" not in rendered, rendered[:120])
check("bullets become a list", "<ul>" in rendered and rendered.count("<li>") == 2)
check("list is closed", rendered.endswith("</ul>"))
check("ampersand in a product name is escaped", "A &amp; B" in rendered)
check("paragraphs wrapped", "<p>Sales fell 21 percentage points.</p>" in rendered)

print("\nendpoint")
os.environ["DEMAND_DIGEST_TOKEN"] = "self-test-token"
try:
    from fastapi.testclient import TestClient

    from app.main import app

    client = TestClient(app)
    empty = {"alerts": {"alert1": [], "alert2": [], "alert3": [], "alert4": []}}
    hdr = {"X-Digest-Token": "self-test-token"}

    check("no token -> 401", client.post("/api/demand-digest", json=empty).status_code == 401)
    check(
        "wrong token -> 401",
        client.post("/api/demand-digest", json=empty, headers={"X-Digest-Token": "x"}).status_code
        == 401,
    )
    check(
        "latest is protected too",
        client.get("/api/demand-digest/latest").status_code == 401,
    )
    check(
        "malformed body -> 400",
        client.post("/api/demand-digest", content=b"nope", headers=hdr).status_code == 400,
    )

    r = client.post("/api/demand-digest", json=empty, headers=hdr)
    body = r.json() if r.status_code == 200 else {}
    check("valid token -> 200", r.status_code == 200, r.text[:200])
    check("no alerts skips the LLM entirely", body.get("model") == "none")
    check("brief withheld by default", "brief" not in body)

    r = client.post("/api/demand-digest", json={**empty, "include_brief": True}, headers=hdr)
    check("include_brief returns the brief", "brief" in r.json())
except Exception as e:
    check("endpoint tests ran", False, f"{type(e).__name__}: {e}")

print()
if failures:
    print(f"{len(failures)} FAILED: {', '.join(failures)}")
    sys.exit(1)
print("all checks passed")
