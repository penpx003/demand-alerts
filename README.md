# Demand Planning alerts

Two halves of one pipeline over the SAP IBP extract in the `DBAlerts` worksheet of
the `DemandAlertsScripts` workbook.

```
DBAlerts (IBP extract)
        |
        v
office-script/        Excel Automate, TypeScript
  Demand Alerts       -> 4 alert worksheets, two aggregation levels each
        |
        v
digest-bot/           FastAPI on Render, Groq + Supabase
  /api/demand-digest  -> weekly narrative posted to a Teams channel
```

**Building this from scratch? → [SETUP.md](SETUP.md)** — the full step-by-step
guide, from an empty workbook to a narrative arriving in Teams every Monday.

| Folder | What it is |
|---|---|
| [office-script/](office-script/) | The Office Script. Paste `src/DemandAlerts.ts` into Excel → Automate. |
| [digest-bot/](digest-bot/) | The service that narrates the alert tables each week. |

The two are independent: the Office Script is useful on its own, and the bot only
needs the four output tables in whatever way they reach it.

## The four alerts

| # | Alert | Window | Detects |
|---|---|---|---|
| 1 | FDP change vs previous cycle | W+1..W+12 | Material changes to the current Final Demand Plan |
| 2 | Accuracy / Bias deterioration | W-1 vs W-5..W-2 | Performance getting worse |
| 3 | Forecast vs recent sales | W-4..W-1 vs W+1..W+4 | Future plan detached from recent demand |
| 4 | Statistical Forecast vs FDP | W-4..W-1 | Negative forecast value added |

Each produces results at **Product-Customer** and **Product** level, in the same
table, distinguished by the `Aggregation Level` column.

## KPI rules enforced across both halves

These are the rules the whole pipeline is built around, and both halves enforce
them structurally rather than by convention:

- Product-level results aggregate from the **source rows**, never from the
  Product-Customer results.
- Accuracy and bias are **always recalculated** from underlying volumes at the
  exact entity level and period. Source KPI columns are never averaged or summed.
- `Accuracy = 1 - SUM(weekly ABS error) / SUM(weekly Actual)` — errors accumulate
  per week so positive and negative weeks cannot cancel.
- `Bias = SUM(Forecast - Actual) / SUM(Actual)`; positive means over-forecast.
- `Bias deterioration = ABS(current) - ABS(baseline)` — movement away from zero.
- Zero Actual means the KPI is not calculable, not zero.
- The current week is excluded from every historical and future calculation.
- Accuracy and bias *differences* are percentage points, never percent.

## Relationship to the RAG bot

None, beyond using the same free-tier providers (Groq, Supabase, Power Automate,
Teams). Separate repo, separate Render service, separate Supabase tables, separate
Teams team and channel. Nothing here imports from, deploys with, or writes to the
RAG project.

## Status

- Office Script: type-checks clean; **never run against the real workbook**, which
  was not reachable from this machine. Dimension mapping resolves at runtime and
  stops with a clear error if it cannot. Volume thresholds are placeholders.
- Digest bot: parsing, statistics, endpoint auth and a real Groq generation all
  verified locally. **Not yet deployed**; the Supabase migration has not been run
  and the Power Automate flow has not been built.
