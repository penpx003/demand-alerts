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

**Live and running end to end** as of 2026-07-30.

| Piece | State |
|---|---|
| Office Script | v1.3, runs clean against the real workbook |
| Supabase | migration applied; snapshots and digests persisting |
| Digest bot | deployed on Render; 52 self-checks passing |
| Power Automate | `Weekly Demand Digest` runs Mondays, posts to Teams |

Two things to keep an eye on:

- **Volume thresholds** (`ALERT1_MIN_ABS_DIFF`, `ALERT3_MIN_HISTORICAL_VOLUME`,
  `ALERT3_MIN_FUTURE_VOLUME`) started as placeholder `100`s. Confirm they match your
  volume unit, or the alerts are either noisy or silent — [SETUP.md](SETUP.md) 1.5.
- **The dimension mapping** is resolved at runtime from the header row. It has run
  successfully, but "resolved" is not "resolved correctly" — check the seven
  `Dimension ... ->` lines against the real columns at least once, and after any
  change to the IBP extract layout. [SETUP.md](SETUP.md) 1.2.

Rebuilding from scratch, or picking this up cold? [SETUP.md](SETUP.md) has the full
build, the traps that cost time the first time round, and a condensed checklist at
the end.
