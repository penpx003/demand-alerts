# Demand Alerts — Office Script

Production Office Script (Excel Automate) that reads the IBP extract in the
`DBAlerts` worksheet and generates four Demand Planning alerts, each at two
independent aggregation levels (Product-Customer and Product).

## Files

| Path | Purpose |
|---|---|
| `src/DemandAlerts.ts` | **The script.** Copy the whole file into Excel → Automate → New Script. |
| `types/excelscript.d.ts` | Local-only type stubs so the editor can check the code. **Do not paste into Excel.** |
| `tsconfig.json` | Local type-check config. Not used by Excel. |

## Install

1. Open the `DemandAlertsScripts` workbook.
2. **Automate → New Script**.
3. Delete the generated `main` and paste the entire contents of `src/DemandAlerts.ts`.
4. Name it `Demand Alerts` and **Run**.

The script never writes to `DBAlerts`. It deletes and recreates four output sheets:
`Alert 1 FDP Change`, `Alert 2 Accuracy Bias`, `Alert 3 Forecast vs Sales`, `Alert 4 Stat vs FDP`.

## Before the first run — dimension mapping

The IBP CSV export can leave dimension headers blank. On every run the script:

1. tries to identify each dimension from the header row by name,
2. falls back to the `DIM_COL` column-index constants (`A = 0`, current extract = A–G),
3. verifies the resolved Week column actually parses week labels, re-detecting it by
   content if not,
4. **stops with an explicit error** if Market, Product, Customer or Week cannot be
   resolved, or if any required Key Figure is missing (the message lists every one).

Every decision is written to the run log — check it after the first run and pin
`DIM_COL` to the real indexes if the auto-detection picked anything unexpected.

## Key Figures

Required (the run fails and lists any that are missing):

- `Sales History (Shipments)`
- `5 Final Demand Plan (Shipments)` — current future plan
- `5 Final Demand Plan (Shipments) W-1` — previous cycle
- `5 Final Demand Plan (Shipments) Snapshot` — historical performance
- `0.2 Full Stat. Forecast (Snapshot)` — historical performance

Optional, used for **validation only** at Product-Customer + single week:
`... Accuracy`, `... Error`, `... Bias`, `0.2 Full Stat. Forecast Accuracy/Error`.

Headers are matched after trimming and collapsing repeated internal spaces, so
`5  Final Demand Plan (Shipments)` matches — while `... W-1`, `... Snapshot` and
`... Accuracy` stay distinct.

## Calculation rules enforced

- Product-level results are aggregated **from the source records**, never derived
  from the Product-Customer output.
- Accuracy and Bias are **always recalculated** from underlying volumes at the exact
  entity level and period; source KPI columns are never averaged or summed.
- `Accuracy = 1 - SUM(weekly ABS error) / SUM(weekly Forecast)` — the SAP IBP
  convention `ABS(Forecast - Actual) / Forecast`, so the alerts match what
  planners see in IBP. `ACCURACY_DENOMINATOR` switches it to `SUM(Actual)`.
  Weekly errors are
  accumulated per week so positive and negative errors cannot cancel.
- `Bias = SUM(Forecast - Actual) / SUM(Actual)`; positive = overforecast.
- `Bias Deterioration = ABS(current Bias) - ABS(baseline Bias)` — movement away from zero.
- Zero Actual ⇒ KPI not calculable; the entity is excluded and counted in the summary.
- The current week is excluded from every historical and future calculation.
- Week start comes from the date inside the parentheses (`W01 26 (29/12)`), never
  from the week number alone, so ISO year crossings are handled.

## Tuning

All thresholds are constants at the top of `src/DemandAlerts.ts`:

| Constant | Default | Alert |
|---|---|---|
| `ALERT1_PCT_THRESHOLD` | `0.10` | 1 |
| `ALERT1_MIN_ABS_DIFF` | `100` | 1 |
| `ALERT1_FUTURE_HORIZON_WEEKS` | `12` | 1 |
| `ALERT2_ACCURACY_DETERIORATION_THRESHOLD` | `-0.05` | 2 |
| `ALERT2_BIAS_DETERIORATION_THRESHOLD` | `0.05` | 2 |
| `ALERT2_MIN_BASELINE_WEEKS` | `4` | 2 |
| `ALERT3_PCT_THRESHOLD` | `0.15` | 3 |
| `ALERT3_MIN_HISTORICAL_VOLUME` | `100` | 3 |
| `ALERT3_MIN_FUTURE_VOLUME` | `100` | 3 |
| `ALERT4_ACCURACY_DIFF_THRESHOLD` | `0.05` | 4 |
| `HISTORICAL_WEEKS` / `FUTURE_WEEKS` | `5` / `12` | all |

`ALERT1_MIN_ABS_DIFF`, `ALERT3_MIN_HISTORICAL_VOLUME` and `ALERT3_MIN_FUTURE_VOLUME`
are in the volume unit of the extract — set them to your real materiality before
rolling out, or the first run will be noisy.

## Large extracts

Office Scripts caps the payload of a single API call, so the source cannot be read
with one `getValues()` — a big IBP extract fails with *"Range getValues: the
response payload size has exceeded the limit"*.

The script reads the sheet in row blocks sized from `READ_CHUNK_CELLS` (default
`100000` cells), and only across the columns it actually needs. If a block still
exceeds the limit it halves the block and retries automatically, logging each
retry, so the default rarely needs changing. Lower `READ_CHUNK_CELLS` if you see
repeated retries.

Blocks are parsed straight into the in-memory entity maps and then discarded, so
peak memory does not grow with the number of source rows.

## Empty alerts

An alert that matches nothing still gets its worksheet **and its table**, carrying a
single placeholder row with `Aggregation Level` = `No alerts this week` and no Market
or Product. A note below the table explains it.

This matters for automation: Power Automate reads each alert table by name, so a
missing table would fail the whole weekly run on a quiet week. Consumers must skip
rows that have no Market and no Product — the digest bot does.

## Output of a run

`main` returns an `ExecutionSummary` (also printed to the run log): rows read, rows
ignored for invalid dates or numbers, per-alert counts split by aggregation level,
exclusions for insufficient history or zero Actual, source-KPI validation results,
and the list of worksheets created.

## Local type-checking (optional)

Requires Node.js, which is not installed on this machine:

```
npm i -D typescript
npx tsc -p tsconfig.json
```
