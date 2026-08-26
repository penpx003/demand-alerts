/**
 * ============================================================================
 * DEMAND PLANNING ALERTS  —  Microsoft Office Script (Excel Automate)
 * ============================================================================
 * Version  : 1.8
 * Runs on  : Excel for the web / desktop, Automate tab, "New Script"
 * Workbook : DemandAlertsScripts
 * Source   : worksheet "DBAlerts" (IBP CSV extract)
 *
 * Generates four Demand Planning alerts, each at two INDEPENDENT aggregation
 * levels (Product-Customer and Product):
 *
 *   Alert 1 — Final Demand Plan change vs previous cycle (W-1)   -> W+1..W+12
 *   Alert 2 — Accuracy / Bias deterioration                      -> W-1 vs W-5..W-2
 *   Alert 3 — Future Final Demand Plan vs recent Sales History    -> W-4..W-1 vs W+1..W+4
 *   Alert 4 — Statistical Forecast vs Final Demand Plan accuracy  -> W-4..W-1 (negative FVA)
 *
 * DESIGN RULES ENFORCED BY THIS SCRIPT
 *  - Product-level results are aggregated from the ORIGINAL source records,
 *    never derived from the Product-Customer alert output.
 *  - Accuracy and Bias are ALWAYS recalculated from the underlying volumes at
 *    the exact entity level and period analysed. Source Accuracy / Bias / Error
 *    Key Figures are used for VALIDATION ONLY (Product-Customer, single week).
 *  - Accuracy = 1 - SUM(weekly ABS error) / SUM(weekly Forecast), matching the
 *    SAP IBP convention  ABS(Forecast - Actual) / Forecast  so the alerts agree
 *    with the figures planners already see in IBP. Switch ACCURACY_DENOMINATOR
 *    to "actual" for the MAPE-style variant.
 *    Either way the absolute error is accumulated PER WEEK — never
 *    ABS(SUM(Forecast) - SUM(Actual)), which would let weekly errors cancel.
 *  - Bias     = SUM(Forecast - Actual) / SUM(Actual)   (+ = overforecast).
 *  - Bias deterioration = ABS(current Bias) - ABS(baseline Bias)  (away from 0).
 *  - Historical performance uses SNAPSHOT Key Figures, never current plan values.
 *  - Bulk read (one getValues), in-memory aggregation, bulk setValues writes.
 *
 * DIMENSION MAPPING
 *  The IBP CSV export can leave dimension headers blank. The script therefore:
 *    1. tries to identify each dimension from the header row by name,
 *    2. falls back to the configurable column-index constants below,
 *    3. verifies the resolved Week column actually parses,
 *    4. stops with an explicit error if a REQUIRED dimension cannot be resolved.
 *  Every resolution decision is written to the console — nothing is guessed
 *  silently.
 * ============================================================================
 */

/* ===========================================================================
 * 1. CONFIGURATION
 * =========================================================================*/

/** Worksheet holding the IBP extract. Never modified by this script. */
const SOURCE_SHEET_NAME: string = "DBAlerts";

/**
 * Fallback dimension column indexes, 0-based and ABSOLUTE (A = 0, B = 1, ...).
 * Used when the header cell is blank or does not match a known name.
 * Current IBP extraction: first seven columns are dimensions, column 8 (H, 7)
 * onwards are Key Figures.
 * Set a value to -1 to declare the dimension as "not present in the extract".
 */
const DIM_COL: { [key: string]: number } = {
    market: 0,            // A  Market / Sales Organisation
    product: 1,           // B  Product ID
    productDesc: 2,       // C  Product Description          (optional)
    productLocation: 3,   // D  Product-location / planning combination (optional)
    customer: 4,          // E  Customer
    customerHierarchy: 5, // F  Additional customer hierarchy dimension  (optional)
    week: 6               // G  Week / Time Period, e.g. "W01 26 (29/12)"
};

/** First column index that may contain Key Figures. */
const FIRST_KEY_FIGURE_COL: number = 7;

/** Try to identify dimensions from the header row before using DIM_COL. */
const USE_HEADER_NAME_DETECTION: boolean = true;

/** Header-name candidates per dimension (normalised, lower case, contains-match). */
const DIM_HEADER_CANDIDATES: { [key: string]: string[] } = {
    market: ["market", "sales organisation", "sales organization", "salesorg", "sales org"],
    product: ["product id", "productid", "product", "material", "item"],
    productDesc: ["product description", "product desc", "description", "material description"],
    productLocation: ["product location", "product-location", "planning combination", "location"],
    customer: ["customer id", "customer", "ship to", "ship-to", "sold to", "sold-to"],
    customerHierarchy: ["customer hierarchy", "hierarchy", "customer group", "channel"],
    week: ["week", "time period", "period", "calendar week", "cal week"]
};

/* ---- Alert 1: Final Demand Plan change vs W-1 --------------------------- */
const ALERT1_PCT_THRESHOLD: number = 0.10;     // |Variation %| must exceed this
const ALERT1_MIN_ABS_DIFF: number = 100;       // minimum absolute volume change
const ALERT1_FUTURE_HORIZON_WEEKS: number = 12; // W+1 .. W+12

/* ---- Alert 2: Accuracy / Bias deterioration ----------------------------- */
const ALERT2_ACCURACY_DETERIORATION_THRESHOLD: number = -0.05; // acc diff below this (-5 pp)
const ALERT2_BIAS_DETERIORATION_THRESHOLD: number = 0.05;      // |bias| growth above this (+5 pp)
const ALERT2_BIAS_MATERIALITY: number = 0.01;                  // dead-band for "No material change"
/** Baseline weeks required (W-5..W-2). 4 = require all four; lower to relax. */
const ALERT2_MIN_BASELINE_WEEKS: number = 4;
/**
 * Minimum W-1 actual volume for an Accuracy/Bias alert to be worth raising.
 *
 * Accuracy on a tiny base produces dramatic percentages from trivial absolute
 * misses — a combination selling 158 units in a week can show 0% accuracy from
 * an 88-unit error. Left at 0 every such combination alerts, which is what makes
 * Alerts 2 and 4 far noisier than 1 and 3. Set it to real materiality.
 */
const ALERT2_MIN_ACTUAL: number = 0;

/* ---- Alert 3: Forecast vs recent Sales ---------------------------------- */
const ALERT3_PCT_THRESHOLD: number = 0.15;     // |Variation %| must exceed this
const ALERT3_MIN_HISTORICAL_VOLUME: number = 100; // minimum Sales last 4 weeks
const ALERT3_MIN_FUTURE_VOLUME: number = 100;  // minimum forecast when there is no history

/* ---- Alert 4: Statistical Forecast vs Final Demand Plan ----------------- */
const ALERT4_ACCURACY_DIFF_THRESHOLD: number = 0.05; // Stat acc - FDP acc above this (+5 pp)
/** Completed weeks required in W-4..W-1. 4 = require all four. */
const ALERT4_MIN_HISTORY_WEEKS: number = 4;
/** Minimum Sales History over W-4..W-1. Same reasoning as ALERT2_MIN_ACTUAL. */
const ALERT4_MIN_ACTUAL: number = 0;

/* ---- Horizons ----------------------------------------------------------- */
const HISTORICAL_WEEKS: number = 5;  // keep W-1 .. W-5
const FUTURE_WEEKS: number = 12;     // keep W+1 .. W+12

/* ---- Behaviour ---------------------------------------------------------- */

/**
 * Denominator for Accuracy. The two conventions give materially different
 * numbers for the same data, so the choice matters:
 *
 *   "actual"   Accuracy = 1 - SUM(weekly |Forecast - Actual|) / SUM(Actual)
 *              What the written specification asks for. Standard MAPE-style:
 *              the error is measured against what actually happened, which
 *              keeps it comparable across periods and entities.
 *
 *   "forecast" Accuracy = 1 - SUM(weekly |Forecast - Actual|) / SUM(Forecast)
 *              What SAP IBP itself reports in the
 *              "5 Final Demand Plan (Shipments) Accuracy" Key Figure — verified
 *              against the BE30 / F19157 extract (Aug 2026), matching to three
 *              decimals across several customers and weeks. Use this to make the
 *              alerts agree with the figures planners already see in IBP.
 *
 * ACTIVE: "forecast" — chosen so the alerts agree with IBP. A planner opening
 * IBP after reading an alert must not find a different accuracy for the same
 * SKU; that costs more trust than the theoretically cleaner formula earns.
 *
 * Consequence of "forecast": a combination with zero forecast has no calculable
 * accuracy and is excluded, even when it had actual sales. Under "actual" the
 * reverse holds. Both are correct — the denominator cannot be zero either way.
 *
 * Bias is unaffected: it stays SUM(Forecast - Actual) / SUM(Actual) per spec.
 */
const ACCURACY_DENOMINATOR: string = "forecast";   // "actual" | "forecast"

/** Cap recalculated Accuracy into [0%, 100%] (IBP convention). */
const CAP_ACCURACY_0_100: boolean = true;
/** Skip the whole source row when a Key Figure cell cannot be parsed. */
const SKIP_ROWS_WITH_INVALID_NUMBERS: boolean = false;
/** Cross-check recalculated single-week KPIs against the source Accuracy/Bias. */
const ENABLE_SOURCE_KPI_VALIDATION: boolean = true;
const VALIDATION_TOLERANCE: number = 0.005;  // 0.5 pp
const VALIDATION_MAX_SAMPLES: number = 500;

/* ---- Output ------------------------------------------------------------- */
const OUT_SHEET_ALERT1: string = "Alert 1 FDP Change";
const OUT_SHEET_ALERT2: string = "Alert 2 Accuracy Bias";
const OUT_SHEET_ALERT3: string = "Alert 3 Forecast vs Sales";
const OUT_SHEET_ALERT4: string = "Alert 4 Stat vs FDP";

const OUT_TABLE_ALERT1: string = "tblAlert1FDPChange";
const OUT_TABLE_ALERT2: string = "tblAlert2AccuracyBias";
const OUT_TABLE_ALERT3: string = "tblAlert3ForecastVsSales";
const OUT_TABLE_ALERT4: string = "tblAlert4StatVsFDP";

const TABLE_STYLE: string = "TableStyleMedium2";
const FMT_PERCENT: string = "0.0%";
const FMT_VOLUME: string = "#,##0";
const FMT_DATE: string = "dd/mm/yyyy";
const MAX_DESCRIPTION_COLUMN_WIDTH: number = 240; // points

/**
 * The payload limit applies to WRITES as well as reads. Alert 1 is the largest
 * output (one row per combination per future week), so it is the first to hit it
 * — and because it is published first, a failure there leaves the other three
 * worksheets showing stale results from the previous run.
 *
 * Output is therefore written in blocks of roughly this many cells, halving and
 * retrying if a block is still too large.
 */
const WRITE_CHUNK_CELLS: number = 100000;
const MIN_WRITE_CHUNK_ROWS: number = 50;

/**
 * Office Scripts caps the payload of a single API call, so the source CANNOT be
 * read with one getValues() — a large IBP extract fails with
 * "Range getValues: The response payload size has exceeded the limit".
 * See https://learn.microsoft.com/office/dev/scripts/testing/platform-limits
 *
 * The sheet is therefore read in row blocks of roughly this many cells. If a
 * block still exceeds the limit the reader halves it and retries automatically,
 * so this value only needs to be in the right ballpark. Lower it if you hit the
 * limit repeatedly; raise it to reduce the number of calls on a small extract.
 */
const READ_CHUNK_CELLS: number = 100000;
/** Smallest block the reader will fall back to before giving up. */
const MIN_READ_CHUNK_ROWS: number = 50;
/** Rows read up front to find the header and sample the Week column. */
const PROBE_ROWS: number = 250;

/* ---- Exact Key Figure names (matched after normalisation) --------------- */
const KF_SALES_HISTORY: string = "Sales History (Shipments)";
const KF_FDP_CURRENT: string = "5 Final Demand Plan (Shipments)";
const KF_FDP_PREVIOUS: string = "5 Final Demand Plan (Shipments) W-1";
const KF_FDP_SNAPSHOT: string = "5 Final Demand Plan (Shipments) Snapshot";
const KF_FDP_ACCURACY: string = "5 Final Demand Plan (Shipments) Accuracy";
const KF_FDP_ERROR: string = "5 Final Demand Plan (Shipments) Error";
const KF_FDP_BIAS: string = "5 Final Demand Plan (Shipments) Bias";
const KF_STAT_SNAPSHOT: string = "0.2 Full Stat. Forecast (Snapshot)";
const KF_STAT_ERROR: string = "0.2 Full Stat. Forecast Error";
const KF_STAT_ACCURACY: string = "0.2 Full Stat. Forecast Accuracy";

/** Key Figures without which the alerts cannot run. */
const REQUIRED_KEY_FIGURES: string[] = [
    KF_SALES_HISTORY,
    KF_FDP_CURRENT,
    KF_FDP_PREVIOUS,
    KF_FDP_SNAPSHOT,
    KF_STAT_SNAPSHOT
];

/** Key Figures used for validation only — absence is not fatal. */
const OPTIONAL_KEY_FIGURES: string[] = [
    KF_FDP_ACCURACY,
    KF_FDP_ERROR,
    KF_FDP_BIAS,
    KF_STAT_ERROR,
    KF_STAT_ACCURACY
];

const AGG_PRODUCT_CUSTOMER: string = "Product-Customer";
const AGG_PRODUCT: string = "Product";
const CUSTOMER_ALL: string = "ALL";

/**
 * Written into the Aggregation Level cell of the single placeholder row an alert
 * table carries when nothing met its criteria. Consumers skip it: it has no
 * Market and no Product.
 */
const NO_ALERTS_MARKER: string = "No alerts this week";

const MS_PER_DAY: number = 86400000;
const MS_PER_WEEK: number = 604800000;

/* ===========================================================================
 * 2. INTERFACES
 * =========================================================================*/

interface DimensionMap {
    market: number;
    product: number;
    productDesc: number;
    productLocation: number;
    customer: number;
    customerHierarchy: number;
    week: number;
}

interface KeyFigureMap {
    salesHistory: number;
    fdpCurrent: number;
    fdpPrevious: number;
    fdpSnapshot: number;
    fdpAccuracy: number;
    fdpError: number;
    fdpBias: number;
    statSnapshot: number;
    statError: number;
    statAccuracy: number;
}

/** Parsed week label. */
interface ParsedWeek {
    valid: boolean;
    mondayUtc: number;   // ms, UTC midnight of the Monday starting the week
    label: string;
}

/** Volumes aggregated for one entity and one week bucket. */
interface WeekAggregate {
    sales: number;
    fdpCurrent: number;
    fdpPrevious: number;
    fdpSnapshot: number;
    statSnapshot: number;
    sourceRows: number;
    srcFdpAccuracy: number | null;
    srcFdpBias: number | null;
    srcStatAccuracy: number | null;
}

/** One aggregation entity (Product-Customer or Product) with its week buckets. */
interface Entity {
    market: string;
    product: string;
    description: string;
    customer: string;
    weeks: Map<number, WeekAggregate>;
}

interface PeriodTotals {
    actual: number;
    forecast: number;
    absError: number;
    signedError: number;
    weeks: number;
}

interface AlertOutput {
    headers: string[];
    rows: (string | number)[][];
    percentColumns: number[];
    volumeColumns: number[];
    dateColumns: number[];
    descriptionColumn: number;
    productCustomerCount: number;
    productCount: number;
}

interface ExecutionSummary {
    sourceRowsRead: number;
    sourceRowsIgnoredInvalidDate: number;
    sourceRowsIgnoredInvalidNumber: number;
    sourceRowsWithInvalidNumericCells: number;
    sourceRowsOutsideAnalysedHorizon: number;
    currentWeekMonday: string;
    alert1ProductCustomerCount: number;
    alert1ProductCount: number;
    alert2ProductCustomerCount: number;
    alert2ProductCount: number;
    alert3ProductCustomerCount: number;
    alert3ProductCount: number;
    alert4ProductCustomerCount: number;
    alert4ProductCount: number;
    excludedInsufficientHistoricalWeeks: number;
    excludedZeroActual: number;
    excludedBelowMinimumVolume: number;
    sourceKpiValidationChecks: number;
    sourceKpiValidationMismatches: number;
    outputWorksheets: string[];
}

/** Mutable counters shared across the run. */
interface RunCounters {
    excludedInsufficientWeeks: number;
    excludedZeroActual: number;
    excludedBelowMinimumVolume: number;
    validationChecks: number;
    validationMismatches: number;
}

/* ===========================================================================
 * 3. MAIN
 * =========================================================================*/

function main(workbook: ExcelScript.Workbook): ExecutionSummary {

    const sourceSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet(SOURCE_SHEET_NAME);
    if (!sourceSheet) {
        throw new Error(
            "Source worksheet '" + SOURCE_SHEET_NAME + "' was not found in this workbook. " +
            "Update the SOURCE_SHEET_NAME constant or add the worksheet."
        );
    }

    /* ---- 3.1 Probe: header + a sample of rows, in one small read -------- */
    const bounds: SourceBounds = getSourceBounds(sourceSheet);
    const probe: (string | number | boolean)[][] = readBlock(
        sourceSheet, 0, Math.min(bounds.rowCount, PROBE_ROWS), bounds.columnCount
    );

    const headerRowIndex: number = findHeaderRow(probe);
    if (headerRowIndex < 0) {
        throw new Error("Could not locate a header row in '" + SOURCE_SHEET_NAME + "'. The sheet appears to be empty.");
    }
    const headerRow: (string | number | boolean)[] = probe[headerRowIndex];

    /* ---- 3.2 Resolve columns ------------------------------------------- */
    const headerMap: Map<string, number> = buildHeaderMap(headerRow);
    const dims: DimensionMap = resolveDimensionColumns(headerRow, probe, headerRowIndex);
    const kfs: KeyFigureMap = validateRequiredColumns(headerMap, headerRow);

    // Only columns up to the last one actually used need to be read. On a wide
    // extract this alone can cut the payload substantially.
    const readColumns: number = Math.min(
        bounds.columnCount,
        lastNeededColumn(dims, kfs) + 1
    );
    const blockRows: number = blockRowsFor(readColumns);
    console.log(
        "Reading " + (bounds.rowCount - headerRowIndex - 1) + " data row(s) x " +
        readColumns + " column(s) in blocks of " + blockRows + " rows."
    );

    /* ---- 3.3 Reference dates -------------------------------------------- */
    const currentMonday: number = getMondayOfCurrentWeek();
    console.log("Current week Monday: " + formatDate(currentMonday));

    /* ---- 3.4 Aggregate ------------------------------------------------- */
    const counters: RunCounters = {
        excludedInsufficientWeeks: 0,
        excludedZeroActual: 0,
        excludedBelowMinimumVolume: 0,
        validationChecks: 0,
        validationMismatches: 0
    };

    const productCustomerMap: Map<string, Entity> = new Map<string, Entity>();
    const productMap: Map<string, Entity> = new Map<string, Entity>();
    const weekLabels: Map<number, string> = new Map<number, string>();
    const weekStarts: Map<number, number> = new Map<number, number>();

    let rowsRead: number = 0;
    let rowsInvalidDate: number = 0;
    let rowsInvalidNumberSkipped: number = 0;
    let rowsWithInvalidCells: number = 0;
    let rowsOutsideHorizon: number = 0;

    // Totals per Key Figure over the rows actually analysed. A Key Figure that
    // is empty in the extract makes whole alerts silently impossible — Accuracy
    // and Bias need actuals, so a blank Sales History excludes every entity and
    // the run reports "no alerts" with nothing to indicate why.
    let totSales: number = 0;
    let totFdpCur: number = 0;
    let totFdpPrev: number = 0;
    let totFdpSnap: number = 0;
    let totStatSnap: number = 0;

    // Where the volumes actually sit, week by week. A Key Figure can be well
    // populated in the extract yet contribute nothing because its values fall
    // outside the analysed window — which looks identical to an empty column.
    const salesByOffset: Map<number, number> = new Map<number, number>();
    const fdpByOffset: Map<number, number> = new Map<number, number>();
    let minOffsetSeen: number = 9999;
    let maxOffsetSeen: number = -9999;

    const minOffset: number = -HISTORICAL_WEEKS;
    const maxOffset: number = FUTURE_WEEKS;

    // Stream the sheet a block at a time. Each block is parsed straight into the
    // entity maps and then dropped, so peak memory stays flat regardless of how
    // many rows the extract has.
    let nextRow: number = headerRowIndex + 1;
    let blocksRead: number = 0;

    while (nextRow < bounds.rowCount) {
        const wanted: number = Math.min(blockRows, bounds.rowCount - nextRow);
        const block: (string | number | boolean)[][] =
            readBlock(sourceSheet, nextRow, wanted, readColumns);
        if (block.length === 0) {
            break; // defensive: never spin on a zero-length read
        }
        nextRow += block.length;
        blocksRead++;

        for (let i: number = 0; i < block.length; i++) {
            const row: (string | number | boolean)[] = block[i];
            if (isEmptyRow(row)) {
                continue;
            }
            rowsRead++;

            const week: ParsedWeek = parseSourceWeek(row[dims.week]);
            if (!week.valid) {
                rowsInvalidDate++;
                continue;
            }

            const offset: number = calculateWeekOffset(week.mondayUtc, currentMonday);
            if (offset < minOffsetSeen) {
                minOffsetSeen = offset;
            }
            if (offset > maxOffsetSeen) {
                maxOffsetSeen = offset;
            }
            if (offset === 0 || offset < minOffset || offset > maxOffset) {
                // Current week is excluded from every calculation, by specification.
                rowsOutsideHorizon++;
                continue;
            }

            const sales: number = toNumber(row[kfs.salesHistory]);
            const fdpCur: number = toNumber(row[kfs.fdpCurrent]);
            const fdpPrev: number = toNumber(row[kfs.fdpPrevious]);
            const fdpSnap: number = toNumber(row[kfs.fdpSnapshot]);
            const statSnap: number = toNumber(row[kfs.statSnapshot]);

            const invalidCell: boolean =
                isUnparseable(row[kfs.salesHistory]) || isUnparseable(row[kfs.fdpCurrent]) ||
                isUnparseable(row[kfs.fdpPrevious]) || isUnparseable(row[kfs.fdpSnapshot]) ||
                isUnparseable(row[kfs.statSnapshot]);

            if (invalidCell) {
                rowsWithInvalidCells++;
                if (SKIP_ROWS_WITH_INVALID_NUMBERS) {
                    rowsInvalidNumberSkipped++;
                    continue;
                }
            }

            salesByOffset.set(offset, (salesByOffset.get(offset) || 0) + sales);
            fdpByOffset.set(offset, (fdpByOffset.get(offset) || 0) + fdpCur);

            totSales += sales;
            totFdpCur += fdpCur;
            totFdpPrev += fdpPrev;
            totFdpSnap += fdpSnap;
            totStatSnap += statSnap;

            const market: string = readText(row[dims.market]);
            const product: string = readText(row[dims.product]);
            const description: string = dims.productDesc >= 0 ? readText(row[dims.productDesc]) : "";
            const customer: string = readText(row[dims.customer]);

            if (!weekLabels.has(offset)) {
                weekLabels.set(offset, week.label);
                weekStarts.set(offset, week.mondayUtc);
            }

            const srcFdpAcc: number | null = kfs.fdpAccuracy >= 0 ? toRatioOrNull(row[kfs.fdpAccuracy]) : null;
            const srcFdpBias: number | null = kfs.fdpBias >= 0 ? toRatioOrNull(row[kfs.fdpBias]) : null;
            const srcStatAcc: number | null = kfs.statAccuracy >= 0 ? toRatioOrNull(row[kfs.statAccuracy]) : null;

            addToEntity(
                productCustomerMap, createCompositeKey([market, product, customer]),
                market, product, description, customer,
                offset, sales, fdpCur, fdpPrev, fdpSnap, statSnap,
                srcFdpAcc, srcFdpBias, srcStatAcc
            );

            addToEntity(
                productMap, createCompositeKey([market, product]),
                market, product, description, CUSTOMER_ALL,
                offset, sales, fdpCur, fdpPrev, fdpSnap, statSnap,
                null, null, null
            );
        }
    }

    console.log(
        "Source rows read: " + rowsRead +
        " (in " + blocksRead + " block(s))" +
        " | invalid week: " + rowsInvalidDate +
        " | outside horizon / current week: " + rowsOutsideHorizon +
        " | Product-Customer entities: " + productCustomerMap.size +
        " | Product entities: " + productMap.size
    );

    console.log(
        "Key Figure totals over analysed rows: " +
        "Sales History=" + Math.round(totSales) +
        " | Current FDP=" + Math.round(totFdpCur) +
        " | FDP W-1=" + Math.round(totFdpPrev) +
        " | FDP Snapshot=" + Math.round(totFdpSnap) +
        " | Stat Snapshot=" + Math.round(totStatSnap)
    );

    // Week-by-week placement. If the totals above are zero but the extract
    // clearly holds values, they are in weeks this run does not analyse — and
    // the offsets-present range says how far the extract reaches.
    const perWeek: string[] = [];
    for (let o: number = minOffset; o <= maxOffset; o++) {
        if (o === 0) {
            continue;
        }
        perWeek.push(
            "W" + (o > 0 ? "+" : "") + o + " sales=" + Math.round(salesByOffset.get(o) || 0) +
            " fdp=" + Math.round(fdpByOffset.get(o) || 0)
        );
    }
    console.log("Volumes by analysed week: " + perWeek.join(" | "));
    console.log(
        "Week offsets present in the extract: " +
        (minOffsetSeen === 9999 ? "none" : minOffsetSeen + " .. " + maxOffsetSeen) +
        " (analysed: " + minOffset + " .. " + maxOffset + ", current week excluded)"
    );

    const emptyKfs: string[] = [];
    if (totSales === 0) { emptyKfs.push(KF_SALES_HISTORY); }
    if (totFdpCur === 0) { emptyKfs.push(KF_FDP_CURRENT); }
    if (totFdpPrev === 0) { emptyKfs.push(KF_FDP_PREVIOUS); }
    if (totFdpSnap === 0) { emptyKfs.push(KF_FDP_SNAPSHOT); }
    if (totStatSnap === 0) { emptyKfs.push(KF_STAT_SNAPSHOT); }
    if (emptyKfs.length > 0) {
        console.log(
            "WARNING: these Key Figures are entirely zero/blank in the analysed rows: " +
            emptyKfs.join(", ") + ".\n" +
            "  Alerts depending on them CANNOT produce results:\n" +
            "    Sales History  -> Alerts 2, 3 and 4 (no actuals, so no Accuracy or Bias)\n" +
            "    FDP Snapshot   -> Alerts 2 and 4\n" +
            "    Stat Snapshot  -> Alert 4\n" +
            "    Current FDP    -> Alerts 1 and 3\n" +
            "    FDP W-1        -> Alert 1\n" +
            "  Check the extract really contains this Key Figure, and that its column\n" +
            "  header matches the KF_* constant exactly."
        );
    }

    if (ENABLE_SOURCE_KPI_VALIDATION) {
        validateAgainstSourceKpis(productCustomerMap, counters);
    }

    /* ---- 3.5 Run the alerts -------------------------------------------- */
    const outputSheets: string[] = [];

    const alert1: AlertOutput = runAlert1(productCustomerMap, productMap, weekLabels, weekStarts);
    publishAlert(workbook, OUT_SHEET_ALERT1, OUT_TABLE_ALERT1, alert1);
    outputSheets.push(OUT_SHEET_ALERT1);

    const alert2: AlertOutput = runAlert2(productCustomerMap, productMap, counters);
    publishAlert(workbook, OUT_SHEET_ALERT2, OUT_TABLE_ALERT2, alert2);
    outputSheets.push(OUT_SHEET_ALERT2);

    const alert3: AlertOutput = runAlert3(productCustomerMap, productMap, currentMonday);
    publishAlert(workbook, OUT_SHEET_ALERT3, OUT_TABLE_ALERT3, alert3);
    outputSheets.push(OUT_SHEET_ALERT3);

    const alert4: AlertOutput = runAlert4(productCustomerMap, productMap, counters);
    publishAlert(workbook, OUT_SHEET_ALERT4, OUT_TABLE_ALERT4, alert4);
    outputSheets.push(OUT_SHEET_ALERT4);

    verifyOutputTables(workbook);

    /* ---- 3.6 Summary ---------------------------------------------------- */
    const summary: ExecutionSummary = {
        sourceRowsRead: rowsRead,
        sourceRowsIgnoredInvalidDate: rowsInvalidDate,
        sourceRowsIgnoredInvalidNumber: rowsInvalidNumberSkipped,
        sourceRowsWithInvalidNumericCells: rowsWithInvalidCells,
        sourceRowsOutsideAnalysedHorizon: rowsOutsideHorizon,
        currentWeekMonday: formatDate(currentMonday),
        alert1ProductCustomerCount: alert1.productCustomerCount,
        alert1ProductCount: alert1.productCount,
        alert2ProductCustomerCount: alert2.productCustomerCount,
        alert2ProductCount: alert2.productCount,
        alert3ProductCustomerCount: alert3.productCustomerCount,
        alert3ProductCount: alert3.productCount,
        alert4ProductCustomerCount: alert4.productCustomerCount,
        alert4ProductCount: alert4.productCount,
        excludedInsufficientHistoricalWeeks: counters.excludedInsufficientWeeks,
        excludedZeroActual: counters.excludedZeroActual,
        excludedBelowMinimumVolume: counters.excludedBelowMinimumVolume,
        sourceKpiValidationChecks: counters.validationChecks,
        sourceKpiValidationMismatches: counters.validationMismatches,
        outputWorksheets: outputSheets
    };

    returnExecutionSummary(summary);
    return summary;
}

/* ===========================================================================
 * 4. SOURCE READING AND COLUMN RESOLUTION
 * =========================================================================*/

/** Extent of the used range, anchored at A1 so column indexes are absolute. */
interface SourceBounds {
    rowCount: number;   // rows from row 1 down to the last used row
    columnCount: number; // columns from A across to the last used column
}

function getSourceBounds(sheet: ExcelScript.Worksheet): SourceBounds {
    const used: ExcelScript.Range | undefined = sheet.getUsedRange();
    if (!used) {
        throw new Error("Worksheet '" + SOURCE_SHEET_NAME + "' is empty — nothing to analyse.");
    }
    // Anchor at A1 so the DIM_COL constants refer to real worksheet columns
    // (A = 0, B = 1, ...) even when the used range starts further in.
    return {
        rowCount: used.getRowIndex() + used.getRowCount(),
        columnCount: used.getColumnIndex() + used.getColumnCount()
    };
}

/**
 * Reads a block of rows, halving the block and retrying if the response exceeds
 * the Office Scripts payload limit. Returns the rows read; the caller advances
 * by the returned length rather than by the requested size.
 */
function readBlock(
    sheet: ExcelScript.Worksheet,
    startRow: number,
    rowsWanted: number,
    columnCount: number
): (string | number | boolean)[][] {

    let rows: number = rowsWanted;
    for (;;) {
        try {
            return sheet.getRangeByIndexes(startRow, 0, rows, columnCount).getValues();
        } catch (e) {
            if (rows <= MIN_READ_CHUNK_ROWS) {
                throw new Error(
                    "Could not read rows " + (startRow + 1) + "-" + (startRow + rows) +
                    " of '" + SOURCE_SHEET_NAME + "' even at the minimum block size (" +
                    MIN_READ_CHUNK_ROWS + " rows). The sheet may have extremely wide " +
                    "content. Reduce the number of columns in the extract, or lower " +
                    "MIN_READ_CHUNK_ROWS. Underlying error: " + String(e)
                );
            }
            rows = Math.max(MIN_READ_CHUNK_ROWS, Math.floor(rows / 2));
            console.log(
                "  read block too large; retrying with " + rows + " rows from row " + (startRow + 1)
            );
        }
    }
}

/** Rows per read, sized so a block stays under the payload limit. */
function blockRowsFor(columnCount: number): number {
    return Math.max(MIN_READ_CHUNK_ROWS, Math.floor(READ_CHUNK_CELLS / Math.max(1, columnCount)));
}

/**
 * Rightmost column the script actually reads. Columns beyond this are ignored,
 * which shrinks every block on an extract carrying unused Key Figures.
 */
function lastNeededColumn(dims: DimensionMap, kfs: KeyFigureMap): number {
    const indexes: number[] = [
        dims.market, dims.product, dims.productDesc, dims.productLocation,
        dims.customer, dims.customerHierarchy, dims.week,
        kfs.salesHistory, kfs.fdpCurrent, kfs.fdpPrevious, kfs.fdpSnapshot,
        kfs.fdpAccuracy, kfs.fdpError, kfs.fdpBias,
        kfs.statSnapshot, kfs.statError, kfs.statAccuracy
    ];
    let max: number = 0;
    for (let i: number = 0; i < indexes.length; i++) {
        if (indexes[i] > max) {
            max = indexes[i];
        }
    }
    return max;
}

/** First row containing at least three non-empty cells is treated as the header. */
function findHeaderRow(grid: (string | number | boolean)[][]): number {
    const limit: number = Math.min(grid.length, 50);
    for (let r: number = 0; r < limit; r++) {
        let filled: number = 0;
        for (let c: number = 0; c < grid[r].length; c++) {
            if (readText(grid[r][c]) !== "") {
                filled++;
            }
        }
        if (filled >= 3) {
            return r;
        }
    }
    return grid.length > 0 ? 0 : -1;
}

/**
 * Trims, converts non-breaking spaces and collapses repeated internal spaces.
 * Key Figures such as "5  Final Demand Plan (Shipments)" (double space) are
 * matched safely, while genuinely different Key Figures stay distinct.
 */
function normaliseHeader(value: string | number | boolean): string {
    let text: string = readText(value);
    text = text.replace(/ /g, " ");
    text = text.replace(/\s+/g, " ");
    return text.trim();
}

/** Normalised header -> column index. First occurrence wins. */
function buildHeaderMap(headerRow: (string | number | boolean)[]): Map<string, number> {
    const map: Map<string, number> = new Map<string, number>();
    for (let c: number = 0; c < headerRow.length; c++) {
        const key: string = normaliseHeader(headerRow[c]).toLowerCase();
        if (key !== "" && !map.has(key)) {
            map.set(key, c);
        }
    }
    return map;
}

/** Locates a Key Figure column by exact normalised name. -1 when absent. */
function findKeyFigureColumn(headerMap: Map<string, number>, name: string): number {
    const key: string = normaliseHeader(name).toLowerCase();
    const found: number | undefined = headerMap.get(key);
    return found === undefined ? -1 : found;
}

/** Validates every required Key Figure and returns their column indexes. */
function validateRequiredColumns(
    headerMap: Map<string, number>,
    headerRow: (string | number | boolean)[]
): KeyFigureMap {

    const missing: string[] = [];
    for (let i: number = 0; i < REQUIRED_KEY_FIGURES.length; i++) {
        if (findKeyFigureColumn(headerMap, REQUIRED_KEY_FIGURES[i]) < 0) {
            missing.push(REQUIRED_KEY_FIGURES[i]);
        }
    }

    if (missing.length > 0) {
        const available: string[] = [];
        for (let c: number = FIRST_KEY_FIGURE_COL; c < headerRow.length; c++) {
            const h: string = normaliseHeader(headerRow[c]);
            if (h !== "") {
                available.push(h);
            }
        }
        throw new Error(
            "Missing required Key Figure column(s) in '" + SOURCE_SHEET_NAME + "':\n  - " +
            missing.join("\n  - ") +
            "\nKey Figure headers found in the extract:\n  - " +
            (available.length > 0 ? available.join("\n  - ") : "(none)") +
            "\nCheck the IBP extract, or correct the Key Figure name constants at the top of the script."
        );
    }

    const map: KeyFigureMap = {
        salesHistory: findKeyFigureColumn(headerMap, KF_SALES_HISTORY),
        fdpCurrent: findKeyFigureColumn(headerMap, KF_FDP_CURRENT),
        fdpPrevious: findKeyFigureColumn(headerMap, KF_FDP_PREVIOUS),
        fdpSnapshot: findKeyFigureColumn(headerMap, KF_FDP_SNAPSHOT),
        fdpAccuracy: findKeyFigureColumn(headerMap, KF_FDP_ACCURACY),
        fdpError: findKeyFigureColumn(headerMap, KF_FDP_ERROR),
        fdpBias: findKeyFigureColumn(headerMap, KF_FDP_BIAS),
        statSnapshot: findKeyFigureColumn(headerMap, KF_STAT_SNAPSHOT),
        statError: findKeyFigureColumn(headerMap, KF_STAT_ERROR),
        statAccuracy: findKeyFigureColumn(headerMap, KF_STAT_ACCURACY)
    };

    // Which column each Key Figure resolved to. Without this, a mis-mapped or
    // empty Key Figure is invisible — the run simply reports no alerts.
    console.log(
        "Key Figures -> " +
        "SalesHistory=" + kfColumn(map.salesHistory) +
        " FDP=" + kfColumn(map.fdpCurrent) +
        " FDP W-1=" + kfColumn(map.fdpPrevious) +
        " FDP Snapshot=" + kfColumn(map.fdpSnapshot) +
        " Stat Snapshot=" + kfColumn(map.statSnapshot) +
        " | validation: FDPAcc=" + kfColumn(map.fdpAccuracy) +
        " FDPBias=" + kfColumn(map.fdpBias) +
        " StatAcc=" + kfColumn(map.statAccuracy)
    );

    const absentOptional: string[] = [];
    for (let i: number = 0; i < OPTIONAL_KEY_FIGURES.length; i++) {
        if (findKeyFigureColumn(headerMap, OPTIONAL_KEY_FIGURES[i]) < 0) {
            absentOptional.push(OPTIONAL_KEY_FIGURES[i]);
        }
    }
    if (absentOptional.length > 0) {
        console.log("Optional validation Key Figures not present (alerts unaffected): " + absentOptional.join(", "));
    }

    return map;
}

/**
 * Resolves the dimension columns: header names first, configured indexes as a
 * documented fallback, then a data-driven check of the Week column.
 */
function resolveDimensionColumns(
    headerRow: (string | number | boolean)[],
    grid: (string | number | boolean)[][],
    headerRowIndex: number
): DimensionMap {

    const names: string[] = [
        "market", "product", "productDesc", "productLocation",
        "customer", "customerHierarchy", "week"
    ];
    const resolved: { [key: string]: number } = {};

    for (let i: number = 0; i < names.length; i++) {
        const dim: string = names[i];
        const configured: number = DIM_COL[dim];
        let index: number = -1;
        let how: string = "";

        if (USE_HEADER_NAME_DETECTION) {
            index = matchDimensionByName(headerRow, DIM_HEADER_CANDIDATES[dim]);
            if (index >= 0) {
                how = "header name '" + normaliseHeader(headerRow[index]) + "'";
            }
        }

        if (index < 0 && configured >= 0 && configured < headerRow.length) {
            index = configured;
            how = "configured column index (header " +
                (normaliseHeader(headerRow[configured]) === "" ? "blank" : "'" + normaliseHeader(headerRow[configured]) + "'") + ")";
        }

        resolved[dim] = index;
        console.log(
            "Dimension '" + dim + "' -> " +
            (index < 0 ? "NOT RESOLVED" : "column " + columnLetter(index) + " (" + index + ") via " + how)
        );
    }

    const dims: DimensionMap = {
        market: resolved["market"],
        product: resolved["product"],
        productDesc: resolved["productDesc"],
        productLocation: resolved["productLocation"],
        customer: resolved["customer"],
        customerHierarchy: resolved["customerHierarchy"],
        week: resolved["week"]
    };

    // The Week column must actually contain parseable week labels.
    if (dims.week < 0 || !weekColumnParses(grid, headerRowIndex, dims.week)) {
        const detected: number = detectWeekColumn(grid, headerRowIndex, headerRow.length);
        if (detected >= 0) {
            console.log(
                "Week column re-resolved by content inspection to column " +
                columnLetter(detected) + " (" + detected + ")."
            );
            dims.week = detected;
        } else {
            throw new Error(
                "The Week / Time Period dimension could not be identified in '" + SOURCE_SHEET_NAME + "'.\n" +
                "Expected values such as \"W01 26 (29/12)\" containing the week start date in parentheses.\n" +
                "Set DIM_COL.week to the correct 0-based column index (A = 0) and run again."
            );
        }
    }

    const requiredDims: string[] = ["market", "product", "customer"];
    const requiredValues: number[] = [dims.market, dims.product, dims.customer];
    const unresolved: string[] = [];
    for (let i: number = 0; i < requiredDims.length; i++) {
        if (requiredValues[i] < 0) {
            unresolved.push(requiredDims[i]);
        }
    }
    if (unresolved.length > 0) {
        throw new Error(
            "Required dimension(s) could not be identified in '" + SOURCE_SHEET_NAME + "': " +
            unresolved.join(", ") + ".\n" +
            "The IBP export left the header blank and no fallback column index is configured.\n" +
            "Set the correct 0-based column indexes in the DIM_COL constant (A = 0, B = 1, ...) and run again."
        );
    }

    if (dims.productDesc < 0) {
        console.log("Product Description not identified — the output column will be left blank.");
    }

    // Guard against two dimensions resolving to the same column.
    const seen: Map<number, string> = new Map<number, string>();
    const check: string[] = ["market", "product", "customer", "week"];
    const checkValues: number[] = [dims.market, dims.product, dims.customer, dims.week];
    for (let i: number = 0; i < check.length; i++) {
        const idx: number = checkValues[i];
        if (seen.has(idx)) {
            throw new Error(
                "Dimensions '" + seen.get(idx) + "' and '" + check[i] + "' both resolved to column " +
                columnLetter(idx) + ". Correct the DIM_COL constant."
            );
        }
        seen.set(idx, check[i]);
    }

    return dims;
}

/**
 * Contains-match of a header against the candidate names for a dimension.
 * The dimension zone (columns before FIRST_KEY_FIGURE_COL) is searched first so
 * that a Key Figure header never wins over a genuine dimension header.
 */
function matchDimensionByName(headerRow: (string | number | boolean)[], candidates: string[]): number {
    const zoneEnd: number = Math.min(FIRST_KEY_FIGURE_COL, headerRow.length);
    const found: number = matchInRange(headerRow, candidates, 0, zoneEnd);
    if (found >= 0) {
        return found;
    }
    return matchInRange(headerRow, candidates, 0, headerRow.length);
}

/** Exact match first, then contains-match, so "Product ID" beats "Product Description". */
function matchInRange(
    headerRow: (string | number | boolean)[],
    candidates: string[],
    from: number,
    to: number
): number {
    for (let ci: number = 0; ci < candidates.length; ci++) {
        for (let c: number = from; c < to; c++) {
            if (normaliseHeader(headerRow[c]).toLowerCase() === candidates[ci]) {
                return c;
            }
        }
    }
    for (let ci: number = 0; ci < candidates.length; ci++) {
        for (let c: number = from; c < to; c++) {
            const h: string = normaliseHeader(headerRow[c]).toLowerCase();
            if (h !== "" && h.indexOf(candidates[ci]) >= 0) {
                return c;
            }
        }
    }
    return -1;
}

/** True when most sampled values in the column parse as week labels. */
function weekColumnParses(grid: (string | number | boolean)[][], headerRowIndex: number, col: number): boolean {
    let checked: number = 0;
    let ok: number = 0;
    for (let r: number = headerRowIndex + 1; r < grid.length && checked < 200; r++) {
        const raw: string = readText(grid[r][col]);
        if (raw === "") {
            continue;
        }
        checked++;
        if (parseSourceWeek(grid[r][col]).valid) {
            ok++;
        }
    }
    return checked > 0 && ok / checked >= 0.5;
}

/** Scans all columns for one whose content parses as week labels. */
function detectWeekColumn(grid: (string | number | boolean)[][], headerRowIndex: number, columnCount: number): number {
    for (let c: number = 0; c < columnCount; c++) {
        if (weekColumnParses(grid, headerRowIndex, c)) {
            return c;
        }
    }
    return -1;
}

/* ===========================================================================
 * 5. DATE AND WEEK LOGIC
 * =========================================================================*/

/** UTC midnight of the Monday of the week in which the script runs. */
function getMondayOfCurrentWeek(): number {
    const now: Date = new Date();
    const todayUtc: number = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
    return mondayOf(todayUtc);
}

/** UTC midnight of the Monday of the week containing the given instant. */
function mondayOf(utcMs: number): number {
    const dow: number = new Date(utcMs).getUTCDay(); // 0 = Sunday
    const shift: number = (dow + 6) % 7;             // Monday = 0
    return utcMs - shift * MS_PER_DAY;
}

/** Whole weeks between a week start and the current week Monday. 0 = current week. */
function calculateWeekOffset(weekMondayUtc: number, currentMondayUtc: number): number {
    return Math.round((weekMondayUtc - currentMondayUtc) / MS_PER_WEEK);
}

/**
 * Parses IBP week labels such as "W01 26 (29/12)".
 * The date inside the parentheses is authoritative for the week start; the week
 * number alone is unreliable because ISO weeks cross calendar years.
 * Supported: (dd/mm), (dd/mm/yy), (dd/mm/yyyy), and "." or "-" separators.
 */
function parseSourceWeek(raw: string | number | boolean): ParsedWeek {
    const invalid: ParsedWeek = { valid: false, mondayUtc: 0, label: "" };

    // A real Excel date serial in the week column is also accepted.
    if (typeof raw === "number" && isFinite(raw) && raw > 20000 && raw < 80000) {
        const utc: number = excelSerialToUtc(raw);
        return { valid: true, mondayUtc: mondayOf(utc), label: formatDate(mondayOf(utc)) };
    }

    const label: string = normaliseHeader(raw);
    if (label === "") {
        return invalid;
    }

    const paren: RegExpMatchArray | null = label.match(/\(([^)]*)\)/);
    if (!paren) {
        return invalid;
    }

    const dm: RegExpMatchArray | null = paren[1].match(/(\d{1,2})\s*[\/\.\-]\s*(\d{1,2})(?:\s*[\/\.\-]\s*(\d{2,4}))?/);
    if (!dm) {
        return invalid;
    }

    const day: number = parseInt(dm[1], 10);
    const month: number = parseInt(dm[2], 10);
    if (day < 1 || day > 31 || month < 1 || month > 12) {
        return invalid;
    }

    let year: number = -1;

    if (dm[3] !== undefined && dm[3] !== null && dm[3] !== "") {
        year = parseInt(dm[3], 10);
        if (year < 100) {
            year = 2000 + year;
        }
    } else {
        // Derive the year from the "Wnn yy" prefix, handling the year crossing.
        const prefix: string = label.substring(0, label.indexOf("("));
        const wm: RegExpMatchArray | null = prefix.match(/W\s*(\d{1,2})\D+(\d{2,4})/i);
        if (wm) {
            const weekNo: number = parseInt(wm[1], 10);
            let weekYear: number = parseInt(wm[2], 10);
            if (weekYear < 100) {
                weekYear = 2000 + weekYear;
            }
            year = weekYear;
            if (month === 12 && weekNo <= 3) {
                year = weekYear - 1;   // e.g. W01 26 (29/12) -> 29 Dec 2025
            } else if (month === 1 && weekNo >= 50) {
                year = weekYear + 1;   // e.g. W52 25 (01/01) -> 01 Jan 2026
            }
        } else {
            // No week/year prefix: choose the calendar year closest to today.
            const today: Date = new Date();
            const base: number = today.getFullYear();
            let best: number = base;
            let bestDist: number = Number.MAX_VALUE;
            const todayUtc: number = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate());
            for (let y: number = base - 1; y <= base + 1; y++) {
                const dist: number = Math.abs(Date.UTC(y, month - 1, day) - todayUtc);
                if (dist < bestDist) {
                    bestDist = dist;
                    best = y;
                }
            }
            year = best;
        }
    }

    const utcDate: number = Date.UTC(year, month - 1, day);
    const check: Date = new Date(utcDate);
    if (check.getUTCFullYear() !== year || check.getUTCMonth() !== month - 1 || check.getUTCDate() !== day) {
        return invalid; // e.g. 31/02
    }

    return { valid: true, mondayUtc: mondayOf(utcDate), label: label };
}

function excelSerialToUtc(serial: number): number {
    return Math.round(serial - 25569) * MS_PER_DAY;
}

function utcToExcelSerial(utcMs: number): number {
    return Math.round(utcMs / MS_PER_DAY) + 25569;
}

function formatDate(utcMs: number): string {
    const d: Date = new Date(utcMs);
    return pad2(d.getUTCDate()) + "/" + pad2(d.getUTCMonth() + 1) + "/" + d.getUTCFullYear();
}

function pad2(n: number): string {
    return n < 10 ? "0" + n : "" + n;
}

/** Column letter for a resolved Key Figure, or "(absent)". */
function kfColumn(index: number): string {
    return index < 0 ? "(absent)" : columnLetter(index);
}

function columnLetter(index: number): string {
    let n: number = index;
    let out: string = "";
    while (n >= 0) {
        out = String.fromCharCode(65 + (n % 26)) + out;
        n = Math.floor(n / 26) - 1;
    }
    return out;
}

/* ===========================================================================
 * 6. PARSING HELPERS
 * =========================================================================*/

function readText(value: string | number | boolean | null | undefined): string {
    if (value === null || value === undefined) {
        return "";
    }
    if (typeof value === "string") {
        return value.trim();
    }
    return ("" + value).trim();
}

function isEmptyRow(row: (string | number | boolean)[]): boolean {
    for (let c: number = 0; c < row.length; c++) {
        if (readText(row[c]) !== "") {
            return false;
        }
    }
    return true;
}

/**
 * Robust numeric conversion. Blank and "-" become 0. Handles thousand
 * separators in either convention, trailing "%" and accounting negatives.
 */
function toNumber(value: string | number | boolean | null | undefined): number {
    if (value === null || value === undefined) {
        return 0;
    }
    if (typeof value === "number") {
        return isFinite(value) ? value : 0;
    }
    if (typeof value === "boolean") {
        return 0;
    }

    let text: string = value.replace(/ /g, " ").trim();
    if (text === "" || text === "-" || text === "--") {
        return 0;
    }

    let negative: boolean = false;
    if (text.charAt(0) === "(" && text.charAt(text.length - 1) === ")") {
        negative = true;
        text = text.substring(1, text.length - 1);
    }

    let percent: boolean = false;
    if (text.charAt(text.length - 1) === "%") {
        percent = true;
        text = text.substring(0, text.length - 1);
    }

    text = text.replace(/\s/g, "");

    const lastComma: number = text.lastIndexOf(",");
    const lastDot: number = text.lastIndexOf(".");
    if (lastComma >= 0 && lastDot >= 0) {
        if (lastComma > lastDot) {
            text = text.replace(/\./g, "").replace(",", ".");   // 1.234,56
        } else {
            text = text.replace(/,/g, "");                      // 1,234.56
        }
    } else if (lastComma >= 0) {
        const decimals: number = text.length - lastComma - 1;
        if (decimals > 0 && decimals <= 2 && text.indexOf(",") === lastComma) {
            text = text.replace(",", ".");                      // 1234,56
        } else {
            text = text.replace(/,/g, "");                      // 1,234,567
        }
    }

    const parsed: number = parseFloat(text);
    if (!isFinite(parsed)) {
        return 0;
    }
    let result: number = parsed;
    if (percent) {
        result = result / 100;
    }
    return negative ? -result : result;
}

/** True when a non-empty cell cannot be interpreted as a number. */
function isUnparseable(value: string | number | boolean | null | undefined): boolean {
    if (value === null || value === undefined) {
        return false;
    }
    if (typeof value === "number") {
        return !isFinite(value);
    }
    if (typeof value === "boolean") {
        return true;
    }
    const text: string = value.trim();
    if (text === "" || text === "-" || text === "--") {
        return false;
    }
    return !/\d/.test(text);
}

/** Source Accuracy/Bias to a ratio. 85 and 0.85 both become 0.85. */
function toRatioOrNull(value: string | number | boolean | null | undefined): number | null {
    if (readText(value) === "") {
        return null;
    }
    if (isUnparseable(value)) {
        return null;
    }
    const n: number = toNumber(value);
    if (!isFinite(n)) {
        return null;
    }
    return Math.abs(n) > 1.5 ? n / 100 : n;
}

function createCompositeKey(parts: string[]): string {
    return parts.join("");
}

/* ===========================================================================
 * 7. AGGREGATION
 * =========================================================================*/

/** Adds one source row into an entity/week bucket, creating them as needed. */
function addToEntity(
    map: Map<string, Entity>,
    key: string,
    market: string,
    product: string,
    description: string,
    customer: string,
    offset: number,
    sales: number,
    fdpCurrent: number,
    fdpPrevious: number,
    fdpSnapshot: number,
    statSnapshot: number,
    srcFdpAccuracy: number | null,
    srcFdpBias: number | null,
    srcStatAccuracy: number | null
): void {

    let entity: Entity | undefined = map.get(key);
    if (!entity) {
        entity = {
            market: market,
            product: product,
            description: description,
            customer: customer,
            weeks: new Map<number, WeekAggregate>()
        };
        map.set(key, entity);
    } else if (entity.description === "" && description !== "") {
        entity.description = description;
    }

    let bucket: WeekAggregate | undefined = entity.weeks.get(offset);
    if (!bucket) {
        bucket = {
            sales: 0,
            fdpCurrent: 0,
            fdpPrevious: 0,
            fdpSnapshot: 0,
            statSnapshot: 0,
            sourceRows: 0,
            srcFdpAccuracy: null,
            srcFdpBias: null,
            srcStatAccuracy: null
        };
        entity.weeks.set(offset, bucket);
    }

    bucket.sales += sales;
    bucket.fdpCurrent += fdpCurrent;
    bucket.fdpPrevious += fdpPrevious;
    bucket.fdpSnapshot += fdpSnapshot;
    bucket.statSnapshot += statSnapshot;
    bucket.sourceRows++;
    bucket.srcFdpAccuracy = srcFdpAccuracy;
    bucket.srcFdpBias = srcFdpBias;
    bucket.srcStatAccuracy = srcStatAccuracy;
}

/**
 * Totals over a set of week offsets. Weekly absolute errors are accumulated
 * per week, never after summing the period.
 */
function aggregateByEntityAndWeek(
    entity: Entity,
    offsets: number[],
    actualOf: (w: WeekAggregate) => number,
    forecastOf: (w: WeekAggregate) => number
): PeriodTotals {

    const totals: PeriodTotals = { actual: 0, forecast: 0, absError: 0, signedError: 0, weeks: 0 };

    for (let i: number = 0; i < offsets.length; i++) {
        const bucket: WeekAggregate | undefined = entity.weeks.get(offsets[i]);
        if (!bucket) {
            continue;
        }
        const actual: number = actualOf(bucket);
        const forecast: number = forecastOf(bucket);
        const signed: number = forecast - actual;

        totals.actual += actual;
        totals.forecast += forecast;
        totals.signedError += signed;
        totals.absError += Math.abs(signed);
        totals.weeks++;
    }
    return totals;
}

/**
 * Accuracy = 1 - SUM(weekly ABS error) / denominator, where the denominator is
 * SUM(Actual) or SUM(Forecast) per ACCURACY_DENOMINATOR. null when the
 * denominator is zero — dividing by it would be meaningless, not merely large.
 */
function calculatePeriodAccuracy(
    sumAbsError: number, sumActual: number, sumForecast: number
): number | null {
    const base: number = ACCURACY_DENOMINATOR === "forecast" ? sumForecast : sumActual;
    if (base === 0 || !isFinite(base)) {
        return null;
    }
    let accuracy: number = 1 - sumAbsError / Math.abs(base);
    if (CAP_ACCURACY_0_100) {
        if (accuracy < 0) {
            accuracy = 0;
        }
        if (accuracy > 1) {
            accuracy = 1;
        }
    }
    return accuracy;
}

/** Bias = SUM(Forecast - Actual) / SUM(Actual). null when not calculable. */
function calculatePeriodBias(sumSignedError: number, sumActual: number): number | null {
    if (sumActual === 0 || !isFinite(sumActual)) {
        return null;
    }
    return sumSignedError / Math.abs(sumActual);
}

function offsetRange(from: number, to: number): number[] {
    const out: number[] = [];
    for (let o: number = from; o <= to; o++) {
        if (o !== 0) {
            out.push(o);
        }
    }
    return out;
}

/* Accessors used by aggregateByEntityAndWeek. */
function actualSales(w: WeekAggregate): number { return w.sales; }
function forecastFdpSnapshot(w: WeekAggregate): number { return w.fdpSnapshot; }
function forecastStatSnapshot(w: WeekAggregate): number { return w.statSnapshot; }

/* ===========================================================================
 * 8. SOURCE KPI VALIDATION (Product-Customer, single week only)
 * =========================================================================*/

/**
 * Compares the recalculated single-week Accuracy/Bias with the values supplied
 * by IBP, at the only level where the source values are valid. Differences are
 * logged; the alerts always use the recalculated figures.
 */
function validateAgainstSourceKpis(productCustomerMap: Map<string, Entity>, counters: RunCounters): void {
    let samples: number = 0;
    const messages: string[] = [];

    productCustomerMap.forEach((entity: Entity) => {
        if (samples >= VALIDATION_MAX_SAMPLES) {
            return;
        }
        entity.weeks.forEach((bucket: WeekAggregate, offset: number) => {
            if (samples >= VALIDATION_MAX_SAMPLES || offset >= 0) {
                return;
            }
            if (bucket.sourceRows !== 1 || bucket.sales === 0) {
                return; // duplicates make the source value non-comparable
            }
            if (bucket.srcFdpAccuracy === null && bucket.srcFdpBias === null) {
                return;
            }

            samples++;
            counters.validationChecks++;

            const signed: number = bucket.fdpSnapshot - bucket.sales;
            const recalcAccuracy: number | null =
                calculatePeriodAccuracy(Math.abs(signed), bucket.sales, bucket.fdpSnapshot);
            const recalcBias: number | null = calculatePeriodBias(signed, bucket.sales);

            let mismatch: boolean = false;
            if (bucket.srcFdpAccuracy !== null && recalcAccuracy !== null &&
                Math.abs(bucket.srcFdpAccuracy - recalcAccuracy) > VALIDATION_TOLERANCE) {
                mismatch = true;
            }
            if (bucket.srcFdpBias !== null && recalcBias !== null &&
                Math.abs(bucket.srcFdpBias - recalcBias) > VALIDATION_TOLERANCE) {
                mismatch = true;
            }

            if (mismatch) {
                counters.validationMismatches++;
                if (messages.length < 5) {
                    messages.push(
                        entity.market + " / " + entity.product + " / " + entity.customer + " W" + offset +
                        ": source acc=" + fmtRatio(bucket.srcFdpAccuracy) + " recalculated acc=" + fmtRatio(recalcAccuracy) +
                        ", source bias=" + fmtRatio(bucket.srcFdpBias) + " recalculated bias=" + fmtRatio(recalcBias)
                    );
                }
            }
        });
    });

    console.log(
        "Source KPI validation: " + counters.validationChecks + " single-week Product-Customer checks, " +
        counters.validationMismatches + " outside the " + (VALIDATION_TOLERANCE * 100) + " pp tolerance."
    );
    for (let i: number = 0; i < messages.length; i++) {
        console.log("  validation sample: " + messages[i]);
    }
}

function fmtRatio(v: number | null): string {
    return v === null ? "n/a" : (Math.round(v * 1000) / 10) + "%";
}

/* ===========================================================================
 * 9. ALERT 1 — Final Demand Plan change vs W-1
 * =========================================================================*/

function runAlert1(
    productCustomerMap: Map<string, Entity>,
    productMap: Map<string, Entity>,
    weekLabels: Map<number, string>,
    weekStarts: Map<number, number>
): AlertOutput {

    const headers: string[] = [
        "Aggregation Level", "Market", "Product", "Product Description", "Customer",
        "Week", "Week Start Date",
        "Final Demand Plan W-1", "Current Final Demand Plan",
        "Difference", "Absolute Difference", "Variation %", "Alert Type"
    ];

    interface Row1 {
        level: string; market: string; product: string; description: string; customer: string;
        week: string; weekStart: number; previous: number; current: number;
        difference: number; absDifference: number; variation: number | null; alertType: string;
        sortVariation: number;
    }

    const rows: Row1[] = [];
    let pcCount: number = 0;
    let pCount: number = 0;

    const collect = (map: Map<string, Entity>, level: string): number => {
        let count: number = 0;
        map.forEach((entity: Entity) => {
            for (let offset: number = 1; offset <= ALERT1_FUTURE_HORIZON_WEEKS; offset++) {
                const bucket: WeekAggregate | undefined = entity.weeks.get(offset);
                if (!bucket) {
                    continue;
                }
                const current: number = bucket.fdpCurrent;
                const previous: number = bucket.fdpPrevious;

                if (current === 0 && previous === 0) {
                    continue;
                }

                const difference: number = current - previous;
                const absDifference: number = Math.abs(difference);

                let variation: number | null = null;
                let alertType: string = "";
                let qualifies: boolean = false;

                if (previous === 0) {
                    // New forecast — no percentage can be calculated.
                    alertType = "New Forecast";
                    variation = null;
                    qualifies = absDifference >= ALERT1_MIN_ABS_DIFF;
                } else if (current === 0) {
                    alertType = "Forecast Removed";
                    variation = -1;
                    qualifies = Math.abs(variation) > ALERT1_PCT_THRESHOLD && absDifference >= ALERT1_MIN_ABS_DIFF;
                } else {
                    variation = difference / Math.abs(previous);
                    alertType = difference > 0 ? "Forecast Increase" : "Forecast Decrease";
                    qualifies = Math.abs(variation) > ALERT1_PCT_THRESHOLD && absDifference >= ALERT1_MIN_ABS_DIFF;
                }

                if (!qualifies) {
                    continue;
                }

                const label: string | undefined = weekLabels.get(offset);
                const start: number | undefined = weekStarts.get(offset);

                rows.push({
                    level: level,
                    market: entity.market,
                    product: entity.product,
                    description: entity.description,
                    customer: entity.customer,
                    week: label === undefined ? "W+" + offset : label,
                    weekStart: start === undefined ? 0 : start,
                    previous: previous,
                    current: current,
                    difference: difference,
                    absDifference: absDifference,
                    variation: variation,
                    alertType: alertType,
                    sortVariation: variation === null ? -1 : Math.abs(variation)
                });
                count++;
            }
        });
        return count;
    };

    pcCount = collect(productCustomerMap, AGG_PRODUCT_CUSTOMER);
    pCount = collect(productMap, AGG_PRODUCT);

    rows.sort((a: Row1, b: Row1) => {
        if (b.sortVariation !== a.sortVariation) {
            return b.sortVariation - a.sortVariation;
        }
        return b.absDifference - a.absDifference;
    });

    const out: (string | number)[][] = [];
    for (let i: number = 0; i < rows.length; i++) {
        const r: Row1 = rows[i];
        out.push([
            r.level, r.market, r.product, r.description, r.customer,
            r.week, r.weekStart === 0 ? "" : utcToExcelSerial(r.weekStart),
            round2(r.previous), round2(r.current), round2(r.difference), round2(r.absDifference),
            r.variation === null ? "" : r.variation,
            r.alertType
        ]);
    }

    return {
        headers: headers,
        rows: out,
        percentColumns: [11],
        volumeColumns: [7, 8, 9, 10],
        dateColumns: [6],
        descriptionColumn: 3,
        productCustomerCount: pcCount,
        productCount: pCount
    };
}

/* ===========================================================================
 * 10. ALERT 2 — Accuracy and Bias deterioration
 * =========================================================================*/

function runAlert2(
    productCustomerMap: Map<string, Entity>,
    productMap: Map<string, Entity>,
    counters: RunCounters
): AlertOutput {

    const headers: string[] = [
        "Aggregation Level", "Market", "Product", "Product Description", "Customer",
        "Accuracy W-1", "Baseline Accuracy W-5 to W-2", "Accuracy Difference",
        "Bias W-1", "Baseline Bias W-5 to W-2", "Bias Signed Difference",
        "Bias Deterioration", "Bias Direction", "Alert Reason",
        "W-1 Actual", "Baseline Actual W-5 to W-2"
    ];

    interface Row2 {
        level: string; market: string; product: string; description: string; customer: string;
        accW1: number; accBase: number; accDiff: number;
        biasW1: number; biasBase: number; biasSignedDiff: number; biasDeterioration: number;
        biasDirection: string; reason: string; actualW1: number; actualBase: number;
    }

    const baselineOffsets: number[] = [-5, -4, -3, -2];
    const rows: Row2[] = [];

    const collect = (map: Map<string, Entity>, level: string): number => {
        let count: number = 0;
        map.forEach((entity: Entity) => {
            const w1: WeekAggregate | undefined = entity.weeks.get(-1);
            if (!w1) {
                counters.excludedInsufficientWeeks++;
                return;
            }
            if (w1.sales === 0) {
                counters.excludedZeroActual++;
                return;
            }
            if (w1.sales < ALERT2_MIN_ACTUAL) {
                counters.excludedBelowMinimumVolume++;
                return;
            }

            const baseline: PeriodTotals = aggregateByEntityAndWeek(
                entity, baselineOffsets, actualSales, forecastFdpSnapshot
            );
            if (baseline.weeks < ALERT2_MIN_BASELINE_WEEKS) {
                counters.excludedInsufficientWeeks++;
                return;
            }
            if (baseline.actual === 0) {
                counters.excludedZeroActual++;
                return;
            }

            const signedW1: number = w1.fdpSnapshot - w1.sales;
            const accW1: number | null =
                calculatePeriodAccuracy(Math.abs(signedW1), w1.sales, w1.fdpSnapshot);
            const biasW1: number | null = calculatePeriodBias(signedW1, w1.sales);
            const accBase: number | null =
                calculatePeriodAccuracy(baseline.absError, baseline.actual, baseline.forecast);
            const biasBase: number | null = calculatePeriodBias(baseline.signedError, baseline.actual);

            if (accW1 === null || biasW1 === null || accBase === null || biasBase === null) {
                counters.excludedZeroActual++;
                return;
            }

            const accDiff: number = accW1 - accBase;
            const biasSignedDiff: number = biasW1 - biasBase;
            const biasDeterioration: number = Math.abs(biasW1) - Math.abs(biasBase);

            const accuracyAlert: boolean = accDiff < ALERT2_ACCURACY_DETERIORATION_THRESHOLD;
            const biasAlert: boolean = biasDeterioration > ALERT2_BIAS_DETERIORATION_THRESHOLD;
            if (!accuracyAlert && !biasAlert) {
                return;
            }

            let reason: string = "Accuracy and Bias deterioration";
            if (accuracyAlert && !biasAlert) {
                reason = "Accuracy deterioration";
            } else if (!accuracyAlert && biasAlert) {
                reason = "Bias deterioration";
            }

            let direction: string = "No material change";
            if (biasDeterioration > ALERT2_BIAS_MATERIALITY) {
                direction = "Further from zero";
            } else if (biasDeterioration < -ALERT2_BIAS_MATERIALITY) {
                direction = "Closer to zero";
            }

            rows.push({
                level: level,
                market: entity.market,
                product: entity.product,
                description: entity.description,
                customer: entity.customer,
                accW1: accW1,
                accBase: accBase,
                accDiff: accDiff,
                biasW1: biasW1,
                biasBase: biasBase,
                biasSignedDiff: biasSignedDiff,
                biasDeterioration: biasDeterioration,
                biasDirection: direction,
                reason: reason,
                actualW1: w1.sales,
                actualBase: baseline.actual
            });
            count++;
        });
        return count;
    };

    const pcCount: number = collect(productCustomerMap, AGG_PRODUCT_CUSTOMER);
    const pCount: number = collect(productMap, AGG_PRODUCT);

    rows.sort((a: Row2, b: Row2) => {
        if (a.accDiff !== b.accDiff) {
            return a.accDiff - b.accDiff;                       // largest deterioration first
        }
        return b.biasDeterioration - a.biasDeterioration;
    });

    const out: (string | number)[][] = [];
    for (let i: number = 0; i < rows.length; i++) {
        const r: Row2 = rows[i];
        out.push([
            r.level, r.market, r.product, r.description, r.customer,
            r.accW1, r.accBase, r.accDiff,
            r.biasW1, r.biasBase, r.biasSignedDiff, r.biasDeterioration,
            r.biasDirection, r.reason,
            round2(r.actualW1), round2(r.actualBase)
        ]);
    }

    return {
        headers: headers,
        rows: out,
        percentColumns: [5, 6, 7, 8, 9, 10, 11],
        volumeColumns: [14, 15],
        dateColumns: [],
        descriptionColumn: 3,
        productCustomerCount: pcCount,
        productCount: pCount
    };
}

/* ===========================================================================
 * 11. ALERT 3 — Forecast vs recent Sales
 * =========================================================================*/

function runAlert3(
    productCustomerMap: Map<string, Entity>,
    productMap: Map<string, Entity>,
    currentMonday: number
): AlertOutput {

    const headers: string[] = [
        "Aggregation Level", "Market", "Product", "Product Description", "Customer",
        "Sales Last 4 Weeks", "Forecast Next 4 Weeks", "Difference", "Absolute Difference",
        "Variation %", "Alert Type", "Historical Period", "Future Period"
    ];

    const historyOffsets: number[] = [-4, -3, -2, -1];
    const futureOffsets: number[] = [1, 2, 3, 4];

    const historyLabel: string =
        "W-4 to W-1 (" + formatDate(currentMonday - 4 * MS_PER_WEEK) +
        " - " + formatDate(currentMonday - 1 * MS_PER_WEEK + 6 * MS_PER_DAY) + ")";
    const futureLabel: string =
        "W+1 to W+4 (" + formatDate(currentMonday + 1 * MS_PER_WEEK) +
        " - " + formatDate(currentMonday + 4 * MS_PER_WEEK + 6 * MS_PER_DAY) + ")";

    interface Row3 {
        level: string; market: string; product: string; description: string; customer: string;
        sales: number; forecast: number; difference: number; absDifference: number;
        variation: number | null; alertType: string; sortVariation: number;
    }

    const rows: Row3[] = [];

    const collect = (map: Map<string, Entity>, level: string): number => {
        let count: number = 0;
        map.forEach((entity: Entity) => {
            let sales: number = 0;
            for (let i: number = 0; i < historyOffsets.length; i++) {
                const b: WeekAggregate | undefined = entity.weeks.get(historyOffsets[i]);
                if (b) {
                    sales += b.sales;
                }
            }
            let forecast: number = 0;
            for (let i: number = 0; i < futureOffsets.length; i++) {
                const b: WeekAggregate | undefined = entity.weeks.get(futureOffsets[i]);
                if (b) {
                    forecast += b.fdpCurrent;
                }
            }

            const difference: number = forecast - sales;
            const absDifference: number = Math.abs(difference);

            let variation: number | null = null;
            let alertType: string = "";
            let qualifies: boolean = false;

            if (sales <= 0 && forecast === 0) {
                return;                                          // nothing to report
            }

            if (sales <= 0) {
                alertType = "Forecast with no recent sales";
                variation = null;
                qualifies = forecast >= ALERT3_MIN_FUTURE_VOLUME;
            } else if (forecast === 0) {
                alertType = "No future forecast";
                variation = -1;
                qualifies = Math.abs(variation) > ALERT3_PCT_THRESHOLD && sales >= ALERT3_MIN_HISTORICAL_VOLUME;
            } else {
                variation = difference / sales;
                alertType = difference > 0 ? "Forecast above recent sales" : "Forecast below recent sales";
                qualifies = Math.abs(variation) > ALERT3_PCT_THRESHOLD && sales >= ALERT3_MIN_HISTORICAL_VOLUME;
            }

            if (!qualifies) {
                return;
            }

            rows.push({
                level: level,
                market: entity.market,
                product: entity.product,
                description: entity.description,
                customer: entity.customer,
                sales: sales,
                forecast: forecast,
                difference: difference,
                absDifference: absDifference,
                variation: variation,
                alertType: alertType,
                sortVariation: variation === null ? -1 : Math.abs(variation)
            });
            count++;
        });
        return count;
    };

    const pcCount: number = collect(productCustomerMap, AGG_PRODUCT_CUSTOMER);
    const pCount: number = collect(productMap, AGG_PRODUCT);

    rows.sort((a: Row3, b: Row3) => {
        if (b.sortVariation !== a.sortVariation) {
            return b.sortVariation - a.sortVariation;
        }
        return b.absDifference - a.absDifference;
    });

    const out: (string | number)[][] = [];
    for (let i: number = 0; i < rows.length; i++) {
        const r: Row3 = rows[i];
        out.push([
            r.level, r.market, r.product, r.description, r.customer,
            round2(r.sales), round2(r.forecast), round2(r.difference), round2(r.absDifference),
            r.variation === null ? "" : r.variation,
            r.alertType, historyLabel, futureLabel
        ]);
    }

    return {
        headers: headers,
        rows: out,
        percentColumns: [9],
        volumeColumns: [5, 6, 7, 8],
        dateColumns: [],
        descriptionColumn: 3,
        productCustomerCount: pcCount,
        productCount: pCount
    };
}

/* ===========================================================================
 * 12. ALERT 4 — Statistical Forecast vs Final Demand Plan (negative FVA)
 * =========================================================================*/

function runAlert4(
    productCustomerMap: Map<string, Entity>,
    productMap: Map<string, Entity>,
    counters: RunCounters
): AlertOutput {

    const headers: string[] = [
        "Aggregation Level", "Market", "Product", "Product Description", "Customer",
        "Sales History Last 4 Weeks", "Statistical Forecast Snapshot Last 4 Weeks",
        "Final Demand Plan Snapshot Last 4 Weeks",
        "Statistical Forecast Absolute Error", "Final Demand Plan Absolute Error",
        "Statistical Forecast Accuracy", "Final Demand Plan Accuracy",
        "Accuracy Difference", "FVA Classification"
    ];

    const historyOffsets: number[] = [-4, -3, -2, -1];

    interface Row4 {
        level: string; market: string; product: string; description: string; customer: string;
        sales: number; stat: number; fdp: number; statAbsError: number; fdpAbsError: number;
        statAccuracy: number; fdpAccuracy: number; accuracyDifference: number;
    }

    const rows: Row4[] = [];

    const collect = (map: Map<string, Entity>, level: string): number => {
        let count: number = 0;
        map.forEach((entity: Entity) => {
            const statTotals: PeriodTotals = aggregateByEntityAndWeek(
                entity, historyOffsets, actualSales, forecastStatSnapshot
            );
            const fdpTotals: PeriodTotals = aggregateByEntityAndWeek(
                entity, historyOffsets, actualSales, forecastFdpSnapshot
            );

            if (statTotals.weeks < ALERT4_MIN_HISTORY_WEEKS) {
                counters.excludedInsufficientWeeks++;
                return;
            }
            if (statTotals.actual <= 0) {
                counters.excludedZeroActual++;
                return;
            }
            if (statTotals.actual < ALERT4_MIN_ACTUAL) {
                counters.excludedBelowMinimumVolume++;
                return;
            }

            const statAccuracy: number | null =
                calculatePeriodAccuracy(statTotals.absError, statTotals.actual, statTotals.forecast);
            const fdpAccuracy: number | null =
                calculatePeriodAccuracy(fdpTotals.absError, fdpTotals.actual, fdpTotals.forecast);
            if (statAccuracy === null || fdpAccuracy === null) {
                counters.excludedZeroActual++;
                return;
            }

            const accuracyDifference: number = statAccuracy - fdpAccuracy;
            if (accuracyDifference <= ALERT4_ACCURACY_DIFF_THRESHOLD) {
                return;
            }

            rows.push({
                level: level,
                market: entity.market,
                product: entity.product,
                description: entity.description,
                customer: entity.customer,
                sales: statTotals.actual,
                stat: statTotals.forecast,
                fdp: fdpTotals.forecast,
                statAbsError: statTotals.absError,
                fdpAbsError: fdpTotals.absError,
                statAccuracy: statAccuracy,
                fdpAccuracy: fdpAccuracy,
                accuracyDifference: accuracyDifference
            });
            count++;
        });
        return count;
    };

    const pcCount: number = collect(productCustomerMap, AGG_PRODUCT_CUSTOMER);
    const pCount: number = collect(productMap, AGG_PRODUCT);

    rows.sort((a: Row4, b: Row4) => {
        if (b.accuracyDifference !== a.accuracyDifference) {
            return b.accuracyDifference - a.accuracyDifference;
        }
        return b.sales - a.sales;
    });

    const out: (string | number)[][] = [];
    for (let i: number = 0; i < rows.length; i++) {
        const r: Row4 = rows[i];
        out.push([
            r.level, r.market, r.product, r.description, r.customer,
            round2(r.sales), round2(r.stat), round2(r.fdp),
            round2(r.statAbsError), round2(r.fdpAbsError),
            r.statAccuracy, r.fdpAccuracy, r.accuracyDifference,
            "Negative FVA"
        ]);
    }

    return {
        headers: headers,
        rows: out,
        percentColumns: [10, 11, 12],
        volumeColumns: [5, 6, 7, 8, 9],
        dateColumns: [],
        descriptionColumn: 3,
        productCustomerCount: pcCount,
        productCount: pCount
    };
}

/* ===========================================================================
 * 13. OUTPUT MANAGEMENT
 * =========================================================================*/

/**
 * Confirms every output table exists before the script reports success.
 *
 * Without this, a failure part-way through publishing leaves some worksheets
 * holding the PREVIOUS run's results while the script still looks like it
 * worked — and the downstream flow then reads stale data, or fails on a table
 * that was never created. Better to fail loudly here.
 */
function verifyOutputTables(workbook: ExcelScript.Workbook): void {
    const expected: string[] = [
        OUT_TABLE_ALERT1, OUT_TABLE_ALERT2, OUT_TABLE_ALERT3, OUT_TABLE_ALERT4
    ];
    const missing: string[] = [];
    for (let i: number = 0; i < expected.length; i++) {
        if (!workbook.getTable(expected[i])) {
            missing.push(expected[i]);
        }
    }
    if (missing.length > 0) {
        throw new Error(
            "Output incomplete — these tables were not created: " + missing.join(", ") +
            ". Worksheets from a previous run may still be present and STALE. " +
            "Re-run the script; if it fails again, lower WRITE_CHUNK_CELLS."
        );
    }
    console.log("All four output tables verified present.");
}

function publishAlert(
    workbook: ExcelScript.Workbook,
    sheetName: string,
    tableName: string,
    output: AlertOutput
): void {
    const sheet: ExcelScript.Worksheet = createOrResetOutputSheet(workbook, sheetName, tableName);
    writeOutputTable(sheet, tableName, output);
    console.log(
        sheetName + ": " + output.rows.length + " row(s) — " +
        output.productCustomerCount + " Product-Customer, " + output.productCount + " Product."
    );
}

/** Deletes any previous table with the same name, then recreates the worksheet. */
function createOrResetOutputSheet(
    workbook: ExcelScript.Workbook,
    sheetName: string,
    tableName: string
): ExcelScript.Worksheet {

    const existingTable: ExcelScript.Table | undefined = workbook.getTable(tableName);
    if (existingTable) {
        existingTable.delete();
    }

    const existingSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet(sheetName);
    if (existingSheet) {
        existingSheet.delete();
    }

    return workbook.addWorksheet(sheetName);
}

function writeOutputTable(
    sheet: ExcelScript.Worksheet,
    tableName: string,
    output: AlertOutput
): void {

    const columnCount: number = output.headers.length;
    const isEmpty: boolean = output.rows.length === 0;

    // The table is created even when nothing alerted. A downstream Power Automate
    // flow reads each alert table BY NAME, so a missing table would fail the whole
    // weekly run on a quiet week. An empty result therefore still gets its table,
    // carrying one placeholder row that consumers are expected to skip (it has no
    // Market and no Product).
    const bodyRows: (string | number)[][] = isEmpty
        ? [buildNoAlertsRow(columnCount)]
        : output.rows;
    const rowCount: number = bodyRows.length;

    // Header, then the body in one bulk write (chunked for very large results).
    sheet.getRangeByIndexes(0, 0, 1, columnCount).setValues([output.headers]);

    writeRowsChunked(sheet, bodyRows, columnCount);

    const tableRange: ExcelScript.Range = sheet.getRangeByIndexes(0, 0, rowCount + 1, columnCount);
    const table: ExcelScript.Table = sheet.addTable(tableRange, true);
    table.setName(tableName);
    table.setPredefinedTableStyle(TABLE_STYLE);
    table.setShowFilterButton(true);

    if (isEmpty) {
        // Explain the placeholder to whoever opens the worksheet.
        sheet.getRangeByIndexes(rowCount + 2, 0, 1, 1).setValue(
            "No combinations met the alert criteria for this execution (" +
            formatDate(getMondayOfCurrentWeek()) + "). The table above is kept so " +
            "downstream automation still finds it; adjust the threshold constants " +
            "at the top of the script if this is unexpected."
        );
        sheet.getRangeByIndexes(0, 0, 1, columnCount).getFormat().autofitColumns();
        return;
    }

    formatOutput(sheet, output, rowCount, columnCount);
}

/**
 * Writes the body rows in blocks that stay under the payload limit, halving a
 * block and retrying if it is still rejected. Mirrors readBlock.
 */
function writeRowsChunked(
    sheet: ExcelScript.Worksheet,
    rows: (string | number)[][],
    columnCount: number
): void {

    const perChunk: number = Math.max(
        MIN_WRITE_CHUNK_ROWS,
        Math.floor(WRITE_CHUNK_CELLS / Math.max(1, columnCount))
    );

    let written: number = 0;
    while (written < rows.length) {
        let size: number = Math.min(perChunk, rows.length - written);
        for (;;) {
            try {
                sheet.getRangeByIndexes(1 + written, 0, size, columnCount)
                    .setValues(rows.slice(written, written + size));
                break;
            } catch (e) {
                if (size <= MIN_WRITE_CHUNK_ROWS) {
                    throw new Error(
                        "Could not write output rows " + (written + 1) + "-" +
                        (written + size) + " even at the minimum block size (" +
                        MIN_WRITE_CHUNK_ROWS + " rows). Lower WRITE_CHUNK_CELLS. " +
                        "Underlying error: " + String(e)
                    );
                }
                size = Math.max(MIN_WRITE_CHUNK_ROWS, Math.floor(size / 2));
                console.log("  write block too large; retrying with " + size + " rows");
            }
        }
        written += size;
    }
}

/** Placeholder body row for an alert that produced nothing this week. */
function buildNoAlertsRow(columnCount: number): (string | number)[] {
    const row: (string | number)[] = [NO_ALERTS_MARKER];
    for (let c: number = 1; c < columnCount; c++) {
        row.push("");
    }
    return row;
}

function formatOutput(
    sheet: ExcelScript.Worksheet,
    output: AlertOutput,
    rowCount: number,
    columnCount: number
): void {

    for (let i: number = 0; i < output.percentColumns.length; i++) {
        applyColumnFormat(sheet, output.percentColumns[i], rowCount, FMT_PERCENT);
    }
    for (let i: number = 0; i < output.volumeColumns.length; i++) {
        applyColumnFormat(sheet, output.volumeColumns[i], rowCount, FMT_VOLUME);
    }
    for (let i: number = 0; i < output.dateColumns.length; i++) {
        applyColumnFormat(sheet, output.dateColumns[i], rowCount, FMT_DATE);
    }

    sheet.getFreezePanes().freezeRows(1);

    sheet.getRangeByIndexes(0, 0, rowCount + 1, columnCount).getFormat().autofitColumns();

    if (output.descriptionColumn >= 0) {
        const descFormat: ExcelScript.RangeFormat =
            sheet.getRangeByIndexes(0, output.descriptionColumn, 1, 1).getFormat();
        if (descFormat.getColumnWidth() > MAX_DESCRIPTION_COLUMN_WIDTH) {
            descFormat.setColumnWidth(MAX_DESCRIPTION_COLUMN_WIDTH);
        }
    }
}

/** Applies one number format to a whole data column in a single call. */
function applyColumnFormat(
    sheet: ExcelScript.Worksheet,
    columnIndex: number,
    rowCount: number,
    format: string
): void {
    if (rowCount <= 0) {
        return;
    }
    const formats: string[][] = [];
    for (let r: number = 0; r < rowCount; r++) {
        formats.push([format]);
    }
    sheet.getRangeByIndexes(1, columnIndex, rowCount, 1).setNumberFormat(formats);
}

function round2(value: number): number {
    return Math.round(value * 100) / 100;
}

/* ===========================================================================
 * 14. EXECUTION SUMMARY
 * =========================================================================*/

function returnExecutionSummary(summary: ExecutionSummary): void {
    console.log("=== Demand Planning alerts — execution summary ===");
    console.log("Current week Monday                     : " + summary.currentWeekMonday);
    console.log("Source rows read                        : " + summary.sourceRowsRead);
    console.log("Rows ignored — invalid week/date        : " + summary.sourceRowsIgnoredInvalidDate);
    console.log("Rows ignored — invalid numeric values   : " + summary.sourceRowsIgnoredInvalidNumber);
    console.log("Rows with invalid numeric cells         : " + summary.sourceRowsWithInvalidNumericCells);
    console.log("Rows outside horizon / current week     : " + summary.sourceRowsOutsideAnalysedHorizon);
    console.log("Alert 1 Product-Customer / Product      : " + summary.alert1ProductCustomerCount + " / " + summary.alert1ProductCount);
    console.log("Alert 2 Product-Customer / Product      : " + summary.alert2ProductCustomerCount + " / " + summary.alert2ProductCount);
    console.log("Alert 3 Product-Customer / Product      : " + summary.alert3ProductCustomerCount + " / " + summary.alert3ProductCount);
    console.log("Alert 4 Product-Customer / Product      : " + summary.alert4ProductCustomerCount + " / " + summary.alert4ProductCount);
    console.log("Excluded — insufficient history weeks   : " + summary.excludedInsufficientHistoricalWeeks);
    console.log("Excluded — total Actual equals zero     : " + summary.excludedZeroActual);
    console.log("Excluded — below minimum volume         : " + summary.excludedBelowMinimumVolume +
        "  (ALERT2_MIN_ACTUAL=" + ALERT2_MIN_ACTUAL + ", ALERT4_MIN_ACTUAL=" + ALERT4_MIN_ACTUAL + ")");
    console.log("Source KPI validation checks / mismatch : " + summary.sourceKpiValidationChecks + " / " + summary.sourceKpiValidationMismatches);
    console.log("Output worksheets                       : " + summary.outputWorksheets.join(", "));
}
