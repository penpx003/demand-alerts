/**
 * ============================================================================
 * ListTables — diagnostic Office Script
 * ============================================================================
 * Run this in the SAME workbook you are pointing Power Automate at, to prove
 * what tables actually exist and on which worksheets.
 *
 * Use it when "List rows present in a table" reports that the workbook has no
 * tables. It answers the only question that matters: is the table missing, or
 * is the flow looking at a different file?
 *
 * Paste into Excel -> Automate -> New Script, name it "List Tables", run it,
 * and read the output pane.
 * ============================================================================
 */

function main(workbook: ExcelScript.Workbook): string {
    const lines: string[] = [];

    const sheets: ExcelScript.Worksheet[] = workbook.getWorksheets();
    lines.push("Worksheets (" + sheets.length + "):");
    for (let i: number = 0; i < sheets.length; i++) {
        lines.push("  - " + sheets[i].getName());
    }

    const tables: ExcelScript.Table[] = workbook.getTables();
    lines.push("");
    lines.push("Tables (" + tables.length + "):");

    if (tables.length === 0) {
        lines.push("  (none)");
        lines.push("");
        lines.push("No tables in this workbook. Either the 'Demand Alerts' script has");
        lines.push("not been run here, or it failed before writing its output.");
    } else {
        for (let i: number = 0; i < tables.length; i++) {
            const table: ExcelScript.Table = tables[i];
            const range: ExcelScript.Range = table.getRange();
            lines.push(
                "  - " + table.getName() +
                "   on '" + table.getWorksheet().getName() + "'" +
                "   rows(incl. header)=" + range.getRowCount() +
                "   cols=" + range.getColumnCount()
            );
        }
    }

    // The four the weekly flow reads by name.
    const expected: string[] = [
        "tblAlert1FDPChange",
        "tblAlert2AccuracyBias",
        "tblAlert3ForecastVsSales",
        "tblAlert4StatVsFDP"
    ];
    lines.push("");
    lines.push("Expected by the Power Automate flow:");
    for (let i: number = 0; i < expected.length; i++) {
        const found: ExcelScript.Table | undefined = workbook.getTable(expected[i]);
        lines.push("  " + (found ? "FOUND   " : "MISSING ") + expected[i]);
    }

    const report: string = lines.join("\n");
    console.log(report);
    return report;
}
