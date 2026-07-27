/**
 * Minimal ambient declarations for the Office Scripts ExcelScript API.
 *
 * This file exists ONLY so the script type-checks in the local editor.
 * It is NOT part of the Office Script — do not paste it into Excel Automate.
 * It declares just the members used by src/DemandAlerts.ts.
 *
 * Full API reference:
 * https://learn.microsoft.com/office/dev/scripts/develop/scripting-fundamentals
 */

/** Office Scripts exposes console.log for the run log. */
declare const console: {
    log(message?: string | number | boolean): void;
};

declare namespace ExcelScript {

    type CellValue = string | number | boolean;

    interface RangeFormat {
        autofitColumns(): void;
        getColumnWidth(): number;
        setColumnWidth(width: number): void;
    }

    interface Range {
        getValues(): CellValue[][];
        setValues(values: CellValue[][]): void;
        setValue(value: CellValue): void;
        setNumberFormat(numberFormat: string[][]): void;
        getFormat(): RangeFormat;
        getRowIndex(): number;
        getColumnIndex(): number;
        getRowCount(): number;
        getColumnCount(): number;
    }

    interface Table {
        setName(name: string): void;
        getName(): string;
        setPredefinedTableStyle(style: string): void;
        setShowFilterButton(showFilterButton: boolean): void;
        delete(): void;
    }

    interface FreezePane {
        freezeRows(count: number): void;
    }

    interface Worksheet {
        getName(): string;
        getUsedRange(): Range | undefined;
        getRange(address: string): Range;
        getRangeByIndexes(
            startRow: number,
            startColumn: number,
            rowCount: number,
            columnCount: number
        ): Range;
        addTable(address: Range | string, hasHeaders: boolean): Table;
        getFreezePanes(): FreezePane;
        activate(): void;
        delete(): void;
    }

    interface Workbook {
        getWorksheet(name: string): Worksheet | undefined;
        getWorksheets(): Worksheet[];
        addWorksheet(name?: string): Worksheet;
        getTable(name: string): Table | undefined;
        getTables(): Table[];
    }
}
