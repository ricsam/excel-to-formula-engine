/**
 * The conversion pipeline: an .xlsx package in, `WorkbookData` out.
 */

import type {
  SerializedCellValue,
  WorkbookDataCellStyle,
  WorkbookDataConditionalStyle,
  WorkbookDataNamedExpression,
  WorkbookDataSheet,
  WorkbookDataTable,
} from "@ricsam/formula-engine";
import { readXlsxPackage } from "./xlsx/package";
import {
  parseSharedStrings,
  parseTables,
  parseWorkbook,
  type RawTable,
} from "./xlsx/workbook";
import { parseWorksheet, type RawCell } from "./xlsx/worksheet";
import { parseStyles, type ExcelFormat } from "./xlsx/styles";
import { offsetFormulaReferences, translateFormula } from "./translate/formula";
import { translateConditionalFormat } from "./translate/conditional-format";
import { excelSerialToIso } from "./translate/dates";
import {
  parseA1Range,
  rangeReferenceToSpreadsheetRange,
  toA1,
  toSpreadsheetRange,
} from "./a1";
import type {
  ConvertOptions,
  ConvertResult,
  Diagnostic,
  ExcelCellMetadata,
  ExcelSheetMetadata,
  ExcelWorkbookData,
} from "./types";

/**
 * Excel's built-in defined names. They configure print areas and filters rather
 * than naming a value, so importing them as named expressions would be noise.
 */
const BUILTIN_DEFINED_NAMES = new Set([
  "_xlnm.print_area",
  "_xlnm.print_titles",
  "_xlnm._filterdatabase",
  "_xlnm.criteria",
  "_xlnm.extract",
  "_xlnm.database",
  "_xlnm.sheet_title",
]);

export async function convertWorkbook(
  source: Blob | ArrayBuffer | Uint8Array,
  options: ConvertOptions = {}
): Promise<ConvertResult> {
  const {
    styles: importStyles = true,
    conditionalFormatting: importConditionalFormatting = true,
    tables: importTables = true,
    definedNames: importDefinedNames = true,
    cellMetadata: importCellMetadata = true,
    convertDates = true,
  } = options;

  const diagnostics: Diagnostic[] = [];
  const pkg = await readXlsxPackage(source);

  const sharedStrings = parseSharedStrings(pkg);
  const { sheets: sheetEntries, definedNames, date1904 } = parseWorkbook(pkg);
  const styleTable = parseStyles(pkg);

  const selected = options.sheets
    ? sheetEntries.filter((entry) => options.sheets!.includes(entry.name))
    : sheetEntries;

  const sheets: WorkbookDataSheet<ExcelCellMetadata, ExcelSheetMetadata>[] = [];
  const tables: WorkbookDataTable[] = [];
  const cellStyles: WorkbookDataCellStyle[] = [];
  const conditionalStyles: WorkbookDataConditionalStyle[] = [];
  const importedSheetNames = new Set<string>();

  for (const entry of selected) {
    const worksheet = parseWorksheet(pkg, entry.partPath, sharedStrings);
    importedSheetNames.add(entry.name);

    const content = new Map<string, SerializedCellValue>();
    const cellMetadata = new Map<string, ExcelCellMetadata>();
    // Cells sharing one style are collected so each distinct style becomes a
    // single engine rule over many areas rather than one rule per cell.
    const styleGroups = new Map<number, RawCell[]>();

    for (const cell of worksheet.cells) {
      const format =
        cell.styleIndex === undefined
          ? undefined
          : styleTable.formats[cell.styleIndex];

      const value = cellValue(cell, format, {
        convertDates,
        date1904,
        sheetName: entry.name,
        diagnostics,
      });
      if (value !== undefined) {
        content.set(cell.reference, value);
      }

      if (importStyles && cell.styleIndex !== undefined && format) {
        if (Object.keys(format.style).length > 0) {
          const group = styleGroups.get(cell.styleIndex);
          if (group) {
            group.push(cell);
          } else {
            styleGroups.set(cell.styleIndex, [cell]);
          }
        }
      }

      if (importCellMetadata) {
        const metadata: ExcelCellMetadata = {};
        if (format?.numberFormat) {
          metadata.numberFormat = format.numberFormat;
        }
        const hyperlink = worksheet.hyperlinks[cell.reference];
        if (hyperlink) {
          metadata.hyperlink = hyperlink;
        }
        if (Object.keys(metadata).length > 0) {
          cellMetadata.set(cell.reference, metadata);
        }
      }
    }

    if (importStyles) {
      for (const [styleIndex, cells] of styleGroups) {
        const format = styleTable.formats[styleIndex];
        if (!format) {
          continue;
        }
        cellStyles.push({
          areas: coalesceCellsIntoAreas(entry.name, cells),
          style: format.style,
        });
      }
    }

    if (importConditionalFormatting) {
      // Excel evaluates the lowest priority number last, so it wins. The engine
      // applies later rules on top, so rules are emitted in reverse priority.
      const ordered = [...worksheet.conditionalFormats].sort(
        (a, b) => b.priority - a.priority
      );
      for (const rule of ordered) {
        const { result, diagnostics: ruleDiagnostics } =
          translateConditionalFormat(rule, {
            sheetName: entry.name,
            differentialFormats: styleTable.differentialFormats,
          });
        diagnostics.push(...ruleDiagnostics);
        if (result) {
          conditionalStyles.push(result);
        }
      }
    }

    if (importTables) {
      for (const table of parseTables(pkg, entry.partPath)) {
        const converted = convertTable(table, entry.name, diagnostics);
        if (converted) {
          tables.push(converted);
        }
      }
    }

    const sheetMetadata: ExcelSheetMetadata = {};
    if (Object.keys(worksheet.columnWidths).length > 0) {
      sheetMetadata.columnWidths = worksheet.columnWidths;
    }
    if (Object.keys(worksheet.rowHeights).length > 0) {
      sheetMetadata.rowHeights = worksheet.rowHeights;
    }
    if (worksheet.merges.length > 0) {
      sheetMetadata.merges = worksheet.merges;
    }
    if (worksheet.frozen) {
      sheetMetadata.frozen = worksheet.frozen;
    }
    if (entry.state !== "visible") {
      sheetMetadata.visible = false;
    }

    sheets.push({
      name: entry.name,
      content,
      cellMetadata: cellMetadata.size > 0 ? cellMetadata : undefined,
      sheetMetadata:
        Object.keys(sheetMetadata).length > 0 ? sheetMetadata : undefined,
    });
  }

  const namedExpressions = importDefinedNames
    ? convertDefinedNames(
        definedNames,
        sheetEntries.map((entry) => entry.name),
        importedSheetNames,
        diagnostics
      )
    : [];

  const data: ExcelWorkbookData = {
    sheets,
    tables: tables.length > 0 ? tables : undefined,
    namedExpressions:
      namedExpressions.length > 0 ? namedExpressions : undefined,
    cellStyles: cellStyles.length > 0 ? cellStyles : undefined,
    conditionalStyles:
      conditionalStyles.length > 0 ? conditionalStyles : undefined,
    workbookMetadata: {
      fileName: options.fileName,
      sheetOrder: sheetEntries.map((entry) => entry.name),
      diagnostics: diagnostics.length > 0 ? diagnostics : undefined,
    },
  };

  return { data, diagnostics };
}

interface CellValueContext {
  convertDates: boolean;
  date1904: boolean;
  sheetName: string;
  diagnostics: Diagnostic[];
}

/**
 * The engine stores a cell as a formula string, a literal, or nothing. A cell
 * with a formula keeps the formula: Excel's cached result is discarded, because
 * the engine recalculates.
 */
function cellValue(
  cell: RawCell,
  format: ExcelFormat | undefined,
  context: CellValueContext
): SerializedCellValue {
  if (cell.formula !== undefined) {
    const body = cell.formulaOffset
      ? offsetFormulaReferences(
          cell.formula,
          cell.formulaOffset.cols,
          cell.formulaOffset.rows
        )
      : cell.formula;

    const translated = translateFormula(body, {
      sheetName: context.sheetName,
      cellReference: cell.reference,
    });
    context.diagnostics.push(...translated.diagnostics);
    return translated.formula;
  }

  switch (cell.kind) {
    case "number": {
      const value = cell.value as number;
      if (context.convertDates && format?.isDateFormat) {
        return excelSerialToIso(value, context.date1904);
      }
      return value;
    }
    case "boolean":
      return cell.value as boolean;
    case "string":
      return cell.value as string;
    case "date":
      return cell.value as string;
    case "error":
      // Excel's own error text is valid engine cell content.
      return cell.value as string;
    case "empty":
      return undefined;
  }
}

function convertTable(
  table: RawTable,
  sheetName: string,
  diagnostics: Diagnostic[]
): WorkbookDataTable | undefined {
  const parsed = parseA1Range(table.ref);
  if (!parsed) {
    diagnostics.push({
      severity: "warning",
      code: "skipped-table",
      message: `Table "${table.displayName}" has an unreadable range (${table.ref}).`,
      sheetName,
    });
    return undefined;
  }

  if (table.headerRowCount === 0) {
    // The engine requires a header row, since structured references name it.
    diagnostics.push({
      severity: "warning",
      code: "skipped-table",
      message:
        `Table "${table.displayName}" has no header row, which the engine ` +
        `requires for structured references, so it was skipped.`,
      sheetName,
    });
    return undefined;
  }

  const numCols = parsed.end.colIndex - parsed.start.colIndex + 1;
  const dataRows =
    parsed.end.rowIndex -
    parsed.start.rowIndex +
    1 -
    table.headerRowCount -
    table.totalsRowCount;

  if (numCols <= 0 || dataRows < 0) {
    diagnostics.push({
      severity: "warning",
      code: "skipped-table",
      message: `Table "${table.displayName}" has an empty range (${table.ref}).`,
      sheetName,
    });
    return undefined;
  }

  return {
    name: table.displayName,
    sheetName,
    start: toA1(parsed.start),
    numRows: { type: "number", value: dataRows },
    numCols,
  };
}

function convertDefinedNames(
  definedNames: { name: string; formula: string; localSheetId?: number; hidden: boolean }[],
  sheetNames: string[],
  importedSheetNames: Set<string>,
  diagnostics: Diagnostic[]
): WorkbookDataNamedExpression[] {
  const result: WorkbookDataNamedExpression[] = [];

  for (const definedName of definedNames) {
    if (BUILTIN_DEFINED_NAMES.has(definedName.name.toLowerCase())) {
      continue;
    }
    if (definedName.hidden || definedName.formula.trim() === "") {
      continue;
    }
    if (definedName.formula.includes("#REF!")) {
      diagnostics.push({
        severity: "warning",
        code: "skipped-defined-name",
        message: `Defined name "${definedName.name}" refers to a deleted range.`,
        source: definedName.formula,
      });
      continue;
    }

    const sheetName =
      definedName.localSheetId === undefined
        ? undefined
        : sheetNames[definedName.localSheetId];

    if (sheetName !== undefined && !importedSheetNames.has(sheetName)) {
      continue;
    }

    // Named expressions are stored as a formula body with no leading "=".
    const translated = translateFormula(definedName.formula);
    diagnostics.push(...translated.diagnostics);

    result.push({
      name: definedName.name,
      expression: translated.formula.slice(1),
      sheetName,
    });
  }

  return result;
}

/**
 * Turn a set of cells sharing one style into as few rectangular areas as
 * possible.
 *
 * Excel's style model is per-cell, but the engine's is per-area, so emitting
 * one area per cell would produce enormous rule lists on real workbooks. Cells
 * are merged into horizontal runs, then adjacent identical runs are merged
 * vertically into blocks.
 */
export function coalesceCellsIntoAreas(
  sheetName: string,
  cells: { colIndex: number; rowIndex: number }[]
): { sheetName: string; range: ReturnType<typeof toSpreadsheetRange> }[] {
  const byRow = new Map<number, number[]>();
  for (const cell of cells) {
    const row = byRow.get(cell.rowIndex);
    if (row) {
      row.push(cell.colIndex);
    } else {
      byRow.set(cell.rowIndex, [cell.colIndex]);
    }
  }

  interface Run {
    rowIndex: number;
    startCol: number;
    endCol: number;
  }

  const runs: Run[] = [];
  for (const [rowIndex, columns] of byRow) {
    columns.sort((a, b) => a - b);
    let startCol = columns[0]!;
    let previous = startCol;
    for (let i = 1; i < columns.length; i++) {
      const current = columns[i]!;
      if (current === previous + 1) {
        previous = current;
        continue;
      }
      runs.push({ rowIndex, startCol, endCol: previous });
      startCol = current;
      previous = current;
    }
    runs.push({ rowIndex, startCol, endCol: previous });
  }

  // Merge runs with identical column spans in consecutive rows into blocks.
  runs.sort((a, b) =>
    a.startCol - b.startCol || a.endCol - b.endCol || a.rowIndex - b.rowIndex
  );

  const areas: { sheetName: string; range: ReturnType<typeof toSpreadsheetRange> }[] = [];
  let index = 0;
  while (index < runs.length) {
    const first = runs[index]!;
    let lastRow = first.rowIndex;
    let next = index + 1;
    while (
      next < runs.length &&
      runs[next]!.startCol === first.startCol &&
      runs[next]!.endCol === first.endCol &&
      runs[next]!.rowIndex === lastRow + 1
    ) {
      lastRow = runs[next]!.rowIndex;
      next++;
    }

    areas.push({
      sheetName,
      range: toSpreadsheetRange({
        start: { colIndex: first.startCol, rowIndex: first.rowIndex },
        end: { colIndex: first.endCol, rowIndex: lastRow },
      }),
    });
    index = next;
  }

  return areas;
}

export { rangeReferenceToSpreadsheetRange };
