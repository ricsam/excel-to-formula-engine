/**
 * excel-to-formula-engine
 *
 * Converts an Excel .xlsx file into `WorkbookData` for `@ricsam/formula-engine`,
 * in the browser and in Node/Bun.
 *
 * ```ts
 * const data = await excelToFormulaEngine(blob);
 * engine.addWorkbook({ workbookName: "Budget", data });
 * ```
 */

import { convertWorkbook } from "./src/convert";
import { convertDelimitedText, readTextFrom } from "./src/csv/convert";
import type {
  ConvertOptions,
  ConvertResult,
  CsvConvertOptions,
  ExcelWorkbookData,
} from "./src/types";

/**
 * Convert an Excel file into engine `WorkbookData`.
 *
 * Accepts a `Blob` or `File` from a file input, or raw bytes. The result is
 * ready to hand to `FormulaEngine.addWorkbook`.
 *
 * Anything that could not be translated is reported in
 * `data.workbookMetadata.diagnostics`; use {@link excelToFormulaEngineDetailed}
 * when you want the diagnostics alongside the data instead.
 *
 * @throws when the input is not a readable .xlsx package.
 */
export async function excelToFormulaEngine(
  source: Blob | ArrayBuffer | Uint8Array,
  options?: ConvertOptions
): Promise<ExcelWorkbookData> {
  const { data } = await convertWorkbook(source, withFileName(source, options));
  return data;
}

/**
 * Convert an Excel file, returning the data and the translation diagnostics as
 * separate values.
 */
export async function excelToFormulaEngineDetailed(
  source: Blob | ArrayBuffer | Uint8Array,
  options?: ConvertOptions
): Promise<ConvertResult> {
  return convertWorkbook(source, withFileName(source, options));
}

/** Pick up the name from a `File` so it lands in workbook metadata. */
function withFileName(
  source: Blob | ArrayBuffer | Uint8Array,
  options: ConvertOptions | undefined
): ConvertOptions {
  if (options?.fileName !== undefined) {
    return options;
  }
  const name =
    typeof File !== "undefined" && source instanceof File
      ? source.name
      : undefined;
  return name ? { ...options, fileName: name } : { ...options };
}

/**
 * Convert delimited text — CSV, TSV, or semicolon-separated — into engine
 * `WorkbookData` holding a single sheet.
 *
 * The delimiter is detected from the text unless one is given: a `.csv`
 * extension says nothing about the separator, and a comma-decimal locale
 * exports semicolons.
 */
export async function csvToFormulaEngine(
  source: Blob | ArrayBuffer | Uint8Array | string,
  options?: CsvConvertOptions
): Promise<ExcelWorkbookData> {
  const { data } = await csvToFormulaEngineDetailed(source, options);
  return data;
}

/** As {@link csvToFormulaEngine}, with the diagnostics returned separately. */
export async function csvToFormulaEngineDetailed(
  source: Blob | ArrayBuffer | Uint8Array | string,
  options?: CsvConvertOptions
): Promise<ConvertResult> {
  const text = await readTextFrom(source);
  const resolved = withCsvFileName(source, options);
  return convertDelimitedText(text, resolved);
}

/**
 * Convert whichever kind of spreadsheet file this is.
 *
 * For a drop target, which is handed whatever the user let go of and has to
 * decide afterwards. An `.xlsx` is a ZIP and delimited text is not, so the
 * format is settled by looking at the bytes rather than by trusting the name —
 * a file saved as `.csv` but actually a workbook is a common enough export
 * mistake to be worth handling quietly.
 */
export async function spreadsheetToFormulaEngine(
  source: Blob | ArrayBuffer | Uint8Array,
  options?: ConvertOptions & CsvConvertOptions
): Promise<ExcelWorkbookData> {
  const { data } = await spreadsheetToFormulaEngineDetailed(source, options);
  return data;
}

/** As {@link spreadsheetToFormulaEngine}, with the diagnostics separately. */
export async function spreadsheetToFormulaEngineDetailed(
  source: Blob | ArrayBuffer | Uint8Array,
  options?: ConvertOptions & CsvConvertOptions
): Promise<ConvertResult> {
  const bytes = await toBytes(source);
  if (looksLikeZip(bytes)) {
    return convertWorkbook(bytes, withFileName(source, options));
  }
  return csvToFormulaEngineDetailed(bytes, withCsvFileName(source, options));
}

/** Every .xlsx is a ZIP, and every ZIP starts `PK\x03\x04`. */
function looksLikeZip(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 4 &&
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    bytes[2] === 0x03 &&
    bytes[3] === 0x04
  );
}

async function toBytes(
  source: Blob | ArrayBuffer | Uint8Array
): Promise<Uint8Array> {
  if (source instanceof Uint8Array) return source;
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  return new Uint8Array(await source.arrayBuffer());
}

/**
 * Default the sheet name to the file's own name, so a dropped `sales.csv`
 * becomes a sheet called `sales` rather than `Sheet1`.
 */
function withCsvFileName(
  source: Blob | ArrayBuffer | Uint8Array | string,
  options: CsvConvertOptions | undefined
): CsvConvertOptions {
  const name =
    typeof File !== "undefined" && source instanceof File ? source.name : undefined;
  const fileName = options?.fileName ?? name;
  if (!fileName) return { ...options };
  return {
    ...options,
    fileName,
    sheetName:
      options?.sheetName ??
      (fileName.replace(/\.(csv|tsv|tab|txt)$/i, "").trim() || "Sheet1")
  };
}

export { convertWorkbook };
export { convertDelimitedText } from "./src/csv/convert";
export {
  detectDelimiter,
  parseDelimitedText,
  stripBom,
} from "./src/csv/parse";

export type {
  ConvertOptions,
  ConvertResult,
  CsvConvertOptions,
  Diagnostic,
  DiagnosticCode,
  DiagnosticSeverity,
  ExcelCellMetadata,
  ExcelSheetMetadata,
  ExcelWorkbookData,
  ExcelWorkbookMetadata,
} from "./src/types";

export {
  findUnsupportedFunctions,
  offsetFormulaReferences,
  translateFormula,
} from "./src/translate/formula";

export { excelSerialToIso } from "./src/translate/dates";

export {
  columnIndexToLabel,
  columnLabelToIndex,
  parseA1,
  parseA1Range,
  toA1,
} from "./src/a1";
