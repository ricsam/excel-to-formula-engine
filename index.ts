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
import type {
  ConvertOptions,
  ConvertResult,
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

export { convertWorkbook };

export type {
  ConvertOptions,
  ConvertResult,
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
