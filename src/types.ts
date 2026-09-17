import type { WorkbookData } from "@ricsam/formula-engine";

/**
 * Severity of a translation note.
 *
 * `info` records a deliberate, lossless-enough choice. `warning` records
 * something the engine cannot represent, where the imported workbook will
 * behave differently from Excel.
 */
export type DiagnosticSeverity = "info" | "warning";

export type DiagnosticCode =
  /** The formula uses a function the engine does not implement. */
  | "unsupported-function"
  /** The formula could not be translated and was kept verbatim. */
  | "untranslatable-formula"
  /** An Excel feature with no engine equivalent was dropped. */
  | "unsupported-feature"
  /** A conditional formatting rule could not be represented. */
  | "unsupported-conditional-format"
  /** A defined name was skipped, usually because it is a built-in. */
  | "skipped-defined-name"
  /** A table could not be created. */
  | "skipped-table"
  /** A sheet was skipped, for example a chart sheet. */
  | "skipped-sheet"
  /** A delimiter had to be guessed from the text. */
  | "guessed-delimiter";

export interface Diagnostic {
  severity: DiagnosticSeverity;
  code: DiagnosticCode;
  message: string;
  sheetName?: string;
  /** A1 reference when the diagnostic belongs to one cell. */
  cellReference?: string;
  /** The Excel source text that produced the diagnostic. */
  source?: string;
}

/**
 * Per-cell metadata the converter attaches to imported cells.
 *
 * The engine stores this verbatim; it is here so a UI can show the original
 * Excel number format or hyperlink.
 */
export interface ExcelCellMetadata {
  /** Excel number format code, for example `"0.00%"`. */
  numberFormat?: string;
  /** Hyperlink target. */
  hyperlink?: string;
  /** Cell comment or note text. */
  note?: string;
}

export interface ExcelSheetMetadata {
  /** Column widths in Excel character units, keyed by 0-based column index. */
  columnWidths?: Record<number, number>;
  /** Row heights in points, keyed by 0-based row index. */
  rowHeights?: Record<number, number>;
  /** Merged cell ranges in A1 notation, for example `"A1:B2"`. */
  merges?: string[];
  /** Frozen pane split, in cells from the top-left. */
  frozen?: { rows: number; cols: number };
  /** False when the sheet is hidden or very hidden in Excel. */
  visible?: boolean;
  /** Sheet tab colour as a hex string. */
  tabColor?: string;
}

export interface ExcelWorkbookMetadata {
  /** Source file name when the Blob carried one. */
  fileName?: string;
  /** Sheet names in their original workbook order. */
  sheetOrder?: string[];
  /** Everything the converter could not represent. */
  diagnostics?: Diagnostic[];
}

/**
 * The converter's output: `WorkbookData` the engine can import directly, with
 * the converter's metadata types applied.
 */
export type ExcelWorkbookData = WorkbookData<
  ExcelCellMetadata,
  ExcelSheetMetadata,
  ExcelWorkbookMetadata,
  unknown
>;

export interface ConvertOptions {
  /**
   * Name recorded in workbook metadata. Purely informational — the workbook's
   * real name is chosen when it is handed to `addWorkbook`.
   */
  fileName?: string;
  /**
   * Import cell fills, fonts, borders and alignment as engine cell styles.
   * @default true
   */
  styles?: boolean;
  /**
   * Import Excel conditional formatting rules the engine can express.
   * @default true
   */
  conditionalFormatting?: boolean;
  /** Import Excel tables as engine tables. @default true */
  tables?: boolean;
  /** Import defined names as named expressions. @default true */
  definedNames?: boolean;
  /**
   * Attach per-cell metadata (number formats, hyperlinks, notes).
   * @default true
   */
  cellMetadata?: boolean;
  /**
   * Convert Excel date/time serial numbers to ISO strings instead of leaving
   * them as the raw serial number. The engine has no date type, so a serial
   * number would otherwise display as e.g. `45000`.
   * @default true
   */
  convertDates?: boolean;
  /**
   * Import only these sheets, by name. Defaults to every worksheet.
   */
  sheets?: string[];
}

/**
 * Options for reading delimited text (CSV, TSV).
 *
 * The parsing defaults match what a spreadsheet does with the same file, on the
 * grounds that the file usually came out of one.
 *
 * There is deliberately no option to keep `=`-leading fields as text. The engine
 * stores a formula as a string beginning with `=` and has no escape for text
 * that merely looks like one, so such an option could only be honoured by
 * altering the value — which is worse than not offering it.
 */
export interface CsvConvertOptions {
  /** Name for the single sheet the file becomes. @default "Sheet1" */
  sheetName?: string;
  /** Recorded in workbook metadata. */
  fileName?: string;
  /**
   * Field separator. Detected from the text when omitted, which is the right
   * choice for `.csv` — the extension says nothing about the separator, and a
   * comma-decimal locale writes semicolons.
   */
  delimiter?: string;
  /**
   * Convert plain decimals to numbers. Leading-zero strings and integers too
   * long to survive as floats are kept as text either way.
   * @default true
   */
  parseNumbers?: boolean;
  /** Convert `TRUE` / `FALSE` (any case) to booleans. @default true */
  parseBooleans?: boolean;
  /**
   * Strip surrounding whitespace from every field.
   *
   * Off by default because the whitespace may be data, and a quoted field that
   * carries it meant to.
   * @default false
   */
  trimFields?: boolean;
}

export interface ConvertResult {
  data: ExcelWorkbookData;
  diagnostics: Diagnostic[];
}
