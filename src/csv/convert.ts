/**
 * Delimited text in, `WorkbookData` out.
 *
 * The interesting decisions here are all about what a bare string *is*. A CSV
 * carries no types, so every cell is a judgement call, and the ones that matter
 * are the cases where guessing wrong destroys data rather than merely looking
 * wrong: a leading-zero part number, an identifier too long for a float.
 */

import type { SerializedCellValue } from "@ricsam/formula-engine";
import { columnIndexToLabel } from "../a1";
import { translateFormula } from "../translate/formula";
import type {
  CsvConvertOptions,
  Diagnostic,
  ExcelWorkbookData,
} from "../types";
import { detectDelimiter, parseDelimitedText, stripBom } from "./parse";

export interface ConvertCsvResult {
  data: ExcelWorkbookData;
  diagnostics: Diagnostic[];
}

/** Cells beyond this are refused rather than silently locking up the tab. */
const MAX_CELLS = 2_000_000;

export function convertDelimitedText(
  text: string,
  options: CsvConvertOptions = {}
): ConvertCsvResult {
  const {
    sheetName = "Sheet1",
    fileName,
    parseNumbers = true,
    parseBooleans = true,
    trimFields = false,
  } = options;

  const diagnostics: Diagnostic[] = [];
  const body = stripBom(text);
  const delimiter = options.delimiter ?? detectDelimiter(body);
  const rows = parseDelimitedText(body, delimiter);

  const widest = rows.reduce((max, row) => Math.max(max, row.length), 0);
  if (rows.length * widest > MAX_CELLS) {
    throw new Error(
      `The file has about ${rows.length * widest} cells, more than the ` +
        `${MAX_CELLS} this importer will load at once.`
    );
  }

  const content = new Map<string, SerializedCellValue>();

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex]!;
    for (let colIndex = 0; colIndex < row.length; colIndex++) {
      const raw = trimFields ? row[colIndex]!.trim() : row[colIndex]!;
      // An empty cell is absent rather than an empty string: the engine deletes
      // empty values anyway, and a sparse sheet is the point of its storage.
      if (raw.length === 0) continue;

      const reference = `${columnIndexToLabel(colIndex)}${rowIndex + 1}`;
      content.set(
        reference,
        toCellValue(raw, {
          parseNumbers,
          parseBooleans,
          sheetName,
          reference,
          diagnostics,
        })
      );
    }
  }

  const data: ExcelWorkbookData = {
    sheets: [{ name: sheetName, content }],
    workbookMetadata: {
      fileName,
      sheetOrder: [sheetName],
      diagnostics: diagnostics.length > 0 ? diagnostics : undefined,
    },
  };

  return { data, diagnostics };
}

interface ToCellValueContext {
  parseNumbers: boolean;
  parseBooleans: boolean;
  sheetName: string;
  reference: string;
  diagnostics: Diagnostic[];
}

function toCellValue(raw: string, context: ToCellValueContext): SerializedCellValue {
  // A field beginning with `=` is a formula, the way it is when a spreadsheet
  // opens the same file. See the note on CsvConvertOptions.
  if (raw.startsWith("=") && raw.length > 1) {
    const translated = translateFormula(raw, {
      sheetName: context.sheetName,
      cellReference: context.reference,
    });
    context.diagnostics.push(...translated.diagnostics);
    return translated.formula;
  }

  if (context.parseBooleans) {
    const upper = raw.toUpperCase();
    if (upper === "TRUE") return true;
    if (upper === "FALSE") return false;
  }

  if (context.parseNumbers) {
    const number = toNumber(raw);
    if (number !== undefined) return number;
  }

  return raw;
}

/**
 * The numeric value of a field, or undefined to keep it as text.
 *
 * Deliberately narrower than `Number()`. Three kinds of string look numeric and
 * must not be converted, because converting them loses information that cannot
 * be recovered from the result:
 *
 * - Leading zeros (`007`, `01234`) — part numbers, postcodes, phone extensions.
 *   This is the mangling Excel is notorious for.
 * - Integers beyond `Number.MAX_SAFE_INTEGER` — long identifiers, which silently
 *   change value when they become floats.
 * - Anything `Number()` accepts that is not a plain decimal: `0x1f`, `Infinity`,
 *   `1_000`, and the empty string.
 */
export function toNumber(raw: string): number | undefined {
  const text = raw.trim();
  if (text.length === 0) return undefined;

  // A plain decimal, with an optional sign and exponent. Nothing else.
  if (!/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(text)) {
    return undefined;
  }

  const digits = text.replace(/^[+-]/, "");
  // `0`, `0.5` and `0e3` are fine; `007` and `01.5` are identifiers.
  if (/^0\d/.test(digits)) return undefined;

  const value = Number(text);
  if (!Number.isFinite(value)) return undefined;

  // A long integer that cannot survive as a float stays text.
  if (
    !text.includes(".") &&
    !/[eE]/.test(text) &&
    !Number.isSafeInteger(value)
  ) {
    return undefined;
  }

  return value;
}

/** Decode bytes as UTF-8, which is what every modern exporter writes. */
export async function readTextFrom(
  source: Blob | ArrayBuffer | Uint8Array | string
): Promise<string> {
  if (typeof source === "string") return source;
  if (source instanceof Uint8Array || source instanceof ArrayBuffer) {
    return new TextDecoder("utf-8").decode(source);
  }
  return source.text();
}
