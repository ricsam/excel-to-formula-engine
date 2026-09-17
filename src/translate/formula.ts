/**
 * Translates an Excel formula into engine formula text.
 *
 * Excel and the engine share almost all of their syntax — A1 references, sheet
 * and workbook qualifiers, structured table references, array literals and the
 * operator set all carry over unchanged. Translation therefore rewrites the few
 * constructs that genuinely differ and reports, rather than rewrites, anything
 * the engine cannot evaluate. Formulas using unsupported functions are kept
 * verbatim so the original intent stays visible in the cell.
 */

import { getFormulaFunctionCatalog } from "@ricsam/formula-engine";
import type { Diagnostic } from "../types";

/** Uppercase names of every function the engine implements, including aliases. */
let supportedFunctionNames: Set<string> | undefined;

function getSupportedFunctions(): Set<string> {
  if (!supportedFunctionNames) {
    const names = new Set<string>();
    for (const descriptor of getFormulaFunctionCatalog()) {
      names.add(descriptor.name.toUpperCase());
      for (const alias of descriptor.aliases ?? []) {
        names.add(alias.toUpperCase());
      }
    }
    supportedFunctionNames = names;
  }
  return supportedFunctionNames;
}

/**
 * Excel writes newer functions with a `_xlfn.` prefix (and `_xlfn._xlws.` for
 * worksheet-scoped ones) when they postdate the file format. The prefix is a
 * storage detail, never part of the user-visible formula.
 */
const XLFN_PREFIX = /_xlfn\.(_xlws\.)?/g;

/**
 * `@` marks implicit intersection in files written by dynamic-array Excel. The
 * engine has no implicit-intersection operator, so it is dropped: for a single
 * -valued reference the behaviour is identical.
 */
const IMPLICIT_INTERSECTION = /(?<![\w$"'\]])@(?=[A-Za-z_\\])/g;

export interface TranslateFormulaOptions {
  sheetName?: string;
  cellReference?: string;
}

export interface TranslatedFormula {
  /** Engine formula text, including the leading `=`. */
  formula: string;
  diagnostics: Diagnostic[];
}

export function translateFormula(
  rawFormula: string,
  options: TranslateFormulaOptions = {}
): TranslatedFormula {
  const diagnostics: Diagnostic[] = [];
  let body = rawFormula.startsWith("=") ? rawFormula.slice(1) : rawFormula;

  body = body.replace(XLFN_PREFIX, "");
  body = body.replace(IMPLICIT_INTERSECTION, "");

  for (const name of findUnsupportedFunctions(body)) {
    diagnostics.push({
      severity: "warning",
      code: "unsupported-function",
      message:
        `The engine does not implement ${name}(). The formula was imported ` +
        `unchanged and will evaluate to an error.`,
      sheetName: options.sheetName,
      cellReference: options.cellReference,
      source: rawFormula,
    });
  }

  return { formula: `=${body}`, diagnostics };
}

/**
 * Find function calls in a formula whose names the engine does not implement.
 *
 * This scans outside of string literals and skips identifiers that are not
 * followed by `(`, so table columns (`Table[Amount]`) and named expressions are
 * not mistaken for calls.
 */
export function findUnsupportedFunctions(formulaBody: string): string[] {
  const supported = getSupportedFunctions();
  const found = new Set<string>();

  let inString = false;
  let identifier = "";

  for (let i = 0; i < formulaBody.length; i++) {
    const char = formulaBody[i]!;

    if (inString) {
      if (char === '"') {
        // A doubled quote is an escaped quote inside the literal.
        if (formulaBody[i + 1] === '"') {
          i++;
        } else {
          inString = false;
        }
      }
      continue;
    }

    if (char === '"') {
      inString = true;
      identifier = "";
      continue;
    }

    if (/[A-Za-z0-9_.]/.test(char)) {
      identifier += char;
      continue;
    }

    if (char === "(" && identifier.length > 0) {
      const name = identifier.toUpperCase();
      // A trailing sheet qualifier (Sheet1!SUM) cannot occur, but a decimal
      // number can end up in `identifier`; require a letter first.
      if (/^[A-Z_][A-Z0-9_.]*$/.test(name) && !supported.has(name)) {
        found.add(name);
      }
    }

    identifier = "";
  }

  return Array.from(found);
}

/**
 * Expand a shared formula.
 *
 * Excel stores one formula for a group of cells and expects readers to offset
 * its relative references for every other cell in the group. Absolute parts
 * (`$A$1`) stay put.
 */
export function offsetFormulaReferences(
  formulaBody: string,
  deltaCols: number,
  deltaRows: number
): string {
  if (deltaCols === 0 && deltaRows === 0) {
    return formulaBody;
  }

  // Matches an optional sheet/workbook qualifier followed by a cell reference,
  // so that the qualifier is preserved while the reference is shifted.
  const reference =
    /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![\w(])/g;

  let result = "";
  let lastIndex = 0;
  let inString = false;

  for (let i = 0; i < formulaBody.length; i++) {
    if (formulaBody[i] === '"') {
      if (inString && formulaBody[i + 1] === '"') {
        i++;
        continue;
      }
      inString = !inString;
    }
    if (inString) {
      continue;
    }

    reference.lastIndex = i;
    const match = reference.exec(formulaBody);
    if (!match || match.index !== i) {
      continue;
    }

    const [whole, colAbs, colLabel, rowAbs, rowDigits] = match;
    // Skip anything that is part of a longer identifier, e.g. the "A1" in
    // "MYA1NAME" or a function name. A "!" before the reference is a sheet
    // qualifier, whose reference still shifts, and a "$" is the reference's own
    // absolute marker.
    const before = i > 0 ? formulaBody[i - 1]! : "";
    if (/[A-Za-z0-9_.]/.test(before)) {
      continue;
    }

    result += formulaBody.slice(lastIndex, i);
    result += shiftReference(
      colAbs!,
      colLabel!,
      rowAbs!,
      rowDigits!,
      deltaCols,
      deltaRows
    );
    lastIndex = i + whole!.length;
    i = lastIndex - 1;
  }

  result += formulaBody.slice(lastIndex);
  return result;
}

function shiftReference(
  colAbs: string,
  colLabel: string,
  rowAbs: string,
  rowDigits: string,
  deltaCols: number,
  deltaRows: number
): string {
  const colIndex = columnLabelToIndex(colLabel);
  const rowIndex = Number.parseInt(rowDigits, 10) - 1;

  const newColIndex = colAbs ? colIndex : colIndex + deltaCols;
  const newRowIndex = rowAbs ? rowIndex : rowIndex + deltaRows;

  if (newColIndex < 0 || newRowIndex < 0) {
    return "#REF!";
  }

  return `${colAbs}${columnIndexToLabel(newColIndex)}${rowAbs}${newRowIndex + 1}`;
}

function columnLabelToIndex(label: string): number {
  let index = 0;
  for (let i = 0; i < label.length; i++) {
    index = index * 26 + (label.toUpperCase().charCodeAt(i) - 64);
  }
  return index - 1;
}

function columnIndexToLabel(index: number): string {
  let label = "";
  let remaining = index + 1;
  while (remaining > 0) {
    const rest = (remaining - 1) % 26;
    label = String.fromCharCode(65 + rest) + label;
    remaining = Math.floor((remaining - rest) / 26);
  }
  return label;
}
