/**
 * A1-notation helpers. These mirror the engine's own addressing (0-based row
 * and column indices) so translated references line up exactly.
 */

import type { SpreadsheetRange } from "@ricsam/formula-engine";

const A1_CELL = /^\$?([A-Za-z]+)\$?(\d+)$/;

export function columnLabelToIndex(label: string): number {
  let index = 0;
  for (let i = 0; i < label.length; i++) {
    index = index * 26 + (label.toUpperCase().charCodeAt(i) - 64);
  }
  return index - 1;
}

export function columnIndexToLabel(index: number): string {
  let label = "";
  let remaining = index + 1;
  while (remaining > 0) {
    const rest = (remaining - 1) % 26;
    label = String.fromCharCode(65 + rest) + label;
    remaining = Math.floor((remaining - rest) / 26);
  }
  return label;
}

export interface ParsedCell {
  colIndex: number;
  rowIndex: number;
}

export function parseA1(reference: string): ParsedCell | undefined {
  const match = A1_CELL.exec(reference.trim());
  if (!match || match[1] === undefined || match[2] === undefined) {
    return undefined;
  }
  return {
    colIndex: columnLabelToIndex(match[1]),
    rowIndex: Number.parseInt(match[2], 10) - 1,
  };
}

export function toA1({ colIndex, rowIndex }: ParsedCell): string {
  return `${columnIndexToLabel(colIndex)}${rowIndex + 1}`;
}

export interface ParsedRange {
  start: ParsedCell;
  end: ParsedCell;
}

/**
 * Parse an Excel range such as `A1:B10`, `A1`, `A:C` or `2:5`.
 *
 * Whole-column and whole-row ranges are clamped to Excel's grid limits rather
 * than being left open-ended, because Excel's own semantics are bounded.
 */
export const EXCEL_MAX_ROW_INDEX = 1_048_575;
export const EXCEL_MAX_COL_INDEX = 16_383;

export function parseA1Range(reference: string): ParsedRange | undefined {
  const trimmed = reference.trim().replace(/\$/g, "");
  const [rawStart, rawEnd] = trimmed.split(":");
  if (rawStart === undefined) {
    return undefined;
  }
  if (rawEnd === undefined) {
    const cell = parseA1(rawStart);
    return cell ? { start: cell, end: cell } : undefined;
  }

  const startCell = parseA1(rawStart);
  const endCell = parseA1(rawEnd);
  if (startCell && endCell) {
    return { start: startCell, end: endCell };
  }

  // Whole columns, e.g. "A:C".
  if (/^[A-Za-z]+$/.test(rawStart) && /^[A-Za-z]+$/.test(rawEnd)) {
    return {
      start: { colIndex: columnLabelToIndex(rawStart), rowIndex: 0 },
      end: {
        colIndex: columnLabelToIndex(rawEnd),
        rowIndex: EXCEL_MAX_ROW_INDEX,
      },
    };
  }

  // Whole rows, e.g. "2:5".
  if (/^\d+$/.test(rawStart) && /^\d+$/.test(rawEnd)) {
    return {
      start: { colIndex: 0, rowIndex: Number.parseInt(rawStart, 10) - 1 },
      end: {
        colIndex: EXCEL_MAX_COL_INDEX,
        rowIndex: Number.parseInt(rawEnd, 10) - 1,
      },
    };
  }

  return undefined;
}

export function toSpreadsheetRange(range: ParsedRange): SpreadsheetRange {
  return {
    start: { col: range.start.colIndex, row: range.start.rowIndex },
    end: {
      col: { type: "number", value: range.end.colIndex },
      row: { type: "number", value: range.end.rowIndex },
    },
  };
}

/** Parse an Excel range reference straight into an engine range. */
export function rangeReferenceToSpreadsheetRange(
  reference: string
): SpreadsheetRange | undefined {
  const parsed = parseA1Range(reference);
  return parsed ? toSpreadsheetRange(parsed) : undefined;
}
