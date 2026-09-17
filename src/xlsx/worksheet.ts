/**
 * Parses one xl/worksheets/sheetN.xml part.
 *
 * Cell values in Excel are typed by the `t` attribute and stored in `<v>`, with
 * formulas alongside in `<f>`. Shared formulas are stored once on the group's
 * master cell and re-derived here for every other member.
 */

import {
  asArray,
  attr,
  boolAttr,
  numAttr,
  textOf,
  type XlsxPackage,
  type XmlNode,
} from "./package";
import { parseA1, toA1 } from "../a1";

export type CellKind =
  | "number"
  | "string"
  | "boolean"
  | "error"
  | "date"
  | "empty";

export interface RawCell {
  reference: string;
  colIndex: number;
  rowIndex: number;
  kind: CellKind;
  /** Literal value, before date conversion. */
  value: string | number | boolean | undefined;
  /** Excel formula text without the leading `=`, when the cell has one. */
  formula?: string;
  /**
   * How far this cell sits from the master cell of its shared-formula group.
   * The formula's relative references are shifted by this much at translation
   * time. Absent or zero for an ordinary formula.
   */
  formulaOffset?: { cols: number; rows: number };
  /** Index into the style table's `formats`. */
  styleIndex?: number;
}

export interface RawConditionalFormat {
  /** Ranges the rule applies to, as written in the file (`"A1:B10 D1:D5"`). */
  ranges: string[];
  type: string;
  operator?: string;
  /** Rule formulas, in the order Excel stores them. */
  formulas: string[];
  text?: string;
  /** Index into the style table's `differentialFormats`. */
  dxfId?: number;
  priority: number;
  /** Colour-scale stops, for `colorScale` rules. */
  colorScale?: {
    cfvo: { type: string; val?: string }[];
    colors: (string | undefined)[];
  };
  /** Present for `top10` rules. */
  rank?: number;
  bottom?: boolean;
  percent?: boolean;
  stopIfTrue?: boolean;
}

export interface RawWorksheet {
  cells: RawCell[];
  merges: string[];
  columnWidths: Record<number, number>;
  rowHeights: Record<number, number>;
  frozen?: { rows: number; cols: number };
  conditionalFormats: RawConditionalFormat[];
  /** Hyperlink targets keyed by A1 reference. */
  hyperlinks: Record<string, string>;
}

export function parseWorksheet(
  pkg: XlsxPackage,
  partPath: string,
  sharedStrings: string[]
): RawWorksheet {
  const root = pkg.xml(partPath);
  const worksheet = root?.["worksheet"] as XmlNode | undefined;
  const result: RawWorksheet = {
    cells: [],
    merges: [],
    columnWidths: {},
    rowHeights: {},
    conditionalFormats: [],
    hyperlinks: {},
  };
  if (!worksheet) {
    return result;
  }

  parseCells(worksheet, sharedStrings, result);
  parseColumns(worksheet, result);
  parsePanes(worksheet, result);
  parseMerges(worksheet, result);
  parseConditionalFormats(worksheet, result);
  parseHyperlinks(pkg, partPath, worksheet, result);

  return result;
}

/** Master formula of a shared-formula group, keyed by the group's `si`. */
interface SharedFormula {
  formula: string;
  colIndex: number;
  rowIndex: number;
}

interface CellFormula {
  formula: string;
  offset?: { cols: number; rows: number };
}

function parseCells(
  worksheet: XmlNode,
  sharedStrings: string[],
  result: RawWorksheet
): void {
  const sheetData = worksheet["sheetData"] as XmlNode | undefined;
  const shared = new Map<string, SharedFormula>();

  for (const rawRow of asArray(sheetData?.["row"])) {
    const row = rawRow as XmlNode;
    const rowNumber = numAttr(row, "r");
    const height = numAttr(row, "ht");
    if (rowNumber !== undefined && height !== undefined) {
      result.rowHeights[rowNumber - 1] = height;
    }

    for (const rawCell of asArray(row["c"])) {
      const cell = rawCell as XmlNode;
      const reference = attr(cell, "r");
      if (!reference) {
        continue;
      }
      const address = parseA1(reference);
      if (!address) {
        continue;
      }

      const styleIndex = numAttr(cell, "s");
      const type = attr(cell, "t") ?? "n";
      const formula = readFormula(cell, address, shared);
      const { kind, value } = readValue(cell, type, sharedStrings);

      if (kind === "empty" && formula === undefined && styleIndex === undefined) {
        continue;
      }

      result.cells.push({
        reference,
        colIndex: address.colIndex,
        rowIndex: address.rowIndex,
        kind,
        value,
        formula: formula?.formula,
        formulaOffset: formula?.offset,
        styleIndex,
      });
    }
  }
}

function readFormula(
  cell: XmlNode,
  address: { colIndex: number; rowIndex: number },
  shared: Map<string, SharedFormula>
): CellFormula | undefined {
  const node = cell["f"];
  if (node === undefined || node === null) {
    return undefined;
  }

  const body = textOf(node);
  const type = attr(node, "t");
  const si = attr(node, "si");

  if (type === "shared" && si !== undefined) {
    if (body) {
      // The master cell of the group carries the formula text.
      shared.set(si, {
        formula: body,
        colIndex: address.colIndex,
        rowIndex: address.rowIndex,
      });
      return { formula: body };
    }
    const master = shared.get(si);
    if (!master) {
      return undefined;
    }
    // Members store no text of their own: they reuse the master's formula with
    // its relative references shifted by the distance between the two cells.
    return {
      formula: master.formula,
      offset: {
        cols: address.colIndex - master.colIndex,
        rows: address.rowIndex - master.rowIndex,
      },
    };
  }

  // Array formulas ("array"/"dataTable") carry their text on the anchor cell.
  return body ? { formula: body } : undefined;
}

function readValue(
  cell: XmlNode,
  type: string,
  sharedStrings: string[]
): { kind: CellKind; value: string | number | boolean | undefined } {
  // Inline strings live in <is>, everything else in <v>.
  if (type === "inlineStr") {
    const inline = cell["is"];
    return { kind: "string", value: readRichText(inline) };
  }

  const raw = cell["v"];
  if (raw === undefined || raw === null) {
    return { kind: "empty", value: undefined };
  }
  const text = textOf(raw);

  switch (type) {
    case "s": {
      const index = Number.parseInt(text, 10);
      return { kind: "string", value: sharedStrings[index] ?? "" };
    }
    case "str":
      // A formula that returned text.
      return { kind: "string", value: text };
    case "b":
      return { kind: "boolean", value: text === "1" };
    case "e":
      return { kind: "error", value: text };
    case "d":
      // ISO 8601 date, used by strict OOXML.
      return { kind: "date", value: text };
    default: {
      const parsed = Number(text);
      return Number.isFinite(parsed)
        ? { kind: "number", value: parsed }
        : { kind: "string", value: text };
    }
  }
}

/** Flatten a rich-text run container into plain text. */
export function readRichText(node: unknown): string {
  if (node === undefined || node === null) {
    return "";
  }
  if (typeof node === "string") {
    return node;
  }
  const record = node as XmlNode;
  const runs = asArray(record["r"]);
  if (runs.length > 0) {
    return runs.map((run) => textOf((run as XmlNode)["t"])).join("");
  }
  if (record["t"] !== undefined) {
    return textOf(record["t"]);
  }
  return textOf(record);
}

function parseColumns(worksheet: XmlNode, result: RawWorksheet): void {
  for (const rawCol of asArray(
    (worksheet["cols"] as XmlNode | undefined)?.["col"]
  )) {
    const col = rawCol as XmlNode;
    const min = numAttr(col, "min");
    const max = numAttr(col, "max");
    const width = numAttr(col, "width");
    if (min === undefined || max === undefined || width === undefined) {
      continue;
    }
    // A run can legitimately span to Excel's last column; recording every index
    // would be wasteful, so very wide runs are skipped.
    if (max - min > 1024) {
      continue;
    }
    for (let index = min; index <= max; index++) {
      result.columnWidths[index - 1] = width;
    }
  }
}

function parsePanes(worksheet: XmlNode, result: RawWorksheet): void {
  const sheetViews = worksheet["sheetViews"] as XmlNode | undefined;
  const sheetView = asArray(sheetViews?.["sheetView"])[0] as XmlNode | undefined;
  const pane = sheetView?.["pane"] as XmlNode | undefined;
  if (!pane) {
    return;
  }
  if (attr(pane, "state") !== "frozen") {
    return;
  }
  result.frozen = {
    cols: numAttr(pane, "xSplit") ?? 0,
    rows: numAttr(pane, "ySplit") ?? 0,
  };
}

function parseMerges(worksheet: XmlNode, result: RawWorksheet): void {
  for (const merge of asArray(
    (worksheet["mergeCells"] as XmlNode | undefined)?.["mergeCell"]
  )) {
    const ref = attr(merge, "ref");
    if (ref) {
      result.merges.push(ref);
    }
  }
}

function parseConditionalFormats(
  worksheet: XmlNode,
  result: RawWorksheet
): void {
  for (const rawBlock of asArray(worksheet["conditionalFormatting"])) {
    const block = rawBlock as XmlNode;
    const sqref = attr(block, "sqref");
    if (!sqref) {
      continue;
    }
    const ranges = sqref.split(/\s+/).filter(Boolean);

    for (const rawRule of asArray(block["cfRule"])) {
      const rule = rawRule as XmlNode;
      const type = attr(rule, "type");
      if (!type) {
        continue;
      }

      const colorScaleNode = rule["colorScale"] as XmlNode | undefined;

      result.conditionalFormats.push({
        ranges,
        type,
        operator: attr(rule, "operator"),
        formulas: asArray(rule["formula"]).map((formula) => textOf(formula)),
        text: attr(rule, "text"),
        dxfId: numAttr(rule, "dxfId"),
        priority: numAttr(rule, "priority") ?? 0,
        rank: numAttr(rule, "rank"),
        bottom: boolAttr(rule, "bottom"),
        percent: boolAttr(rule, "percent"),
        stopIfTrue: boolAttr(rule, "stopIfTrue"),
        colorScale: colorScaleNode
          ? {
              cfvo: asArray(colorScaleNode["cfvo"]).map((cfvo) => ({
                type: attr(cfvo, "type") ?? "min",
                val: attr(cfvo, "val"),
              })),
              colors: asArray(colorScaleNode["color"]).map((color) =>
                attr(color, "rgb")
              ),
            }
          : undefined,
      });
    }
  }
}

function parseHyperlinks(
  pkg: XlsxPackage,
  partPath: string,
  worksheet: XmlNode,
  result: RawWorksheet
): void {
  const links = asArray(
    (worksheet["hyperlinks"] as XmlNode | undefined)?.["hyperlink"]
  );
  if (links.length === 0) {
    return;
  }

  const relationships = readRelationships(pkg, partPath);

  for (const rawLink of links) {
    const link = rawLink as XmlNode;
    const ref = attr(link, "ref");
    if (!ref) {
      continue;
    }
    const id = attr(link, "r:id") ?? attr(link, "id");
    const location = attr(link, "location");
    const target = id ? relationships.get(id) : undefined;
    const href = target
      ? location
        ? `${target}#${location}`
        : target
      : location;
    if (!href) {
      continue;
    }

    // A hyperlink can cover a range; record it for the range's top-left cell.
    const start = parseA1(ref.split(":")[0] ?? ref);
    if (start) {
      result.hyperlinks[toA1(start)] = href;
    }
  }
}

export function readRelationships(
  pkg: XlsxPackage,
  partPath: string
): Map<string, string> {
  const slash = partPath.lastIndexOf("/");
  const dir = slash === -1 ? "" : partPath.slice(0, slash);
  const file = slash === -1 ? partPath : partPath.slice(slash + 1);
  const relsPath = `${dir}/_rels/${file}.rels`;

  const map = new Map<string, string>();
  const root = pkg.xml(relsPath);
  const container = root?.["Relationships"] as XmlNode | undefined;
  if (!container) {
    return map;
  }
  for (const relationship of asArray(container["Relationship"])) {
    const id = attr(relationship, "Id");
    const target = attr(relationship, "Target");
    if (id && target) {
      map.set(id, target);
    }
  }
  return map;
}
