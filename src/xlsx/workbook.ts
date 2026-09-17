/**
 * Parses the workbook-level parts: the sheet list, defined names, the shared
 * string table, and table definitions.
 */

import {
  asArray,
  attr,
  numAttr,
  type XlsxPackage,
  type XmlNode,
} from "./package";
import { readRelationships, readRichText } from "./worksheet";

export interface SheetEntry {
  name: string;
  /** Zip path of the worksheet part. */
  partPath: string;
  /** `visible`, `hidden` or `veryHidden`. */
  state: string;
}

export interface DefinedName {
  name: string;
  /** Formula text as stored, e.g. `"Sheet1!$A$1:$B$5"`. */
  formula: string;
  /**
   * Index of the sheet this name is scoped to, or undefined for a
   * workbook-scoped name.
   */
  localSheetId?: number;
  hidden: boolean;
}

export interface RawTable {
  name: string;
  displayName: string;
  /** Range in A1 notation, including the header and totals rows. */
  ref: string;
  headerRowCount: number;
  totalsRowCount: number;
  columnNames: string[];
}

/** Read the shared string table, flattening rich text to plain text. */
export function parseSharedStrings(pkg: XlsxPackage): string[] {
  const root = pkg.xml("xl/sharedStrings.xml");
  const sst = root?.["sst"] as XmlNode | undefined;
  if (!sst) {
    return [];
  }
  return asArray(sst["si"]).map((entry) => readRichText(entry));
}

export function parseWorkbook(pkg: XlsxPackage): {
  sheets: SheetEntry[];
  definedNames: DefinedName[];
  /** True when the file uses the 1904 date system rather than 1900. */
  date1904: boolean;
} {
  const root = pkg.xml("xl/workbook.xml");
  const workbook = root?.["workbook"] as XmlNode | undefined;
  if (!workbook) {
    return { sheets: [], definedNames: [], date1904: false };
  }

  const relationships = readRelationships(pkg, "xl/workbook.xml");

  const sheets: SheetEntry[] = [];
  for (const rawSheet of asArray(
    (workbook["sheets"] as XmlNode | undefined)?.["sheet"]
  )) {
    const sheet = rawSheet as XmlNode;
    const name = attr(sheet, "name");
    const id = attr(sheet, "r:id") ?? attr(sheet, "id");
    if (!name || !id) {
      continue;
    }
    const target = relationships.get(id);
    if (!target) {
      continue;
    }
    sheets.push({
      name,
      partPath: resolvePart(target),
      state: attr(sheet, "state") ?? "visible",
    });
  }

  const definedNames: DefinedName[] = [];
  for (const rawName of asArray(
    (workbook["definedNames"] as XmlNode | undefined)?.["definedName"]
  )) {
    const node = rawName as XmlNode;
    const name = attr(node, "name");
    if (!name) {
      continue;
    }
    definedNames.push({
      name,
      formula: textContent(node),
      localSheetId: numAttr(node, "localSheetId"),
      hidden: attr(node, "hidden") === "1",
    });
  }

  const pr = workbook["workbookPr"];
  const date1904 =
    attr(pr, "date1904") === "1" || attr(pr, "date1904") === "true";

  return { sheets, definedNames, date1904 };
}

/** Table parts referenced by one worksheet. */
export function parseTables(
  pkg: XlsxPackage,
  worksheetPartPath: string
): RawTable[] {
  const relationships = readRelationships(pkg, worksheetPartPath);
  const tables: RawTable[] = [];

  for (const target of relationships.values()) {
    if (!target.includes("tables/")) {
      continue;
    }
    const partPath = resolvePart(target, worksheetPartPath);
    const root = pkg.xml(partPath);
    const table = root?.["table"] as XmlNode | undefined;
    if (!table) {
      continue;
    }

    const ref = attr(table, "ref");
    const name = attr(table, "name") ?? attr(table, "displayName");
    if (!ref || !name) {
      continue;
    }

    tables.push({
      name,
      displayName: attr(table, "displayName") ?? name,
      ref,
      // Absent headerRowCount means one header row, per the schema default.
      headerRowCount: numAttr(table, "headerRowCount") ?? 1,
      totalsRowCount: numAttr(table, "totalsRowCount") ?? 0,
      columnNames: asArray(
        (table["tableColumns"] as XmlNode | undefined)?.["tableColumn"]
      ).map((column) => attr(column, "name") ?? ""),
    });
  }

  return tables;
}

function textContent(node: XmlNode): string {
  const value = node["#text"];
  if (value === undefined || value === null) {
    return "";
  }
  return String(value);
}

/**
 * Resolve a relationship target against the part that declared it.
 *
 * Targets are usually relative (`worksheets/sheet1.xml` from `xl/workbook.xml`)
 * but may be absolute (`/xl/worksheets/sheet1.xml`).
 */
function resolvePart(target: string, fromPart = "xl/workbook.xml"): string {
  if (target.startsWith("/")) {
    return target.slice(1);
  }

  const slash = fromPart.lastIndexOf("/");
  const baseDir = slash === -1 ? "" : fromPart.slice(0, slash);
  const segments = baseDir ? baseDir.split("/") : [];

  for (const segment of target.split("/")) {
    if (segment === "." || segment === "") {
      continue;
    }
    if (segment === "..") {
      segments.pop();
      continue;
    }
    segments.push(segment);
  }

  return segments.join("/");
}
