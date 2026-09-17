/**
 * Reads the OPC (zip) container of an .xlsx file and exposes its parts as
 * parsed XML.
 */

import { unzipSync, strFromU8 } from "fflate";
import { XMLParser } from "fast-xml-parser";

export type XmlNode = Record<string, unknown>;

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@",
  // Repeating elements are NOT forced into arrays here: the same tag name can
  // be a container in one part and a repeating element in another (<color> is
  // one of each). Every read site wraps with asArray() instead, which is
  // correct for both shapes.
  // Cell text must survive exactly as stored: Excel writes "007" and " x " and
  // both must round-trip.
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  processEntities: true,
});

export interface XlsxPackage {
  /** Raw part bytes keyed by zip path, e.g. `"xl/worksheets/sheet1.xml"`. */
  readonly files: Map<string, Uint8Array>;
  /** Parse a part as XML, or undefined when the part is absent. */
  xml(path: string): XmlNode | undefined;
  text(path: string): string | undefined;
  has(path: string): boolean;
  /** Every part path beginning with the given prefix. */
  list(prefix: string): string[];
}

export async function readXlsxPackage(
  source: Blob | ArrayBuffer | Uint8Array
): Promise<XlsxPackage> {
  const bytes = await toBytes(source);
  let unzipped: Record<string, Uint8Array>;
  try {
    unzipped = unzipSync(bytes);
  } catch (cause) {
    throw new Error(
      "Could not read the file as an .xlsx package. Legacy .xls files and " +
        "password-protected workbooks are not supported.",
      { cause }
    );
  }

  const files = new Map<string, Uint8Array>();
  for (const [path, content] of Object.entries(unzipped)) {
    // Zip paths are sometimes written with a leading slash or backslashes.
    files.set(normalizePath(path), content);
  }

  if (!files.has("xl/workbook.xml")) {
    throw new Error(
      "The file is a zip archive but does not contain xl/workbook.xml, so it " +
        "is not an Excel workbook."
    );
  }

  const xmlCache = new Map<string, XmlNode | undefined>();

  return {
    files,
    text(path) {
      const content = files.get(normalizePath(path));
      return content ? strFromU8(content) : undefined;
    },
    xml(path) {
      const key = normalizePath(path);
      if (xmlCache.has(key)) {
        return xmlCache.get(key);
      }
      const content = files.get(key);
      const parsed = content ? (parser.parse(strFromU8(content)) as XmlNode) : undefined;
      xmlCache.set(key, parsed);
      return parsed;
    },
    has(path) {
      return files.has(normalizePath(path));
    },
    list(prefix) {
      const normalized = normalizePath(prefix);
      return Array.from(files.keys()).filter((path) =>
        path.startsWith(normalized)
      );
    },
  };
}

function normalizePath(path: string): string {
  return path.replace(/\\/g, "/").replace(/^\/+/, "");
}

async function toBytes(
  source: Blob | ArrayBuffer | Uint8Array
): Promise<Uint8Array> {
  if (source instanceof Uint8Array) {
    return source;
  }
  if (source instanceof ArrayBuffer) {
    return new Uint8Array(source);
  }
  return new Uint8Array(await source.arrayBuffer());
}

/**
 * fast-xml-parser gives a single child as an object and repeated children as an
 * array. Parts we do not force into arrays go through this.
 */
export function asArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined || value === null) {
    return [];
  }
  return Array.isArray(value) ? value : [value];
}

export function attr(node: unknown, name: string): string | undefined {
  if (!node || typeof node !== "object") {
    return undefined;
  }
  const value = (node as Record<string, unknown>)[`@${name}`];
  return value === undefined || value === null ? undefined : String(value);
}

export function numAttr(node: unknown, name: string): number | undefined {
  const raw = attr(node, name);
  if (raw === undefined) {
    return undefined;
  }
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : undefined;
}

export function boolAttr(
  node: unknown,
  name: string,
  fallback = false
): boolean {
  const raw = attr(node, name);
  if (raw === undefined) {
    return fallback;
  }
  return raw === "1" || raw.toLowerCase() === "true";
}

/** Text content of an element parsed by fast-xml-parser. */
export function textOf(node: unknown): string {
  if (node === undefined || node === null) {
    return "";
  }
  if (typeof node === "string") {
    return node;
  }
  if (typeof node === "number" || typeof node === "boolean") {
    return String(node);
  }
  if (typeof node === "object") {
    const value = (node as Record<string, unknown>)["#text"];
    return value === undefined || value === null ? "" : String(value);
  }
  return "";
}
