import { describe, expect, test } from "bun:test";
import { FormulaEngine, parseCellReference } from "@ricsam/formula-engine";
import type { CellAddress } from "@ricsam/formula-engine";
import {
  csvToFormulaEngine,
  csvToFormulaEngineDetailed,
  detectDelimiter,
  parseDelimitedText,
  spreadsheetToFormulaEngine,
  stripBom,
} from "../../index";
import { toNumber } from "../../src/csv/convert";

/**
 * Delimited text has no types, so every cell is a guess. The tests that matter
 * are the ones where guessing wrong destroys data rather than merely looking
 * wrong — a part number that loses its leading zeros cannot be got back.
 */

const WORKBOOK = "Imported";

function at(sheetName: string, reference: string): CellAddress {
  const { rowIndex, colIndex } = parseCellReference(reference);
  return { workbookName: WORKBOOK, sheetName, rowIndex, colIndex };
}

async function cells(
  text: string,
  options?: Parameters<typeof csvToFormulaEngine>[1]
): Promise<Map<string, unknown>> {
  const data = await csvToFormulaEngine(text, options);
  return data.sheets[0]!.content as Map<string, unknown>;
}

describe("parseDelimitedText", () => {
  test("splits plain rows", () => {
    expect(parseDelimitedText("a,b\nc,d", ",")).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  test("keeps a delimiter inside a quoted field", () => {
    expect(parseDelimitedText('"Smith, J",42', ",")).toEqual([["Smith, J", "42"]]);
  });

  test("reads a doubled quote as one literal quote", () => {
    expect(parseDelimitedText('"she said ""hi""",1', ",")).toEqual([
      ['she said "hi"', "1"],
    ]);
  });

  test("keeps a line break inside a quoted field", () => {
    expect(parseDelimitedText('"line one\nline two",x', ",")).toEqual([
      ["line one\nline two", "x"],
    ]);
  });

  test("handles CRLF, LF and bare CR line endings alike", () => {
    expect(parseDelimitedText("a,b\r\nc,d", ",")).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
    expect(parseDelimitedText("a,b\rc,d", ",")).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  test("does not invent a row for a trailing newline", () => {
    expect(parseDelimitedText("a,b\n", ",")).toEqual([["a", "b"]]);
  });

  test("keeps empty fields in position", () => {
    expect(parseDelimitedText("a,,c", ",")).toEqual([["a", "", "c"]]);
  });

  test("keeps ragged rows ragged", () => {
    expect(parseDelimitedText("a,b,c\nd", ",")).toEqual([["a", "b", "c"], ["d"]]);
  });

  test("splits on tabs when asked to", () => {
    expect(parseDelimitedText("a\tb\nc\td", "\t")).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });
});

describe("detectDelimiter", () => {
  test("finds tabs in a TSV", () => {
    expect(detectDelimiter("a\tb\tc\n1\t2\t3")).toBe("\t");
  });

  test("finds commas in a CSV", () => {
    expect(detectDelimiter("a,b,c\n1,2,3")).toBe(",");
  });

  test("finds the semicolons a comma-decimal locale writes", () => {
    expect(detectDelimiter("a;b;c\n1,5;2,5;3,5")).toBe(";");
  });

  test("ignores a delimiter that only appears inside quotes", () => {
    // The comma lives inside a quoted name; the tab is the real separator.
    expect(detectDelimiter('"Smith, J"\t42\n"Jones, A"\t43')).toBe("\t");
  });

  test("falls back to comma for a single column", () => {
    expect(detectDelimiter("alpha\nbeta\ngamma")).toBe(",");
  });

  test("falls back to comma for empty text", () => {
    expect(detectDelimiter("")).toBe(",");
  });
});

describe("stripBom", () => {
  test("removes the BOM Excel writes on every CSV export", () => {
    // Left in, it becomes an invisible part of the first header cell.
    expect(stripBom("﻿Name,Value")).toBe("Name,Value");
  });

  test("leaves text without one alone", () => {
    expect(stripBom("Name,Value")).toBe("Name,Value");
  });
});

describe("toNumber", () => {
  test("converts plain decimals", () => {
    expect(toNumber("42")).toBe(42);
    expect(toNumber("-3.5")).toBe(-3.5);
    expect(toNumber("1.50")).toBe(1.5);
    expect(toNumber("0")).toBe(0);
    expect(toNumber("0.5")).toBe(0.5);
    expect(toNumber("1e3")).toBe(1000);
  });

  test("keeps leading-zero strings as text", () => {
    // Part numbers and postcodes — the mangling Excel is notorious for.
    expect(toNumber("007")).toBeUndefined();
    expect(toNumber("01234")).toBeUndefined();
  });

  test("keeps integers too long to survive as floats", () => {
    expect(toNumber("12345678901234567890")).toBeUndefined();
  });

  test("refuses things Number() would accept but a spreadsheet would not", () => {
    expect(toNumber("0x1f")).toBeUndefined();
    expect(toNumber("Infinity")).toBeUndefined();
    expect(toNumber("1_000")).toBeUndefined();
    expect(toNumber("")).toBeUndefined();
    expect(toNumber("  ")).toBeUndefined();
  });

  test("refuses formatted numbers, which would lose their formatting", () => {
    expect(toNumber("1,000")).toBeUndefined();
    expect(toNumber("$5")).toBeUndefined();
    expect(toNumber("50%")).toBeUndefined();
  });
});

describe("csvToFormulaEngine", () => {
  test("types values the way a spreadsheet would", async () => {
    const content = await cells("Name,Qty,Ok\nPens,12,TRUE\nPads,4,false");

    expect(content.get("A2")).toBe("Pens");
    expect(content.get("B2")).toBe(12);
    expect(content.get("C2")).toBe(true);
    expect(content.get("C3")).toBe(false);
  });

  test("keeps a leading-zero identifier intact", async () => {
    const content = await cells("Part\n007");
    expect(content.get("A2")).toBe("007");
  });

  test("reads a leading = as a formula", async () => {
    const content = await cells("A,B\n2,=A2*3");
    expect(content.get("B2")).toBe("=A2*3");
  });

  test("leaves empty cells out entirely", async () => {
    const content = await cells("a,,c");
    expect(content.has("B1")).toBe(false);
    expect(content.get("C1")).toBe("c");
  });

  test("detects tabs without being told", async () => {
    const content = await cells("Name\tQty\nPens\t12");
    expect(content.get("A1")).toBe("Name");
    expect(content.get("B2")).toBe(12);
  });

  test("honours an explicit delimiter over detection", async () => {
    const content = await cells("a;b\n1;2", { delimiter: ";" });
    expect(content.get("B2")).toBe(2);
  });

  test("can trim fields when asked", async () => {
    expect((await cells("  a  ,b")).get("A1")).toBe("  a  ");
    expect((await cells("  a  ,b", { trimFields: true })).get("A1")).toBe("a");
  });

  test("names the sheet and records the file name", async () => {
    const data = await csvToFormulaEngine("a,b", { sheetName: "Sales", fileName: "sales.csv" });
    expect(data.sheets[0]!.name).toBe("Sales");
    expect(data.workbookMetadata?.fileName).toBe("sales.csv");
  });

  test("reports a formula the engine cannot evaluate", async () => {
    const { diagnostics } = await csvToFormulaEngineDetailed("=VLOOKUP(A1,B:C,2,FALSE)");
    expect(diagnostics.some((d) => d.code === "unsupported-function")).toBe(true);
  });

  test("loads into an engine that then recalculates", async () => {
    const data = await csvToFormulaEngine(
      "Item,Qty,Price,Total\nPens,12,1.5,=B2*C2\nPads,4,3.25,=B3*C3\n,,,=SUM(D2:D3)",
      { sheetName: "Data" }
    );
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: WORKBOOK, data });

    expect(engine.getCellValue(at("Data", "D2"))).toBe(18);
    expect(engine.getCellValue(at("Data", "D4"))).toBe(31);
  });
});

describe("spreadsheetToFormulaEngine", () => {
  test("reads delimited text", async () => {
    const data = await spreadsheetToFormulaEngine(
      new TextEncoder().encode("a,b\n1,2")
    );
    expect((data.sheets[0]!.content as Map<string, unknown>).get("B2")).toBe(2);
  });

  test("reads an xlsx by its bytes rather than its name", async () => {
    // A workbook saved with a .csv extension is a common export mistake; the
    // ZIP signature settles it.
    const bytes = await Bun.file(
      new URL("../fixtures/generated/basic.xlsx", import.meta.url).pathname
    ).bytes();
    const data = await spreadsheetToFormulaEngine(bytes);
    expect(data.sheets.map((sheet) => sheet.name)).toEqual(["Data", "Summary"]);
  });

  test("refuses a corrupt workbook rather than reading it as text", async () => {
    // Every byte sequence "parses" as delimited text, so without this a broken
    // .xlsx becomes a one-cell sheet of mojibake and reports success.
    const broken = new File([new Uint8Array([0x50, 0x01, 0x02, 0x03])], "broken.xlsx");
    await expect(spreadsheetToFormulaEngine(broken)).rejects.toThrow(
      /named as an Excel workbook but is not one/
    );
  });

  test("refuses binary content when the name gives no hint", async () => {
    const binary = new File([new Uint8Array([1, 2, 0, 4])], "mystery");
    await expect(spreadsheetToFormulaEngine(binary)).rejects.toThrow(/binary data/);
  });

  test("names the sheet after a dropped file", async () => {
    const file = new File(["a,b\n1,2"], "quarterly sales.csv", { type: "text/csv" });
    const data = await spreadsheetToFormulaEngine(file);
    expect(data.sheets[0]!.name).toBe("quarterly sales");
    expect(data.workbookMetadata?.fileName).toBe("quarterly sales.csv");
  });
});
