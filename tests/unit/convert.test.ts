import { beforeAll, describe, expect, test } from "bun:test";
import path from "node:path";
import { readFile } from "node:fs/promises";
import { FormulaEngine, parseCellReference } from "@ricsam/formula-engine";
import type { CellAddress } from "@ricsam/formula-engine";
import {
  excelToFormulaEngine,
  excelToFormulaEngineDetailed,
} from "../../index";
import type { ExcelWorkbookData } from "../../index";

const FIXTURE_DIR = path.join(import.meta.dir, "..", "fixtures", "generated");

async function loadFixture(name: string): Promise<Blob> {
  const bytes = await readFile(path.join(FIXTURE_DIR, name));
  return new Blob([bytes as unknown as ArrayBuffer]);
}

const WORKBOOK = "Imported";

function intoEngine(data: ExcelWorkbookData): FormulaEngine {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook({ workbookName: WORKBOOK, data });
  return engine;
}

function at(sheetName: string, reference: string): CellAddress {
  const { rowIndex, colIndex } = parseCellReference(reference);
  return { workbookName: WORKBOOK, sheetName, rowIndex, colIndex };
}

describe("basic.xlsx", () => {
  let data: ExcelWorkbookData;
  let engine: FormulaEngine;

  beforeAll(async () => {
    data = await excelToFormulaEngine(await loadFixture("basic.xlsx"));
    engine = intoEngine(data);
  });

  test("imports sheets in workbook order", () => {
    expect(engine.getOrderedSheetNames(WORKBOOK)).toEqual(["Data", "Summary"]);
  });

  test("imports literal values with their types intact", () => {
    expect(engine.getCellValue(at("Data", "A1"))).toBe("Item");
    expect(engine.getCellValue(at("Data", "B2"))).toBe(12);
    expect(engine.getCellValue(at("Data", "A7"))).toBe(true);
    expect(engine.getCellValue(at("Data", "A8"))).toBe("plain text");
  });

  test("keeps formulas as formulas and recalculates them", () => {
    const content = data.sheets.find((sheet) => sheet.name === "Data")?.content;
    expect(new Map(Object.entries(content ?? {})).get("D2") ?? (content as Map<string, unknown>).get("D2")).toBe("=B2*C2");

    // 12 * 1.5
    expect(engine.getCellValue(at("Data", "D2"))).toBe(18);
    // 18 + 13 + 12
    expect(engine.getCellValue(at("Data", "D5"))).toBe(43);
  });

  test("resolves cross-sheet references", () => {
    expect(engine.getCellValue(at("Summary", "B1"))).toBe(43);
    expect(engine.getCellValue(at("Summary", "B2"))).toBe(86);
  });

  test("imports header styling", () => {
    const style = engine.getCellStyle(at("Data", "A1"));
    expect(style).toBeDefined();
    expect(style?.bold).toBe(true);
    expect(style?.backgroundColor).toBe("#2563eb");
    expect(style?.color).toBe("#ffffff");
    expect(style?.borderSides?.bottom).toBe(true);
  });

  test("does not style cells that had no format", () => {
    expect(engine.getCellStyle(at("Data", "A8"))?.bold).toBeUndefined();
  });

  test("records number formats as cell metadata", () => {
    const metadata = engine.getCellMetadata(at("Data", "A9")) as
      | { numberFormat?: string }
      | undefined;
    expect(metadata?.numberFormat).toBe("0.0%");
  });

  test("reports no warnings for a workbook the engine fully supports", () => {
    const warnings = (data.workbookMetadata?.diagnostics ?? []).filter(
      (diagnostic) => diagnostic.severity === "warning"
    );
    expect(warnings).toEqual([]);
  });
});

describe("table.xlsx", () => {
  test("imports a table whose structured references evaluate", async () => {
    const data = await excelToFormulaEngine(await loadFixture("table.xlsx"));
    const engine = intoEngine(data);

    const table = engine.getTable({
      workbookName: WORKBOOK,
      tableName: "SalesTable",
    });
    expect(table).toBeDefined();
    expect(Array.from(table!.headers.keys())).toEqual(["Region", "Amount"]);
    expect(table!.start).toEqual({ rowIndex: 0, colIndex: 0 });

    // 100 + 250 + 75
    expect(engine.getCellValue(at("Sales", "D1"))).toBe(425);
  });
});

describe("conditional-formatting.xlsx", () => {
  let data: ExcelWorkbookData;

  beforeAll(async () => {
    data = await excelToFormulaEngine(
      await loadFixture("conditional-formatting.xlsx")
    );
  });

  test("translates cellIs rules into engine formula conditions", () => {
    const formulaRules = (data.conditionalStyles ?? []).filter(
      (style) => style.condition.type === "formula"
    );
    expect(formulaRules.length).toBe(2);

    const formulas = formulaRules.map((rule) =>
      rule.condition.type === "formula" ? rule.condition.formula : ""
    );
    expect(formulas).toContain("=A2>80");
    expect(formulas).toContain("=A2<50");
  });

  test("translates a colour scale into a gradient condition", () => {
    const gradient = (data.conditionalStyles ?? []).find(
      (style) => style.condition.type === "gradient"
    );
    expect(gradient).toBeDefined();
    if (gradient?.condition.type === "gradient") {
      expect(gradient.condition.min.type).toBe("lowest_value");
      expect(gradient.condition.max.type).toBe("highest_value");
    }
  });

  test("the imported rules load into the engine", () => {
    const engine = intoEngine(data);
    expect(engine.getAllConditionalStyles().length).toBe(3);
  });
});

describe("names-and-dates.xlsx", () => {
  let data: ExcelWorkbookData;
  let engine: FormulaEngine;

  beforeAll(async () => {
    data = await excelToFormulaEngine(
      await loadFixture("names-and-dates.xlsx")
    );
    engine = intoEngine(data);
  });

  test("converts date-formatted numbers to ISO strings", () => {
    expect(engine.getCellValue(at("Main", "A1"))).toBe("2024-01-15");
    expect(engine.getCellValue(at("Main", "A2"))).toBe("2024-06-30");
  });

  test("leaves plain numbers alone", () => {
    expect(engine.getCellValue(at("Main", "B1"))).toBe(45000);
  });

  test("imports defined names as named expressions", () => {
    const names = data.namedExpressions ?? [];
    expect(names.some((entry) => entry.name === "TaxRate")).toBe(true);
    expect(
      engine.hasNamedExpression({
        expressionName: "TaxRate",
        workbookName: WORKBOOK,
      })
    ).toBe(true);
  });

  test("keeps an unsupported formula and reports it", () => {
    const content = data.sheets.find((sheet) => sheet.name === "Main")?.content;
    const d1 = (content as Map<string, unknown>).get("D1");
    expect(d1).toBe("=VLOOKUP(A1,A1:B2,2,FALSE)");

    const warning = (data.workbookMetadata?.diagnostics ?? []).find(
      (diagnostic) => diagnostic.code === "unsupported-function"
    );
    expect(warning).toBeDefined();
    expect(warning?.message).toContain("VLOOKUP");
    expect(warning?.cellReference).toBe("D1");
  });

  test("records merges, frozen panes and column widths as sheet metadata", () => {
    const metadata = engine.getSheetMetadata({
      workbookName: WORKBOOK,
      sheetName: "Main",
    }) as
      | {
          merges?: string[];
          frozen?: { rows: number; cols: number };
          columnWidths?: Record<number, number>;
        }
      | undefined;

    expect(metadata?.merges).toContain("F1:G2");
    expect(metadata?.frozen).toEqual({ rows: 1, cols: 1 });
    expect(metadata?.columnWidths?.[0]).toBeGreaterThan(20);
  });

  test("imports hidden sheets but marks them hidden", () => {
    expect(engine.getOrderedSheetNames(WORKBOOK)).toContain("Hidden");
    const metadata = engine.getSheetMetadata({
      workbookName: WORKBOOK,
      sheetName: "Hidden",
    }) as { visible?: boolean } | undefined;
    expect(metadata?.visible).toBe(false);
  });
});

describe("shared-formulas.xlsx", () => {
  test("expands a shared formula group across its members", async () => {
    const data = await excelToFormulaEngine(
      await loadFixture("shared-formulas.xlsx")
    );
    const engine = intoEngine(data);

    const content = data.sheets[0]!.content as Map<string, unknown>;
    expect(content.get("C1")).toBe("=A1+B1");
    expect(content.get("C3")).toBe("=A3+B3");
    expect(content.get("C5")).toBe("=A5+B5");

    expect(engine.getCellValue(at("Shared", "C1"))).toBe(11);
    expect(engine.getCellValue(at("Shared", "C3"))).toBe(33);
    expect(engine.getCellValue(at("Shared", "C5"))).toBe(55);
  });
});

describe("options", () => {
  test("sheets option limits which sheets are imported", async () => {
    const data = await excelToFormulaEngine(await loadFixture("basic.xlsx"), {
      sheets: ["Summary"],
    });
    expect(data.sheets.map((sheet) => sheet.name)).toEqual(["Summary"]);
  });

  test("styles can be turned off", async () => {
    const data = await excelToFormulaEngine(await loadFixture("basic.xlsx"), {
      styles: false,
    });
    expect(data.cellStyles).toBeUndefined();
  });

  test("date conversion can be turned off", async () => {
    const data = await excelToFormulaEngine(
      await loadFixture("names-and-dates.xlsx"),
      { convertDates: false }
    );
    const content = data.sheets[0]!.content as Map<string, unknown>;
    expect(typeof content.get("A1")).toBe("number");
  });

  test("detailed conversion returns diagnostics separately", async () => {
    const result = await excelToFormulaEngineDetailed(
      await loadFixture("names-and-dates.xlsx")
    );
    expect(result.data.sheets.length).toBeGreaterThan(0);
    expect(
      result.diagnostics.some(
        (diagnostic) => diagnostic.code === "unsupported-function"
      )
    ).toBe(true);
  });
});

describe("error handling", () => {
  test("rejects a file that is not a zip archive", async () => {
    const blob = new Blob(["this is not a spreadsheet"]);
    await expect(excelToFormulaEngine(blob)).rejects.toThrow(/xlsx package/);
  });

  test("rejects a zip that is not a workbook", async () => {
    // A minimal, valid but empty zip.
    const emptyZip = new Uint8Array([
      0x50, 0x4b, 0x05, 0x06, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      0, 0,
    ]);
    await expect(excelToFormulaEngine(emptyZip)).rejects.toThrow(
      /not an Excel workbook|xlsx package/
    );
  });

  test("accepts an ArrayBuffer as well as a Blob", async () => {
    const bytes = await readFile(path.join(FIXTURE_DIR, "basic.xlsx"));
    const buffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength
    ) as ArrayBuffer;
    const data = await excelToFormulaEngine(buffer);
    expect(data.sheets.length).toBe(2);
  });
});
