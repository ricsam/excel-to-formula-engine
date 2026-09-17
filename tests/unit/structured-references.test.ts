import { describe, expect, test } from "bun:test";
import path from "node:path";
import { readFile } from "node:fs/promises";
import { FormulaEngine, parseCellReference } from "@ricsam/formula-engine";
import type { CellAddress } from "@ricsam/formula-engine";
import { excelToFormulaEngine } from "../../index";
import {
  contractThisRowReferences,
  translateFormula,
} from "../../src/translate/formula";

/**
 * Structured references across the Excel/engine boundary.
 *
 * Excel writes one spelling into the file and shows another in its formula bar,
 * and the engine reads only the one it shows. Two things follow, and both are
 * covered here: the stored `[#This Row]` selector has to be rewritten, and the
 * `@` in `[@Column]` must survive — it looks exactly like the implicit
 * intersection operator that does get stripped.
 */

const FIXTURE_DIR = path.join(import.meta.dir, "..", "fixtures", "generated");

async function loadFixture(name: string): Promise<Blob> {
  const bytes = await readFile(path.join(FIXTURE_DIR, name));
  return new Blob([bytes as unknown as ArrayBuffer]);
}

const WORKBOOK = "Imported";

function at(sheetName: string, reference: string): CellAddress {
  const { rowIndex, colIndex } = parseCellReference(reference);
  return { workbookName: WORKBOOK, sheetName, rowIndex, colIndex };
}

describe("the @ in a structured reference is not implicit intersection", () => {
  test("keeps the @ in a qualified current-row reference", () => {
    // Stripping this would silently turn one row into the whole column.
    expect(translateFormula("TTRinput[@Payload]").formula).toBe(
      "=TTRinput[@Payload]"
    );
  });

  test("keeps the @ in an unqualified current-row reference", () => {
    expect(translateFormula("[@Payload]").formula).toBe("=[@Payload]");
  });

  test("keeps the @ through a whole formula", () => {
    expect(
      translateFormula('LEFT([@Payload], FIND(",", [@Payload])-1)').formula
    ).toBe('=LEFT([@Payload], FIND(",", [@Payload])-1)');
  });

  test("still strips a genuine implicit intersection operator", () => {
    expect(translateFormula("@A1:A10").formula).toBe("=A1:A10");
    expect(translateFormula("SUM(@Table1[Amount])").formula).toBe(
      "=SUM(Table1[Amount])"
    );
  });
});

describe("contractThisRowReferences", () => {
  test("drops the table name when the formula is inside that table", () => {
    // Which is exactly the text Excel shows the author.
    expect(
      contractThisRowReferences("TTRinput[[#This Row],[Payload]]", "TTRinput")
    ).toBe("[@Payload]");
  });

  test("keeps the table name when the formula is somewhere else", () => {
    expect(
      contractThisRowReferences("TTRinput[[#This Row],[Payload]]", "Other")
    ).toBe("TTRinput[@Payload]");
  });

  test("keeps the table name when the containing table is unknown", () => {
    expect(contractThisRowReferences("TTRinput[[#This Row],[Payload]]")).toBe(
      "TTRinput[@Payload]"
    );
  });

  test("matches the table name case-insensitively, as Excel does", () => {
    expect(
      contractThisRowReferences("TTRinput[[#This Row],[Payload]]", "ttrinput")
    ).toBe("[@Payload]");
  });

  test("rewrites every occurrence in a formula", () => {
    expect(
      contractThisRowReferences(
        'LEFT(TTRinput[[#This Row],[Payload]],FIND(",",TTRinput[[#This Row],[Payload]])-1)',
        "TTRinput"
      )
    ).toBe('LEFT([@Payload],FIND(",",[@Payload])-1)');
  });

  test("handles the unspaced spelling some writers emit", () => {
    expect(
      contractThisRowReferences("TTRinput[[#ThisRow],[Amount]]", "TTRinput")
    ).toBe("[@Amount]");
  });

  test("keeps a column span bracketed", () => {
    expect(
      contractThisRowReferences("TTRinput[[#This Row],[A]:[C]]", "TTRinput")
    ).toBe("[@[A]:[C]]");
  });

  test("keeps a column name that needs its brackets", () => {
    expect(
      contractThisRowReferences("TTRinput[[#This Row],[Total #]]", "TTRinput")
    ).toBe("[@[Total #]]");
  });

  test("leaves other selectors alone", () => {
    expect(
      contractThisRowReferences("TTRinput[[#Headers],[Amount]]", "TTRinput")
    ).toBe("TTRinput[[#Headers],[Amount]]");
    expect(contractThisRowReferences("SUM(TTRinput[Amount])", "TTRinput")).toBe(
      "SUM(TTRinput[Amount])"
    );
  });
});

describe("calculated-column.xlsx", () => {
  test("imports the formula as the author wrote it in Excel", async () => {
    const data = await excelToFormulaEngine(
      await loadFixture("calculated-column.xlsx")
    );
    const content = data.sheets[0]!.content as Map<string, unknown>;

    // The file stores TTRinput[[#This Row],[Payload]]; Excel shows [@Payload].
    expect(content.get("C2")).toBe('=LEFT([@Payload],FIND(",",[@Payload])-1)');
  });

  test("and the engine evaluates it", async () => {
    const data = await excelToFormulaEngine(
      await loadFixture("calculated-column.xlsx")
    );
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: WORKBOOK, data });

    // Each row resolves against its own row, which is the whole point of the
    // current-row reference.
    expect(engine.getCellValue(at("Sales", "C2"))).toBe("alpha");
    expect(engine.getCellValue(at("Sales", "C3"))).toBe("gamma");
  });

  test("reports no unsupported-function warnings for it", async () => {
    const data = await excelToFormulaEngine(
      await loadFixture("calculated-column.xlsx")
    );
    const warnings = (data.workbookMetadata?.diagnostics ?? []).filter(
      (diagnostic) => diagnostic.severity === "warning"
    );
    expect(warnings).toEqual([]);
  });

  test("a whole-column reference from outside the table still works", async () => {
    const data = await excelToFormulaEngine(
      await loadFixture("calculated-column.xlsx")
    );
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: WORKBOOK, data });

    expect(engine.getCellValue(at("Sales", "E1"))).toBe(30);
  });
});
