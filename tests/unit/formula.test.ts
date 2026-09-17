import { describe, expect, test } from "bun:test";
import {
  findUnsupportedFunctions,
  offsetFormulaReferences,
  translateFormula,
} from "../../src/translate/formula";

describe("translateFormula", () => {
  test("keeps ordinary formulas unchanged", () => {
    expect(translateFormula("=SUM(A1:A10)").formula).toBe("=SUM(A1:A10)");
    expect(translateFormula("SUM(A1:A10)").formula).toBe("=SUM(A1:A10)");
  });

  test("strips the _xlfn storage prefix", () => {
    expect(translateFormula("=_xlfn.XLOOKUP(A1,B:B,C:C)").formula).toBe(
      "=XLOOKUP(A1,B:B,C:C)"
    );
    expect(translateFormula("=_xlfn._xlws.SORT(A1:A9)").formula).toBe(
      "=SORT(A1:A9)"
    );
  });

  test("drops the implicit intersection operator", () => {
    expect(translateFormula("=@A1:A10").formula).toBe("=A1:A10");
    expect(translateFormula("=SUM(@Table1[Amount])").formula).toBe(
      "=SUM(Table1[Amount])"
    );
  });

  test("keeps an @ that is inside a string literal", () => {
    expect(translateFormula('=CONCATENATE("a@b.com",A1)').formula).toBe(
      '=CONCATENATE("a@b.com",A1)'
    );
  });

  test("keeps formulas with unsupported functions and reports them", () => {
    const result = translateFormula("=VLOOKUP(A1,B:C,2,FALSE)", {
      sheetName: "Sheet1",
      cellReference: "D1",
    });

    expect(result.formula).toBe("=VLOOKUP(A1,B:C,2,FALSE)");
    expect(result.diagnostics).toHaveLength(1);
    expect(result.diagnostics[0]).toMatchObject({
      severity: "warning",
      code: "unsupported-function",
      sheetName: "Sheet1",
      cellReference: "D1",
    });
    expect(result.diagnostics[0]!.message).toContain("VLOOKUP");
  });

  test("reports nothing for functions the engine implements", () => {
    const result = translateFormula("=IFERROR(XLOOKUP(A1,B:B,C:C),0)");
    expect(result.diagnostics).toEqual([]);
  });
});

describe("findUnsupportedFunctions", () => {
  test("ignores table column references", () => {
    expect(findUnsupportedFunctions("SUM(Sales[Amount])")).toEqual([]);
  });

  test("ignores named expressions that are not calls", () => {
    expect(findUnsupportedFunctions("TaxRate * A1")).toEqual([]);
  });

  test("ignores function-like text inside string literals", () => {
    expect(findUnsupportedFunctions('IF(A1="VLOOKUP(",1,0)')).toEqual([]);
  });

  test("finds nested unsupported calls", () => {
    expect(findUnsupportedFunctions("SUM(ROUND(A1,2))")).toEqual(["ROUND"]);
  });

  test("reports each unsupported function once", () => {
    expect(findUnsupportedFunctions("ROUND(A1,2)+ROUND(A2,2)")).toEqual([
      "ROUND",
    ]);
  });
});

describe("offsetFormulaReferences", () => {
  test("shifts relative references", () => {
    expect(offsetFormulaReferences("A1+B1", 0, 1)).toBe("A2+B2");
    expect(offsetFormulaReferences("A1+B1", 2, 0)).toBe("C1+D1");
  });

  test("leaves absolute parts pinned", () => {
    expect(offsetFormulaReferences("$A$1+B1", 1, 1)).toBe("$A$1+C2");
    expect(offsetFormulaReferences("$A1+A$1", 1, 1)).toBe("$A2+B$1");
  });

  test("shifts both ends of a range", () => {
    expect(offsetFormulaReferences("SUM(A1:A10)", 1, 0)).toBe("SUM(B1:B10)");
  });

  test("does not touch function names or string literals", () => {
    expect(offsetFormulaReferences('IF(A1="B1",A1,0)', 0, 1)).toBe(
      'IF(A2="B1",A2,0)'
    );
  });

  test("returns the formula unchanged for a zero offset", () => {
    expect(offsetFormulaReferences("SUM(A1:A10)", 0, 0)).toBe("SUM(A1:A10)");
  });

  test("produces #REF! when a reference falls off the grid", () => {
    expect(offsetFormulaReferences("A1", 0, -1)).toBe("#REF!");
  });

  test("crosses the Z/AA boundary correctly", () => {
    expect(offsetFormulaReferences("Z1", 1, 0)).toBe("AA1");
    expect(offsetFormulaReferences("AA1", -1, 0)).toBe("Z1");
  });

  test("preserves a sheet qualifier", () => {
    expect(offsetFormulaReferences("Sheet2!A1", 0, 1)).toBe("Sheet2!A2");
  });
});
