import { describe, expect, test } from "bun:test";
import { translateConditionalFormat } from "../../src/translate/conditional-format";
import type { RawConditionalFormat } from "../../src/xlsx/worksheet";

const options = {
  sheetName: "Input",
  differentialFormats: [{ backgroundColor: "#ff0000" }],
};

const rule = (overrides: Partial<RawConditionalFormat>): RawConditionalFormat => ({
  ranges: ["$B$2:$B$40", "D2:D40"],
  type: "expression",
  formulas: ["B2>0"],
  dxfId: 0,
  priority: 1,
  ...overrides,
});

/**
 * A skipped rule is reported per sheet, but a person looking at the report
 * wants to know *which cells* lost their formatting, so the note carries the
 * rule's own ranges.
 */
describe("conditional formatting diagnostics", () => {
  test("a rule the engine cannot express names the cells it applied to", () => {
    const { result, diagnostics } = translateConditionalFormat(
      rule({ type: "duplicateValues", formulas: [] }),
      options
    );
    expect(result).toBeUndefined();
    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0]).toMatchObject({
      code: "unsupported-conditional-format",
      sheetName: "Input",
      ranges: ["B2:B40", "D2:D40"],
    });
  });

  test("a rule with no colour to apply names its cells too", () => {
    const { diagnostics } = translateConditionalFormat(rule({ dxfId: undefined }), options);
    expect(diagnostics[0]?.ranges).toEqual(["B2:B40", "D2:D40"]);
  });

  test("a downgraded colour scale is still imported, and says where", () => {
    const { result, diagnostics } = translateConditionalFormat(
      rule({
        type: "colorScale",
        formulas: [],
        colorScale: {
          cfvo: [{ type: "min" }, { type: "percentile", val: "50" }, { type: "max" }],
          colors: ["FFFFFFFF", "FFFFFF00", "FF2563EB"],
        },
      }),
      options
    );
    expect(result).toBeDefined();
    expect(diagnostics.length).toBeGreaterThan(0);
    expect(diagnostics.every((diagnostic) => diagnostic.ranges?.[0] === "B2:B40")).toBe(true);
  });

  test("a rule that translates cleanly reports nothing", () => {
    expect(translateConditionalFormat(rule({}), options).diagnostics).toEqual([]);
  });
});
