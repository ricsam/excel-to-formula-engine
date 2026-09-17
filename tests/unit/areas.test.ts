import { describe, expect, test } from "bun:test";
import { coalesceCellsIntoAreas } from "../../src/convert";
import { parseA1 } from "../../src/a1";

function cells(...references: string[]) {
  return references.map((reference) => {
    const parsed = parseA1(reference);
    if (!parsed) {
      throw new Error(`bad reference ${reference}`);
    }
    return parsed;
  });
}

/** Render an area back to A1 notation so expectations stay readable. */
function describeAreas(
  areas: ReturnType<typeof coalesceCellsIntoAreas>
): string[] {
  return areas
    .map((area) => {
      const { start, end } = area.range;
      const col = (index: number) => {
        let label = "";
        let remaining = index + 1;
        while (remaining > 0) {
          const rest = (remaining - 1) % 26;
          label = String.fromCharCode(65 + rest) + label;
          remaining = Math.floor((remaining - rest) / 26);
        }
        return label;
      };
      const endCol = end.col.type === "number" ? end.col.value : -1;
      const endRow = end.row.type === "number" ? end.row.value : -1;
      return `${col(start.col)}${start.row + 1}:${col(endCol)}${endRow + 1}`;
    })
    .sort();
}

describe("coalesceCellsIntoAreas", () => {
  test("merges a horizontal run into one area", () => {
    const areas = coalesceCellsIntoAreas("S", cells("A1", "B1", "C1", "D1"));
    expect(describeAreas(areas)).toEqual(["A1:D1"]);
  });

  test("merges a vertical run into one area", () => {
    const areas = coalesceCellsIntoAreas("S", cells("A1", "A2", "A3"));
    expect(describeAreas(areas)).toEqual(["A1:A3"]);
  });

  test("merges a solid block into one area", () => {
    const areas = coalesceCellsIntoAreas(
      "S",
      cells("A1", "B1", "A2", "B2", "A3", "B3")
    );
    expect(describeAreas(areas)).toEqual(["A1:B3"]);
  });

  test("splits a horizontal gap into separate areas", () => {
    const areas = coalesceCellsIntoAreas("S", cells("A1", "B1", "D1", "E1"));
    expect(describeAreas(areas)).toEqual(["A1:B1", "D1:E1"]);
  });

  test("splits a vertical gap into separate areas", () => {
    const areas = coalesceCellsIntoAreas("S", cells("A1", "A2", "A5"));
    expect(describeAreas(areas)).toEqual(["A1:A2", "A5:A5"]);
  });

  test("does not merge rows whose runs differ in width", () => {
    const areas = coalesceCellsIntoAreas(
      "S",
      cells("A1", "B1", "C1", "A2", "B2")
    );
    expect(describeAreas(areas)).toEqual(["A1:C1", "A2:B2"]);
  });

  test("handles a single cell", () => {
    expect(describeAreas(coalesceCellsIntoAreas("S", cells("C7")))).toEqual([
      "C7:C7",
    ]);
  });

  test("handles an empty input", () => {
    expect(coalesceCellsIntoAreas("S", [])).toEqual([]);
  });

  test("collapses a large uniform block to a single area", () => {
    // The point of coalescing: a 50x50 styled block must not become 2500 rules.
    const block = [];
    for (let row = 0; row < 50; row++) {
      for (let col = 0; col < 50; col++) {
        block.push({ colIndex: col, rowIndex: row });
      }
    }
    const areas = coalesceCellsIntoAreas("S", block);
    expect(areas).toHaveLength(1);
    expect(describeAreas(areas)).toEqual(["A1:AX50"]);
  });

  test("tags every area with the sheet name", () => {
    const areas = coalesceCellsIntoAreas("Sheet2", cells("A1", "C3"));
    expect(areas.every((area) => area.sheetName === "Sheet2")).toBe(true);
  });
});
