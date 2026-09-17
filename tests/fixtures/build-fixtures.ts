/**
 * Generates the .xlsx fixtures the tests run against.
 *
 * These are written by ExcelJS rather than hand-rolled XML so the tests exercise
 * real OOXML: shared strings, shared formulas, style indirection, table parts
 * and conditional formatting blocks as a spreadsheet application actually emits
 * them.
 *
 * Run with `bun run fixtures`.
 */

import ExcelJS from "exceljs";
import { mkdir } from "node:fs/promises";
import path from "node:path";

const OUT_DIR = path.join(import.meta.dir, "generated");

async function basic(): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();

  const data = wb.addWorksheet("Data");
  data.getCell("A1").value = "Item";
  data.getCell("B1").value = "Qty";
  data.getCell("C1").value = "Price";
  data.getCell("D1").value = "Total";

  const rows: [string, number, number][] = [
    ["Pens", 12, 1.5],
    ["Pads", 4, 3.25],
    ["Clips", 30, 0.4],
  ];
  rows.forEach(([item, qty, price], index) => {
    const row = index + 2;
    data.getCell(`A${row}`).value = item;
    data.getCell(`B${row}`).value = qty;
    data.getCell(`C${row}`).value = price;
    data.getCell(`D${row}`).value = { formula: `B${row}*C${row}` };
  });

  data.getCell("D5").value = { formula: "SUM(D2:D4)" };
  data.getCell("A7").value = true;
  data.getCell("A8").value = "plain text";
  data.getCell("A9").value = 0.075;

  // Header styling, to exercise fills, fonts and borders.
  ["A1", "B1", "C1", "D1"].forEach((ref) => {
    const cell = data.getCell(ref);
    cell.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF2563EB" },
    };
    cell.border = { bottom: { style: "thin", color: { argb: "FF000000" } } };
  });

  data.getCell("C2").numFmt = "0.00";
  data.getCell("A9").numFmt = "0.0%";

  const summary = wb.addWorksheet("Summary");
  summary.getCell("A1").value = "Grand total";
  summary.getCell("B1").value = { formula: "Data!D5" };
  summary.getCell("A2").value = "Doubled";
  summary.getCell("B2").value = { formula: "B1*2" };

  return wb;
}

async function withTable(): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet("Sales");

  sheet.addTable({
    name: "SalesTable",
    ref: "A1",
    headerRow: true,
    columns: [
      { name: "Region" },
      { name: "Amount" },
    ],
    rows: [
      ["North", 100],
      ["South", 250],
      ["East", 75],
    ],
  });

  sheet.getCell("D1").value = { formula: "SUM(SalesTable[Amount])" };

  return wb;
}

async function withConditionalFormatting(): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet("Scores");

  sheet.getCell("A1").value = "Score";
  [55, 72, 91, 40, 88].forEach((score, index) => {
    sheet.getCell(`A${index + 2}`).value = score;
  });

  sheet.addConditionalFormatting({
    ref: "A2:A6",
    rules: [
      {
        type: "cellIs",
        operator: "greaterThan",
        formulae: ["80"],
        priority: 1,
        style: {
          fill: {
            type: "pattern",
            pattern: "solid",
            bgColor: { argb: "FF22C55E" },
          },
        },
      },
      {
        type: "cellIs",
        operator: "lessThan",
        formulae: ["50"],
        priority: 2,
        style: {
          fill: {
            type: "pattern",
            pattern: "solid",
            bgColor: { argb: "FFEF4444" },
          },
        },
      },
    ],
  });

  sheet.addConditionalFormatting({
    ref: "B2:B6",
    rules: [
      {
        type: "colorScale",
        priority: 3,
        cfvo: [{ type: "min" }, { type: "max" }],
        color: [{ argb: "FFFFFFFF" }, { argb: "FF2563EB" }],
      },
    ],
  });
  [1, 2, 3, 4, 5].forEach((value, index) => {
    sheet.getCell(`B${index + 2}`).value = value;
  });

  return wb;
}

async function withNamesAndDates(): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet("Main");

  sheet.getCell("A1").value = new Date(Date.UTC(2024, 0, 15));
  sheet.getCell("A1").numFmt = "yyyy-mm-dd";
  sheet.getCell("A2").value = new Date(Date.UTC(2024, 5, 30));
  sheet.getCell("A2").numFmt = "yyyy-mm-dd";
  sheet.getCell("B1").value = 45000;

  sheet.getCell("C1").value = 0.25;
  wb.definedNames.add("Main!$C$1", "TaxRate");

  sheet.getCell("C2").value = { formula: "C1*100" };

  // A formula using a function the engine does not implement.
  sheet.getCell("D1").value = { formula: "VLOOKUP(A1,A1:B2,2,FALSE)" };

  // Merged cells and a frozen pane.
  sheet.mergeCells("F1:G2");
  sheet.getCell("F1").value = "Merged";
  sheet.views = [{ state: "frozen", xSplit: 1, ySplit: 1 }];
  sheet.getColumn(1).width = 22;

  const hidden = wb.addWorksheet("Hidden");
  hidden.state = "hidden";
  hidden.getCell("A1").value = "tucked away";

  return wb;
}

async function withSharedFormulas(): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet("Shared");

  for (let row = 1; row <= 5; row++) {
    sheet.getCell(`A${row}`).value = row * 10;
    sheet.getCell(`B${row}`).value = row;
  }

  // ExcelJS writes these as a shared formula group with one master.
  sheet.fillFormula("C1:C5", "A1+B1", undefined);

  return wb;
}

const FIXTURES: Record<string, () => Promise<ExcelJS.Workbook>> = {
  "basic.xlsx": basic,
  "table.xlsx": withTable,
  "conditional-formatting.xlsx": withConditionalFormatting,
  "names-and-dates.xlsx": withNamesAndDates,
  "shared-formulas.xlsx": withSharedFormulas,
};

async function main(): Promise<void> {
  await mkdir(OUT_DIR, { recursive: true });
  for (const [fileName, build] of Object.entries(FIXTURES)) {
    const wb = await build();
    const target = path.join(OUT_DIR, fileName);
    await wb.xlsx.writeFile(target);
    console.log("wrote", path.relative(process.cwd(), target));
  }
}

await main();
