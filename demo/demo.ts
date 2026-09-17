/**
 * Browser demo: drop an .xlsx in, see it loaded into a live FormulaEngine.
 *
 * Run with `bun run demo`.
 */

import { FormulaEngine, parseCellReference } from "@ricsam/formula-engine";
import { excelToFormulaEngineDetailed } from "../index";

const fileInput = document.getElementById("file") as HTMLInputElement;
const output = document.getElementById("output") as HTMLElement;
const status = document.getElementById("status") as HTMLElement;

fileInput.addEventListener("change", async () => {
  const file = fileInput.files?.[0];
  if (!file) {
    return;
  }

  status.textContent = `Converting ${file.name}…`;
  output.innerHTML = "";

  try {
    const started = performance.now();
    // This is the whole public API.
    const { data, diagnostics } = await excelToFormulaEngineDetailed(file);

    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: "Imported", data });
    const elapsed = Math.round(performance.now() - started);

    status.textContent =
      `Loaded ${data.sheets.length} sheet(s) in ${elapsed}ms · ` +
      `${diagnostics.length} diagnostic(s)`;
    status.dataset["ready"] = "true";

    for (const sheet of data.sheets) {
      output.appendChild(renderSheet(engine, sheet.name));
    }

    if (diagnostics.length > 0) {
      const list = document.createElement("ul");
      list.className = "diagnostics";
      for (const diagnostic of diagnostics) {
        const item = document.createElement("li");
        item.className = diagnostic.severity;
        item.textContent =
          `[${diagnostic.code}] ` +
          (diagnostic.sheetName ? `${diagnostic.sheetName}!` : "") +
          (diagnostic.cellReference ?? "") +
          ` ${diagnostic.message}`;
        list.appendChild(item);
      }
      output.appendChild(list);
    }
  } catch (error) {
    status.dataset["ready"] = "error";
    status.textContent = `Failed: ${(error as Error).message}`;
  }
});

/** Render the used range of one sheet, showing calculated values. */
function renderSheet(engine: FormulaEngine, sheetName: string): HTMLElement {
  const section = document.createElement("section");
  const heading = document.createElement("h2");
  heading.textContent = sheetName;
  section.appendChild(heading);

  const content = engine.getSheetSerialized({
    workbookName: "Imported",
    sheetName,
  });

  let maxRow = 0;
  let maxCol = 0;
  for (const reference of content.keys()) {
    const { rowIndex, colIndex } = parseCellReference(reference);
    maxRow = Math.max(maxRow, rowIndex);
    maxCol = Math.max(maxCol, colIndex);
  }

  const table = document.createElement("table");
  table.dataset["sheet"] = sheetName;

  for (let row = 0; row <= Math.min(maxRow, 30); row++) {
    const tr = document.createElement("tr");
    for (let col = 0; col <= Math.min(maxCol, 15); col++) {
      const td = document.createElement("td");
      const address = {
        workbookName: "Imported",
        sheetName,
        rowIndex: row,
        colIndex: col,
      };
      const value = engine.getCellValue(address);
      td.textContent = value === undefined ? "" : String(value);
      td.dataset["ref"] = `${columnLabel(col)}${row + 1}`;

      const style = engine.getCellStyle(address);
      if (style) {
        if (style.backgroundColor) td.style.backgroundColor = style.backgroundColor;
        if (style.color) td.style.color = style.color;
        if (style.bold) td.style.fontWeight = "bold";
        if (style.italic) td.style.fontStyle = "italic";
      }
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }

  section.appendChild(table);
  return section;
}

function columnLabel(index: number): string {
  let label = "";
  let remaining = index + 1;
  while (remaining > 0) {
    const rest = (remaining - 1) % 26;
    label = String.fromCharCode(65 + rest) + label;
    remaining = Math.floor((remaining - rest) / 26);
  }
  return label;
}
