# excel-to-formula-engine

Convert an Excel `.xlsx` file into [`@ricsam/formula-engine`](https://github.com/ricsam/formula-engine)
`WorkbookData`, in the browser.

```ts
import { excelToFormulaEngine } from "@ricsam/excel-to-formula-engine";

const data = await excelToFormulaEngine(blob);
engine.addWorkbook({ workbookName: "Budget", data });
```

Pure TypeScript — no WASM, no build step, no server. An `.xlsx` is a ZIP of XML,
so it unzips with [fflate](https://github.com/101arrowz/fflate) and parses with
[fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser).

## Install

```bash
bun add @ricsam/excel-to-formula-engine @ricsam/formula-engine
```

## Usage

`excelToFormulaEngine` accepts a `Blob` or `File` straight from a file input, or
raw bytes as an `ArrayBuffer` / `Uint8Array`.

```ts
import { FormulaEngine } from "@ricsam/formula-engine";
import { excelToFormulaEngine } from "@ricsam/excel-to-formula-engine";

input.addEventListener("change", async () => {
  const file = input.files?.[0];
  if (!file) return;

  const data = await excelToFormulaEngine(file);

  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook({ workbookName: file.name, data });

  // Formulas are live — the engine recalculates them.
  engine.getCellValue({
    workbookName: file.name,
    sheetName: "Data",
    rowIndex: 4,
    colIndex: 3,
  });
});
```

The whole workbook loads as a single undo step, so one `Ctrl-Z` reverts an
import.

### Diagnostics

Anything that could not be translated is reported rather than silently dropped.
`excelToFormulaEngineDetailed` returns the diagnostics beside the data; they are
also stored in `data.workbookMetadata.diagnostics`.

```ts
const { data, diagnostics } = await excelToFormulaEngineDetailed(file);

for (const diagnostic of diagnostics) {
  console.warn(
    `[${diagnostic.code}] ${diagnostic.sheetName}!${diagnostic.cellReference}`,
    diagnostic.message
  );
}
```

### Options

```ts
await excelToFormulaEngine(file, {
  sheets: ["Summary"],          // import only these sheets
  styles: false,                // skip fills, fonts, borders
  conditionalFormatting: false, // skip conditional formatting
  tables: false,                // skip Excel tables
  definedNames: false,          // skip defined names
  cellMetadata: false,          // skip number formats and hyperlinks
  convertDates: false,          // keep date serials as raw numbers
});
```

## What gets translated

| Excel | Engine | Notes |
| --- | --- | --- |
| Cell values | Cell content | Strings, numbers, booleans and error values keep their type. |
| Formulas | Formulas | Kept as formulas, not as Excel's cached results — the engine recalculates. |
| Shared formulas | Expanded formulas | Excel stores one formula per group; each member gets its references offset. |
| Sheet order | Sheet order | Array order becomes tab order. |
| Cross-sheet references | Unchanged | `Data!D5` parses identically. |
| Excel tables | Tables | Structured references (`Sales[Amount]`) work. Tables without a header row are skipped. |
| Defined names | Named expressions | Sheet-scoped names stay sheet-scoped. Built-ins like `_xlnm.Print_Area` are skipped. |
| Fills, fonts, borders | `CellStyle` | Theme and indexed colours are resolved, including tints. |
| Conditional formatting | Conditional styles | `cellIs`, `expression`, text rules and colour scales. See below. |
| Dates | ISO strings | The engine has no date type, so `45306` becomes `"2024-01-15"`. |
| Number formats, hyperlinks | Cell metadata | Stored for your UI; the engine does not interpret them. |
| Merges, frozen panes, column widths, hidden sheets | Sheet metadata | Same — stored for your UI to apply. |

### Formulas the engine cannot evaluate

The engine implements around 40 functions. A formula using anything else — an
Excel file will often reach for `VLOOKUP`, `TEXT`, `ROUND` or `SUMPRODUCT` — is
**imported unchanged** and evaluates to `#NAME?`, with a warning diagnostic
naming the function and the cell.

This keeps the author's intent visible and editable in the cell, rather than
replacing it with a frozen copy of whatever Excel last calculated. Everything
needed to do the opposite is in the diagnostics if you prefer that trade.

### Conditional formatting

Excel's rule catalogue is much larger than the engine's, which has a formula
rule and a two-stop gradient. These map across:

- `cellIs` with every operator, including `between` / `notBetween`
- `expression` (custom formula rules)
- `containsText`, `notContainsText`, `beginsWith`, `endsWith`
- `containsBlanks`, `notContainsBlanks`
- `colorScale` → gradient (a three-colour scale loses its midpoint, with an
  `info` diagnostic)

Data bars, icon sets, `top10`, `aboveAverage` and duplicate rules are reported
as `unsupported-conditional-format` and skipped.

### Not imported

Charts, images, shapes, pivot tables, macros, data validation, sparklines,
comments, and cell protection. The engine has no representation for them.
Password-protected and legacy `.xls` files are rejected with a clear error.

## Demo

```bash
bun run demo:serve
```

Opens a page at <http://localhost:4321> where you can drop in a workbook and see
it recalculated live. `demo/sample.xlsx` is there to try.

## Development

```bash
bun install
bun run fixtures   # generate .xlsx test fixtures with ExcelJS
bun test
bun run typecheck
```

Fixtures are written by ExcelJS rather than hand-rolled XML, so the tests run
against real OOXML — shared strings, shared formulas, style indirection, table
parts and conditional formatting blocks as a spreadsheet application emits them.
`bun test` regenerates them first.

## License

ISC
