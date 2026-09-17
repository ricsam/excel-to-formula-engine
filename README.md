# excel-to-formula-engine

Convert a spreadsheet file — `.xlsx`, `.csv` or `.tsv` — into
[`@ricsam/formula-engine`](https://github.com/ricsam/formula-engine)
`WorkbookData`, in the browser.

```ts
import { excelToFormulaEngine } from "@ricsam/excel-to-formula-engine";

const data = await excelToFormulaEngine(blob);
engine.addWorkbook({ workbookName: "Budget", data });
```

Pure TypeScript — no WASM, no native toolchain, no server round trip. An `.xlsx`
is a ZIP of XML, so it unzips with [fflate](https://github.com/101arrowz/fflate)
and parses with [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser).
Ships as ESM and CJS with type declarations; ~48 kB per bundle, dependencies
external.

## Install

```bash
bun add @ricsam/excel-to-formula-engine @ricsam/formula-engine
```

## Delimited text

CSV and TSV go through `csvToFormulaEngine`, or through
`spreadsheetToFormulaEngine` when you do not know which kind of file you have —
a drop target, typically. That one settles the format from the bytes rather than
the file name, since `.xlsx` is a ZIP and delimited text is not, so a workbook
saved with a `.csv` extension still reads correctly.

```ts
import {
  csvToFormulaEngine,
  spreadsheetToFormulaEngine
} from "@ricsam/excel-to-formula-engine";

// Delimiter is detected: comma, tab, semicolon or pipe.
const data = await csvToFormulaEngine(file);

// Or let it work out which kind of file this is.
const data = await spreadsheetToFormulaEngine(file);
```

Parsing follows RFC 4180, so a quoted field may contain the delimiter, doubled
quotes, and line breaks. The UTF-8 BOM Excel writes on every CSV export is
stripped — left in, it becomes an invisible part of the first header cell.

### What a bare string becomes

A CSV carries no types, so every cell is a judgement call. The rules err toward
keeping data recoverable:

| Field | Becomes | Why |
| --- | --- | --- |
| `42`, `-3.5`, `1.50`, `1e3` | number | Plain decimals. |
| `TRUE` / `false` | boolean | Any case, as in a spreadsheet. |
| `=B2*C2` | formula | What a spreadsheet does with the same file. |
| `007`, `01234` | **text** | Part numbers and postcodes — the mangling Excel is notorious for. |
| `12345678901234567890` | **text** | Beyond float precision; converting would silently change it. |
| `1,000`, `$5`, `50%` | **text** | Converting would drop the formatting. |
| `0x1f`, `Infinity` | **text** | `Number()` accepts these; a spreadsheet does not. |

Empty fields are left out rather than stored as empty strings, which is what
keeps a sparse sheet sparse.

There is deliberately no option to keep an `=`-leading field as text: the engine
stores a formula as a string beginning with `=` and has no escape for text that
merely looks like one, so such an option could only be honoured by altering the
value.

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
bun run test       # regenerates fixtures, then runs the suite
bun run typecheck
bun run build      # emit dist/ (ESM, CJS and .d.ts)
```

Fixtures are written by ExcelJS rather than hand-rolled XML, so the tests run
against real OOXML — shared strings, shared formulas, style indirection, table
parts and conditional formatting blocks as a spreadsheet application emits them.
They are generated rather than committed, so `bun run test` builds them first.

## Releasing

Releases go out from `main` through [changesets](https://github.com/changesets/changesets)
and npm [trusted publishing](https://docs.npmjs.com/trusted-publishers), so no
npm token is stored in the repository — the publish job mints a short-lived
OIDC token instead.

```bash
bunx changeset          # describe the change; pick patch / minor / major
bunx changeset version  # fold pending changesets into package.json + CHANGELOG
git commit -am "Release x.y.z"
git push origin main    # .github/workflows/publish.yml builds, tests, publishes
```

`changeset publish` only publishes when the version in `package.json` is not
already on npm, so pushing unrelated commits to `main` is safe.

## License

ISC
