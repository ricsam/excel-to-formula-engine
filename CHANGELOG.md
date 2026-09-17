# @ricsam/excel-to-formula-engine

## 0.2.1

### Patch Changes

- Refuse unreadable files in `spreadsheetToFormulaEngine` instead of reading them
  as text.
  
  Every byte sequence parses as delimited text, so a corrupt `.xlsx` fell through
  to the CSV reader and became a one-cell sheet of mojibake — reporting success
  for a file that could not be read. A file named `.xlsx` or `.xlsm` that is not a
  ZIP is now refused, as is binary content dropped without a usable file name.

## 0.2.0

### Minor Changes

- Add CSV and TSV import.
  
  `csvToFormulaEngine` reads delimited text into a single-sheet workbook, and
  `spreadsheetToFormulaEngine` takes any of the three formats and settles which
  one it is from the bytes rather than the file name — so a workbook saved with a
  `.csv` extension still reads correctly.
  
  The delimiter is detected between comma, tab, semicolon and pipe by scoring how
  consistently each splits the sample lines, so prose full of commas does not beat
  the real separator. Parsing follows RFC 4180, and the UTF-8 BOM Excel writes on
  every CSV export is stripped.
  
  Values are typed the way a spreadsheet would type them, except where that would
  destroy data: leading-zero identifiers, integers beyond float precision, and
  formatted numbers like `1,000` or `$5` stay text.

## 0.1.1

### Patch Changes

- Fix structured references losing their meaning on import.
  
  Two bugs, both in current-row (`[@Column]`) references:
  
  `[@Column]` was being stripped to `[Column]`. The `@` was read as the implicit
  intersection operator, which the converter drops on purpose. That turned a
  reference to one row into a reference to the whole column — a formula that still
  evaluates, and quietly gives the wrong number.
  
  `Table[[#This Row],[Column]]` was passed through unchanged and does not parse.
  Excel stores that spelling for every calculated column while only ever
  displaying `[@Column]`, so it is the form an importer actually meets. It is now
  rewritten to the `@` form, dropping the table name when the formula sits inside
  the table it names — which reproduces exactly the text the author typed in Excel.
