# @ricsam/excel-to-formula-engine

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
