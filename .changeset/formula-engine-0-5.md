---
"@ricsam/excel-to-formula-engine": minor
---

Require `@ricsam/formula-engine` 0.5.0.

Formula engine 0.5.0 implements `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `VLOOKUP`,
`IFS`, `FILTER`, `SUMPRODUCT`, `TEXT` and more, so formulas using them are no
longer reported as `unsupported-function` and evaluate after import. The peer
range was `^0.3.4`, which made package managers install a separate 0.3.x copy
of the engine for the converter alongside an app's 0.5.0, so those formulas were
still reported unsupported against the copy the app does not use.
