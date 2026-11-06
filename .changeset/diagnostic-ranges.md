---
"@ricsam/excel-to-formula-engine": minor
---

Diagnostics about a conditional formatting rule now carry `ranges`: the A1
ranges the rule applied to, with `$` markers removed.

Until now such a diagnostic only named its sheet, so a report could say a rule
was skipped but not which cells lost their formatting.
