# Formulas

Read raw formula text and referenced cells before modifying formulas. Formula strings begin with `=` and must arrive as complete semantic values; never stream half a formula into the formal engine.

Preserve relative, absolute, mixed, sheet-qualified, named-range, table, and whole-row/column references. Structural operations must use the editor engine so dependent references, chart sources, validations, conditional formats, and named ranges move together.

Prefer formulas that bind to the actual workbook data. Do not replace a formula with its displayed result unless the user explicitly asks to convert it to a value. A style-only request must not trigger formula rewrites or unrelated recalculation.
