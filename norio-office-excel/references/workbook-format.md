# Workbook Format

The public workbook snapshot is the editor's internal state, not a user-facing file format. Validate it with `validateExcelWorkbook` before applying or exporting it.

- Coordinates are one-based.
- Cell record keys are `row:col` and must match the cell's `row` and `col` fields.
- Sheet IDs and names are unique; `activeSheetId` must exist.
- Preserve merges, dimensions, hidden state, groups, validations, conditional formats, tables, pivots, charts, images, named ranges, and watermark fields unless the requested operation explicitly changes them.
- Unknown fields are rejected at the public boundary.
- Values remain strings in the editor snapshot; AI patch input uses explicit typed values so conversion is deliberate.

Use `schemas/workbook.schema.json` for structural generation and the runtime validator for final authority.
