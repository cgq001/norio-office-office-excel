# Patch Operations

An `ExcelAiPatch` binds `documentId`, `requestId`, `contextId`, and `baseRevision` to one or more operations. Validate it before preview and again immediately before commit. Use the exact schema returned by `excel_get_capabilities`; unknown fields and actions are rejected and the batch remains atomic.

## Common Rules

- Rows and columns are one-based. Operation column fields are absolute worksheet columns, not offsets inside the target.
- Use only opaque target references from the current context. Read sources, write targets, create destinations, sheet targets, and object targets are distinct roles.
- A matrix write must fit its authorized target. A blank `insert` anchor may expand right and down only while the complete destination is empty and in bounds.
- `clearCells.clear` is `content`, `style`, or `all`; omitted means `content`. Content clearing preserves style, style clearing preserves values/formulas, and `all` removes both.
- Merge is rejected when it would destroy non-anchor values. Cross-sheet `moveRange` is not supported. A move authorizes the complete destination rectangle implied by the source size; a one-cell insert anchor expands only across empty cells, while overwriting requires a full destination write target.
- `answer` contexts, protected sheets, expired contexts, stale revisions, and targets without the required use are rejected before mutation.
- Identical `requestId` plus identical content is idempotent; the same ID with different content is a conflict.

## Filtering

`filterRange` has five separate actions:

| Action | Required fields | Meaning |
| --- | --- | --- |
| `enable` | `targetRef` | Add filter dropdowns for the authorized range without hiding rows or replacing existing conditions. Repeating it for the same range is a no-op. |
| `apply` | `targetRef`, `col`, `filterState` | Calculate and apply one column condition. `col` is an absolute one-based worksheet column. |
| `clear-column` | `targetRef`, `col` | Remove that column's condition while preserving dropdowns and other column conditions. |
| `clear-all` | `targetRef` | Remove every condition while preserving dropdowns. |
| `disable` | `targetRef` | Remove dropdowns and all filter state. |

Clearing or applying a condition to an existing filter requires authorization covering the complete current filter range because row visibility and the other column conditions are affected across that range. This also applies when `mergeWithExisting` is `true`. Manual hidden-row state is independent and is preserved.

For `apply`, omitted `mergeWithExisting` defaults to `true`; `false` replaces other column conditions. Multiple active columns combine with logical AND to determine row visibility.

| `filterState.mode` | Required configuration | Matching rule |
| --- | --- | --- |
| `non-empty` | none | Display text is not empty. |
| `values` | `selectedValues: string[]` | Matches displayed text. `[]` matches no rows; `[""]` selects blank or empty-display cells. |
| `condition` | 1-2 `conditions`; optional `conditionJoin` (`and`/`or`, default `or`) | Text operators: `eq`, `ne`, `starts-with`, `ends-with`, `contains`, `not-contains`. Number operators: `eq`, `ne`, `gt`, `gte`, `lt`, `lte`, `between`; `between` also needs `secondaryValue`. |
| `background-color` | `selectedColor` | Matches the resolved fill color. |
| `foreground-color` | `selectedColor` | Matches the resolved text color. |
| `colors` | non-empty `selectedFillColors` and/or `selectedTextColors` | Matches either selected fill or text color. |

Colors use `#RRGGBB` or `#AARRGGBB`. Formula cells are matched using their calculated display text for value/text filters and their calculated raw value for numeric conditions. The editor computes `matchedRows`; callers cannot supply it.

Standard XLSX round-trip preserves the filter range, enabled state, value/blank selection, non-empty, text conditions, numeric conditions, and condition joins. Color criteria currently require differential-style mapping and are reported as lossy; strict XLSX export rejects them.

See `examples/filter-operations.json` for the ten canonical filter intents.

## Native Objects And Sheets

- `sheetOperation/add` requires create permission and returns the actual item in `createdSheets`. Rename/delete/activate/hide/show require a matching sheet target plus `sheetId`. Duplicate or invalid names, deleting the final sheet, hiding the final visible sheet, and activating a hidden sheet are rejected.
- `chartOperation/create` requires `sourceTargetRef` and `destinationTargetRef`. The destination chooses the worksheet and default position; optional `x`, `y`, `width`, and `height` override geometry. Update/delete require a chart object target.
- `pivotTableOperation/create` requires a read source, a create-capable sheet destination, and a strict config. It creates a result sheet and returns it in `createdSheets` plus a `pivot-table` entry in `createdObjects`. Refresh/delete require the returned pivot object target.
- Table create uses a range target; update/delete use a table object target. New tables are returned in `createdObjects`.
- Image AI operations currently support update and delete only. Image creation is not exposed.

Offline validation proves structure only. Live preview and commit still verify target existence, authorization closure, protection, locks, context lifetime, and revision.
