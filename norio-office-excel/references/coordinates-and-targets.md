# Coordinates And Targets

All row and column numbers are one-based. Always pair an A1 display address with its sheet ID. A1 text alone is not a stable identity.

The toolbar AI button captures the editor's current native selection or selected object and creates a stable `targetId` plus opaque `targetRef`. Exact Tables and named ranges, selected charts, and selected floating images include an object kind and stable object ID. Focus can move to a chat input without changing the frozen target. Repeated capture of the same selection or object keeps the target; a different selection replaces it.

For targets that are not the current UI selection, use `excel_capture_target`:

- `{ "kind": "sheet", "sheetId": "..." }` creates a sheet-management target.
- `{ "kind": "range", "sheetId": "...", "ranges": [...] }` captures one or more explicit ranges on current, non-current, or hidden sheets.
- `{ "kind": "object", "objectKind": "table|named-range|chart|image|pivot-table", "objectId": "...", "sheetId": "..." }` captures a stable object.

The editor validates existence, bounds, protection, revision, and allowed roles, then returns an opaque reference. Capture does not activate the sheet, change the grid selection, modify the workbook, or enter undo history. A cell/range target never authorizes sheet management; use a real `kind="sheet"` target for rename, delete, activate, hide, and show.

Use `sourceTargets` for data read to answer or derive content, `writeTargets` for authorized mutations, and `destination` for newly created output. These roles are not interchangeable:

- Selecting B2:E12 for analysis grants read access, not permission to overwrite it.
- Selecting blank G2 as an insertion anchor can yield a bounded destination beginning at G2 after capacity and collision checks.
- Creating a chart uses the source target as a live data reference and creates a separate chart object.
- Chart creation requires a read source and a create destination; the destination determines the chart's worksheet and default position.
- Pivot creation requires a read source plus a sheet destination with create permission. Refresh and delete use the returned `pivot-table` object target, not source write permission.
- Updating or deleting a Table, chart, or floating image requires its frozen object target; never identify an object by display name alone.
- Disjoint selections remain separate ranges; do not include cells between them.

If row/column identity cannot be resolved after structure changes, refresh the context. Never search for matching text and silently retarget a stale request.
