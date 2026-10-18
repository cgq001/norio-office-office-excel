---
name: norio-office-excel
description: Create standard XLSX workbooks or inspect and safely modify an active Norio Office Excel workbook through its public AI targets, protocol, and tools. Use for workbook-specific generation and edits, not for generic statistics or prose-only table discussions.
metadata:
  version: 1.0.0
  minimum-package-version: 0.6.1
---

# Norio Office Excel

Use the public `@norio-office/office-excel` AI interfaces. Never inspect Vue private state, scrape grid DOM, simulate toolbar clicks, or invent a `targetRef`.

## Choose The Mode

- For questions about selected data, create an `answer` context and read the frozen target. Do not write the workbook.
- For an in-place edit, require a valid frozen target and create a `modify` context.
- For derived tables, formulas, or charts, treat selected cells as `sourceTargets`; resolve and authorize a separate destination before writing.
- For a new file, build and validate a workbook independently, then export a real `.xlsx`. Internal JSON is not a user deliverable.

## Required Workflow

1. Call `excel_get_capabilities` and obey its permissions, limits, supported formats, and collaboration status.
2. Read the workbook outline. When the task names a non-active sheet, a hidden sheet, an object ID, or a separate output location, call `excel_capture_target` to obtain each real source, write, or destination target without changing the user's UI selection.
3. Read only the required target cells within a budget. Preserve raw values, calculated values, display text, formulas, styles, sheet IDs, and one-based coordinates.
4. Use only `targetRef` values returned for the current edit context. A visible A1 address is descriptive, not authorization.
5. Build a strict patch. Use explicit typed values and modify only requested fields. A style-only request must not rewrite values or formulas.
6. Preview before a meaningful write. A preview or `pending-authority` result is not completion.
7. Apply once. On revision conflict, target invalidation, permission change, or lock conflict, stop and refresh context instead of retrying against stale coordinates.
8. Report actual committed impact, including `createdSheets` and `createdObjects`. `changed: false` is a successful no-op, not a hidden write.

## Targets And Creation

The AI toolbar button freezes a cell, range, disjoint selection, exact Table or named range, selected chart, or selected floating image independently of DOM focus. Object targets carry stable object IDs. Ordinary selection changes do not replace the frozen target. Clicking editor content or closing the target clears the marker.

Use `sortRange`, `filterRange`, `tableOperation`, `pivotTableOperation`, `chartOperation`, and `imageOperation` for native objects and data operations. Never emulate them by rewriting unrelated cells. An `answer` context is read-only. A blank `insert` anchor may expand into a bounded matrix, but existing cells must not be overwritten.

Treat filter intents distinctly:

- "add a filter" means `filterRange/action="enable"`; it only adds the filter UI and must not hide blank rows.
- A stated value, text, number, non-empty, fill-color, or text-color criterion means `action="apply"` with a complete `filterState`.
- "clear this column's condition" means `clear-column`; "clear all conditions but keep dropdowns" means `clear-all`; "turn filtering off" means `disable`.
- `col` is the one-based absolute worksheet column. `selectedValues: []` matches no rows, while `selectedValues: [""]` selects blanks. Omitted `mergeWithExisting` defaults to `true`.
- Applying a condition to an existing filter requires a target that covers the complete existing filter range, including when `mergeWithExisting` is `true`; row visibility and other column conditions are part of the authorized impact.

Read [references/operations.md](references/operations.md) and [examples/filter-operations.json](examples/filter-operations.json) before constructing filter patches.

An empty single-cell target is an insertion anchor and may permit a bounded output region after collision checks. A selected data range does not automatically grant write permission to that range. For charts, bind the chart to the real source range instead of copying the source into detached values.

For `moveRange`, authorize the complete destination rectangle implied by the source size. A one-cell `insert` destination may expand only when every cell in that rectangle is empty; overwriting requires an explicit full destination write target.

Multi-range targets remain disjoint. Never replace them with their bounding rectangle. Reject operations such as merge or whole-range sort when the operation cannot preserve that distinction.

## Streaming

Emit UTF-8 NDJSON events only, one complete JSON event per line, with continuous `seq` values. Do not mix explanations, Markdown fences, partial formulas, or half JSON into the event stream. Treat stream previews as isolated until `stream.end`, final validation, and one atomic commit all succeed.

Read [references/streaming.md](references/streaming.md) before producing stream events. Read [references/operations.md](references/operations.md) for patch construction, and [references/coordinates-and-targets.md](references/coordinates-and-targets.md) for any modification based on a selection.

## File Handling

Use `.xlsx` for formal creation, save, download, and delivery. Do not rename OOXML bytes to `.xls`, claim macro preservation, or silently discard unsupported external connections. Resources must come through host-authorized resolvers and must not depend on temporary local paths or credentials in the delivered file.

For exact workbook fields and compatibility boundaries, read [references/workbook-format.md](references/workbook-format.md), [references/formulas.md](references/formulas.md), and [references/xlsx-format.md](references/xlsx-format.md).

For host wiring and tool mapping, read [references/integration.md](references/integration.md). The CLI is an optional developer utility; importing this Skill does not automatically execute it or connect an AI runtime.
