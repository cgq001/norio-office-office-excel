# Streaming

Use UTF-8 NDJSON with one complete event per line. The decoder supports arbitrary byte chunks, CRLF, multiple lines per chunk, and multibyte characters split across chunks.

Every event includes `protocolVersion: 1`, a stable `streamId`, and contiguous `seq` values beginning at 1. Allowed event types are:

- `workbook.meta`
- `sheet.open` and `sheet.close`
- `cells.chunk`
- `cell.text.delta`
- `operation`
- `stream.end`

Creation streams use temporary `sheetRef` values. Edit streams use opaque live `targetRef` values. Do not interpret one as the other.

An identical repeated sequence event is an idempotent retransmission. A conflicting duplicate, sequence gap, unknown event, unclosed sheet, truncated final line, or event after `stream.end` invalidates the stream.

Preview snapshots are isolated and have `committed: false`. Edit-mode `cell.text.delta` events accumulate per authorized target cell. Create-workbook snapshots are sent to `previewAiWorkbook` so incremental content is visible in the editor before commit. Only a normally ended, fully validated stream may commit once. Cancellation rejects late events and preserves formal workbook state.
