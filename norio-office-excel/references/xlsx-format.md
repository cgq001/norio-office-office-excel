# XLSX Format

Formal output is standard `.xlsx` with the SpreadsheetML MIME type. Use `exportXlsx` for bytes and `importXlsx` for bytes-to-snapshot conversion. These functions do not trigger a browser download.

The current runtime does not promise legacy `.xls`, macro preservation in `.xlsm`, or lossless external data connections. Inspect the returned compatibility report before persisting. In strict workflows, reject lossy output instead of silently overwriting an original file.

Image assets must be resolved from host-authorized stable references. Do not write blob URLs, temporary local paths, access tokens, or private host protocols into a formal workbook.
