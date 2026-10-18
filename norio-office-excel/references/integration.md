# Host Integration

Install `@norio-office/office-excel` in the host project and import the public subpaths listed in `runtime-manifest.json`. Bind the ready component instance with `createExcelAiBridge`; do not reach through Vue refs beyond `OfficeExcelPublicApi`.

Use `createExcelToolHandlers` to map the standard tool definitions to an AI runtime, MCP server, HTTP service, Electron IPC, or another host transport. Tool names may be mapped explicitly, but argument semantics must stay identical.

Register every definition returned by `getExcelToolDefinitions()`, including `excel_capture_target`. Do not copy a shortened operation schema into the host: `excel_get_capabilities().patchSchema` and the `excel_apply_patch.patch` input both expose the same strict operation-level schema.

For a cross-sheet task, capture the source and destination separately, then pass their target IDs into `createEditContext` as `sourceTargetIds` and `destinationTargetId`. A sheet-management task first captures `kind="sheet"`; an object update first captures the stable object ID. The host must never manufacture `targetRef` values.

Native NDJSON streaming is optional. A host may instead collect a normal function-call patch, run public validation, preview it, and apply it. Do not force NDJSON framing onto runtimes that already provide structured tool calls.

Set `showAiTargetButton` to show the manual locator. Listen for `ai-target-change` and `ai-target-invalidated` to synchronize a host-side range card. Clear by matching `targetId`; cleanup from an older request must not clear a newer target.

In collaboration mode, AI mutations require a server-authoritative atomic batch command with one acknowledgment and canonical update. If that capability is unavailable, expose it as unavailable and reject the write. Do not emulate authority with a client-side full snapshot replacement.

The CLI is optional and resolves the npm package from the consuming project. Importing the Skill does not install dependencies, register tools, start a model, or grant file/network access.
