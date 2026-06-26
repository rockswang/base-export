# Handoff: base-export — Feishu Base Excel Export Plugin

## Current State

Feishu Base sidebar plugin (`@lark-base-open/js-sdk@0.5.0`) that exports current table/view to Excel (xlsx) with attachment support.

**Location:** `D:\work\ai-tools\base-export`

## What Works
- Core export pipeline: read table fields/records → build xlsx via `wasm-xlsxwriter` → pack ZIP with `jszip` → browser download
- Non-attachment fields: text, select, number, date, etc.
- ExportPanel UI with progress bar and error handling

## What's Blocked

### Attachment Download — Frontend Route Confirmed Dead

All known frontend approaches fail:
1. **Direct Feishu API URL** (`/space/api/box/stream/download/all/{token}`) — iframe on localhost can't forward auth cookies cross-origin
2. **SDK `field.getAttachmentUrls(recordId)`** — `REQUEST_BIZ_ERROR` code 4
3. **SDK `table.getCellThumbnailUrls(tokens, fieldId, recordId, quality)`** — same error
4. **SDK `table.getCellAttachmentUrls(tokens, fieldId, recordId)`** — same error

Root cause: Feishu host-side function `getAttachmentTemporaryDownloadUrl` actively rejects plugin iframe requests. Host-level business restriction, not plugin permission.

### Current Mitigation
When attachment URLs can't be resolved, file name + token displayed as text in xlsx. Non-attachment export still works.

## Next Steps: Backend Download Script

- **Plugin Frontend**: Export xlsx + `attachments.json` (tokens, file names, types)
- **Python Script**: Reads tokens.json, calls Feishu Open API to download files into `attachments/` folder

### Feishu Open API Reference
- Auth: `POST https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal` with `{ app_id, app_secret }`
- Download: `GET https://open.feishu.cn/open-apis/drive/v1/medias/{file_token}/download` with `Authorization: Bearer {token}`

Attachment `token` in `IOpenAttachment` IS the drive `file_token`.

### App Credentials Needed
Create a Feishu custom app at `https://open.feishu.cn/app` for `app_id` / `app_secret`, with `drive:drive` scope. App must have access to the target Base.

### Python Environment
- Use `D:\dev\uv\ve\sys312` Python 3.12 environment
- `uv pip install requests`
- Script at `scripts/download_attachments.py`

## Relevant Files

### Source
- `src/components/ExportPanel.vue` — Main export logic
- `src/utils/xlsx-builder.js` — wasm-xlsxwriter wrapper
- `src/utils/zip-builder.js` — JSZip wrapper
- `src/utils/downloader.js` — Concurrent fetch downloader

### Config
- `package.json` — Deps: SDK, wasm-xlsxwriter, jszip, Vue 3 + Element Plus + Vite

### SDK Docs (`D:\work\tests\base-jssdk-docs\docs\zh\`)
- `api/table.md` — getCellAttachmentUrls / getCellThumbnailUrls API docs
- `api/field/attachment.md` — IAttachmentField docs
- `api/common/error-code.md` — Error codes (none is code 4)
- `changelog.md` — Version history

## Suggested Skills
- **lark-openapi-explorer** — Discover exact Open API endpoint for Base attachment download
- **xlsx** — If xlsx-builder needs changes
