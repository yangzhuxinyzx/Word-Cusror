# Electron Backend Split

Phase 8 starts by splitting `electron/main.cjs` by domain instead of continuing to add new IPC handlers directly into the main file.

## Completed Slices

Completed so far:

- `files`
- `memory`
- `web/search`
- `excel`
- `fonts`
- `ai/model proxy`
- `ppt`

New modules:

- [`electron/services/files.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/services/files.cjs)
- [`electron/ipc/register-files.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/ipc/register-files.cjs)
- [`electron/services/memory.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/services/memory.cjs)
- [`electron/ipc/register-memory.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/ipc/register-memory.cjs)
- [`electron/services/web-search.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/services/web-search.cjs)
- [`electron/ipc/register-web-search.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/ipc/register-web-search.cjs)
- [`electron/services/excel.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/services/excel.cjs)
- [`electron/ipc/register-excel.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/ipc/register-excel.cjs)
- [`electron/services/fonts.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/services/fonts.cjs)
- [`electron/ipc/register-fonts.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/ipc/register-fonts.cjs)
- [`electron/services/ai-proxy.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/services/ai-proxy.cjs)
- [`electron/ipc/register-ai-proxy.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/ipc/register-ai-proxy.cjs)
- [`electron/services/ppt.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/services/ppt.cjs)
- [`electron/ipc/register-ppt.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/ipc/register-ppt.cjs)

This slice now owns:

- file URL generation
- folder selection
- recursive folder reads
- file reads
- text/binary writes
- append writes
- save dialog
- create/delete/rename
- show in folder
- file info

The memory slice now owns:

- `memory-search`
- `memory-append`
- `memory-append-session`
- `memory-status`
- `memory-status-detail`
- `memory-clear`
- `memory-rebuild-index`

The web/search slice now owns:

- Brave MCP client lifecycle
- locale/region normalization
- web/news/faq/video/discussion result transformation
- `web-search` IPC registration

The excel slice now owns:

- workbook cache management
- `.xls` to `.xlsx` conversion helpers
- formula engine helpers
- read/write/search/update handlers
- row/column/sheet mutation handlers
- merge/unmerge/sort/filter/validation/hyperlink handlers
- chart generation and insertion handlers
- `check-libreoffice` and all `excel-*` IPC registration

The fonts slice now owns:

- `fonts-list`
- `fonts-read`
- font directory validation
- font file base64 reads

The ai/model proxy slice now owns:

- `ai-chat-completions`
- `ai-cancel`
- OpenAI-compatible streaming response parsing
- Anthropic Messages streaming response parsing
- delta/reasoning delta forwarding to renderer
- in-flight request cancellation

The ppt slice now owns:

- `pptx-render-preview`
- `openrouter-gemini-ppt-prompts`
- `ppt-generate-deck`
- `ppt-edit-slides`
- PPT preview cache helpers
- DashScope / LinAPI image generation and edit helpers
- PPT asset metadata load/save helpers
- PPTX image replacement and post-process helpers

## Main Process Role

`electron/main.cjs` should move toward:

- dependency assembly
- app/window lifecycle
- service construction
- IPC registration

It should stop being the default place for:

- file IO details
- workbook logic
- PPT generation details
- provider-specific request handling

## Suggested Next Slices

The planned heavy slices for Phase 8 are now extracted.

## Notes

- Renderer API in [`electron/preload.cjs`](/C:/Users/yangz/Desktop/Github/Word-Cursor/electron/preload.cjs) remains stable during this phase.
- The new file, memory, web/search, excel, fonts, ai/model proxy, and ppt IPC registrations are wired from `main.cjs`.
- Extracted CJS modules load successfully.
- `electron/main.cjs` was rebuilt from a clean stable baseline and re-wired to the extracted `files`, `memory`, `web/search`, `excel`, `fonts`, `ai/model proxy`, and `ppt` slices.
- After the current stabilization pass, `electron/main.cjs` is roughly 32 KB and is now close to an assembly-only role.
- Remaining cleanup is no longer the planned Phase 8 slice work. What remains is legacy cleanup in untouched areas such as Document Builder helpers and unrelated runtime warnings.
