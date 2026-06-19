---
'@eigenpal/docx-editor-core': minor
---

Add `createContentControl` to wrap existing content in a new content control, and extend `setContentControlValue` to inline controls. `createContentControl` wraps an exact text span inside a paragraph (`{ kind: 'text' }`, splitting runs at the boundaries and preserving formatting) or a contiguous block range (`{ kind: 'blocks' }`), returning a new document plus the created control's info with an auto-assigned unique `w:id`. `setContentControlValue` now sets dropdown/date/checkbox values on inline controls too — including inside table cells, and headers/footers with `{ scope: 'all' }` — updating both the visible text and the structured state. Date controls now serialize their format to `<w:dateFormat>` instead of `w:date/@w:fullDate`.
