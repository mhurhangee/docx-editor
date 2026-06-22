---
'@eigenpal/docx-editor-core': patch
---

Preserve text boxes anchored inside table cells on save. A text box whose anchoring run sits in a table cell was previously dropped when a document was opened and saved, so its text disappeared; it now round-trips like a text box in the document body.
