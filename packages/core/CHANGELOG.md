# @eigenpal/docx-editor-core

## 1.3.3

### Patch Changes

- bf748c0: Honor an explicit `w:header="0"` / `w:footer="0"` (header/footer pinned to the page edge) instead of replacing the 0 distance with the 0.5in default. The wrong default over-reserved the header band and could push content onto an extra page versus Word. Fixes #740.
- 15d4f39: Fix header content overlapping the body when a header contains a floating text box (e.g. a centered banner). The floating box is now positioned without pushing the in-flow header paragraphs below it — so a centered banner sits beside the surrounding header text and the body no longer overlaps the header on multi-page documents. Inline and top-and-bottom boxes still reserve vertical space.
- 06fa96b: Fix list-marker alignment when a list paragraph's direct indent has a `hanging` value larger than its `left` indent. The marker now hangs into the left margin to align with the surrounding text (matching Word) instead of being clamped to the content edge and shifted right. Fixes #729.
- bd704e2: Assign every paragraph a stable id when a document is opened, so block ids and `getSelectionInfo().paraId` work before the first edit. Previously a document without `w14:paraId` had null ids until you typed or added a comment. Fixes #738.
- 30df527: Honor an explicit `0` for layout offsets that were previously treated as "unset". A full-bleed page margin (`w:pgMar w:top/left="0"`) no longer snaps to the 1-inch default, and a text-wrapping image pinned flush to the text (`w:distL/distR="0"`) no longer opens a phantom 12px gap. Generalizes the #740 header/footer fix behind a shared nullish helper and a documented size-vs-offset rule, so the falsy-zero trap can't recur.

## 1.3.2

### Patch Changes

- 3bd7bf7: Plain paragraphs that reference a numbering level with `numFmt="none"` are no longer rendered with a fabricated "1." marker. Word shows these as plain text, so the editor now omits the marker while keeping genuine numbered and bulleted lists intact. Fixes #718.
- 0ded2a1: Right-to-left paragraphs now render in the correct reading order. A paragraph whose runs are marked right-to-left (`w:rtl`) but that carries no explicit bidi flag is laid out right-to-left based on its first strong character, so Hebrew and Arabic text no longer reads left-to-right. Alignment and indentation mirror to match. Fixes #719.
- 58e3a7e: Text highlight colors are restored when a document is reloaded. Custom highlight colors outside Word's named palette are saved as character shading (`w:shd`); the importer now reads that shading back into the highlight, so the background no longer disappears on reload even though it was always present in the exported file. Fixes #712.

## 1.3.1

### Patch Changes

- 3fe9c57: Share the layout pipeline across the React and Vue adapters. The Vue editor now renders multi-column section layouts with correct per-section column widths, coalesces a burst of keystrokes into one layout pass per frame, and no longer scrolls the page when you edit. React behavior is unchanged.
- d100115: Fix blank render on documents whose header contains a page-anchored letterhead. The body now clears the header/footer based on in-flow content only, so anchored shapes and text boxes (which Word positions on the page) no longer push the body off the page. Fixes #705.
- db75f4f: Fix dense footnote layout so split-paragraph references reserve space on the correct page.
- 66cf3a8: Share the React/Vue editor orchestration through core so both adapters stay in lockstep. Vue gains three behaviors it was missing: multi-cell selection highlighting, drag-to-edge auto-scroll while selecting, and correct comment/tracked-change ID allocation (IDs are no longer reused after a delete and no longer collide across the comment/revision space). Vue selection rectangles now also cover tab stops and hyperlink text. No public API changes.

## 1.3.0

### Minor Changes

- 5e51a9b: Fix the caret, drag-selection highlight, and table cell-selection highlight appearing in the header while editing the footer. The active header/footer is now resolved per section, so they render in the region being edited. The header/footer caret also stays glued to the text while scrolling instead of drifting away. The hovered region shows a text cursor in edit mode, and the inactive region shows a normal arrow. Fixes #671

  The `@public` `computeHfCaretRectFromView` and `computeHfSelectionRectsFromView` (exported from `@eigenpal/docx-editor-core/layout-bridge`) now take a required `section: 'header' | 'footer'` argument.

- 1be9cf5: Edit and track-change footnote and endnote bodies.

  Note bodies are now serialized on save, so edits and tracked changes (`w:ins` /
  `w:del`) inside footnotes and endnotes persist instead of being dropped — the
  run model preserves the separator markers and the in-body auto-number marks, and
  `repackDocx` writes `word/footnotes.xml` / `word/endnotes.xml` from the model.
  `DocxReviewer.getChanges()` gains `includeFootnotes` / `includeEndnotes` options;
  when set, tracked changes inside note bodies are reported with `noteId` /
  `noteType`.

- 0f3eb97: Fix watermark fidelity when saving to OOXML. Picture watermarks applied across a document's headers now bind each header to its own image relationship (previously the same relationship id was reused across header parts, which could break the image on title or even pages). Watermarks now also appear on title pages and even pages by creating the first/even header parts a section displays but lacks, without disturbing existing header inheritance. Picture watermarks keep the image's aspect ratio instead of being forced into a square.
- eaa6f7f: Add MS Word–style watermarks. Watermarks in opened documents now render behind the body content on every page, and a new Insert → Watermark dialog lets you add, edit, or remove text watermarks (preset or custom, with font, color, diagonal/horizontal layout, and semitransparent options) and picture watermarks (with scale and washout). Watermarks round-trip back to valid OOXML so Word shows the same result.

### Patch Changes

- 15966fc: Stop squashing anchored images that sit near the right edge of the page. A floating image positioned so its width reaches into the page margin (e.g. a logo flush to the top-right) was being capped to the remaining content width by the global `img { max-width: 100% }` reset and then stretched against its fixed height. Painted floating images now keep their exact OOXML size.
- 2003cec: Honor an anchored text box's horizontal position in headers and footers. A text box anchored centered relative to the page (e.g. a "For Internal Use" classification banner) now renders centered instead of pinned to the left.
- cb5f622: Preserve mid-body section breaks (`w:pPr/w:sectPr`) on headless roundtrip. A parseDocx → repackDocx roundtrip no longer collapses a multi-section document down to its final section. Fixes #680.
- 5fcca3b: Content controls (`w:sdt`) inside footnote and endnote bodies now round-trip through the editable model instead of freezing the whole note to a verbatim copy. Notes whose only block-level construct is a content control stay fully editable; the verbatim fallback now applies only to notes carrying block-level bookmarks or `w:customXml`.
- f73706e: Stop dropping several properties on headless roundtrip. Table row-level conditional formatting (`w:trPr/w:cnfStyle`, e.g. header-row/banding context) is now serialized, matching the cell path. Explicit "off" formatting overrides also survive: a run or paragraph that cancels a style value (`<w:strike w:val="0"/>`, `<w:keepNext w:val="0"/>`, and similar for doubleStrike, smallCaps, allCaps, outline, shadow, emboss, imprint, vanish, rtl, cs, keepLines, contextualSpacing, pageBreakBefore, suppressLineNumbers, suppressAutoHyphens, bidi) previously serialized to nothing and silently re-inherited the style value.
- 0d5beed: Fix long content in a table row getting cut off / hidden instead of flowing across pages. A table cell now measures its stacked paragraphs the way it paints them — collapsing adjacent paragraph before/after spacing (like Word) instead of adding it — so the row's height matches what's rendered and page breaks land on whole lines instead of slicing a line in two. Selecting text across a table that spans a page break no longer scatters selection highlights into the gap between pages, and contextual spacing is now suppressed inside table cells. Fixes #570.
- 5b38696: Render vertically-merged table cells like Word when a table crosses a page. Merged cells now keep their column and borders on the continuation page (instead of disappearing and shifting the other cells), and a tall merged cell flows its content across the page break (the row breaks mid-content like Word, honoring `w:cantSplit`). Each fragment closes with a border on the cut edge at the page break — including the merged column when it spans the boundary — and horizontal cell borders no longer render unevenly thick due to sub-pixel positioning. Fixes #666.
- 15966fc: Render anchored text boxes with `topAndBottom` wrapping at their OOXML position instead of in the body flow. A title banner pinned to the top of the page (a shape with `wp:wrapTopAndBottom` and a page-relative vertical anchor) now sits flush at the page top with the body text flowing below it, matching Word, instead of being dropped into the text where its anchor paragraph happens to fall.
- f3d6861: Fix text selection not showing in Vue headers and footers. Selecting text while editing a header or footer now paints the highlight (the body overlay was suppressed in HF mode but the HF rects were never drawn), and double/triple-click word and paragraph selection resolves against the header/footer text instead of a body run at the same position. On multi-page documents, the caret and selection now render on the header/footer instance being edited rather than always on page one's copy. Fixes #691

## 1.2.1

### Patch Changes

- 1c2b098: Fix printing blank pages past the first few in large documents. Virtualized off-screen pages were cloned as empty shells; print now forces every page to render first. Fixes #579

## 1.2.0

### Minor Changes

- 362a65f: Make block-level content controls (`w:sdt`) editable. Block structured document tags wrapping paragraphs or tables now convert to a dedicated ProseMirror node, so their content stays editable and the control survives the full edit cycle (previously it round-tripped on save but was flattened in the editor). The control boundary is drawn around its content in the paged view, and the region remains addressable by its tag/alias.
- d791e05: Add a content-control (SDT) addressing API to the headless surface. `findContentControls`/`findContentControl` discover block-level content controls by tag, alias, id, or type and read their text plus modeled state (`showingPlaceholder`, `checked`, `dateFormat`, `listItems`, `dataBinding`); `setContentControlContent` fills a control by tag (string or block content) and `removeContentControl` deletes or unwraps one. Edits preserve the control's identity and raw properties so the document still round-trips, clear the `w:showingPlcHdr` placeholder flag when writing real content, and refuse locked controls, typed controls (dropdown/date/…), and repeating-section unwraps unless forced. Makes content controls usable as stable anchors for templates and document automation.
- d791e05: Add content-control (SDT) methods to the editor ref. `getContentControls` lists block controls in the live document (filtered by tag/alias/id/type) with their text and position; `scrollToContentControl` brings one into view; `setContentControlContent` fills a control by tag (as a normal undoable edit); `removeContentControl` deletes or unwraps one. Locked controls are refused unless forced. Paired across the React and Vue adapters.
- a60ed77: Add typed value setters for content controls. `setContentControlValue` (headless) and the `setContentControlValue` editor-ref method (React + Vue) set a dropdown selection, toggle a checkbox, or set a date by tag, updating both the visible content and the structured `w:sdtPr` state (dropdown `w:lastValue`, `w14:checked`, `w:date`'s `w:fullDate`). Validates the value against the control type and list items.
- a60ed77: Support repeating sections (`w15:repeatingSection`) with add/remove, matching Word. `addRepeatingSectionItem`/`removeRepeatingSectionItem` (headless) clone an item with fresh unique ids or drop one (keeping at least one); the editor renders ＋/✕ affordances on each repeating item in React and Vue. Items round-trip losslessly.

### Patch Changes

- e30c763: Preserve block-level content controls (`w:sdt`) on save. Block-level structured document tags wrapping paragraphs or tables were silently dropped when a document was loaded and re-saved; they now round-trip losslessly, including their tag, alias, lock, and other properties. Fixes #622
- bc67374: Fix paragraph styles on empty paragraphs and the style that follows a heading on Enter. Applying a heading style to an empty paragraph and then typing now produces styled text instead of plain body text, and the style picker shows the right state. Pressing Enter at the end of a heading now starts the next paragraph in the style's follow-on style (body text) instead of another heading.

## 1.1.0

### Minor Changes

- 9d7138e: Add a `fonts` prop on `<DocxEditor>` for declarative custom-font registration — each entry injects an `@font-face` from the URL you provide, and entries sharing a `family` register different weights. Also exposes `loadFontFromUrl`, `loadFontDefinitions`, and the `FontDefinition` type from `@eigenpal/docx-editor-core/utils`. Fixes #620.
- bf11ee8: Fix undo in suggesting mode marking an existing character as inserted. Undoing a tracked paragraph break (Enter) now only removes the break, without stamping a stray insertion on the boundary character. Raises the prosemirror-history peer dependency to >= 1.5.0. Fixes #633
- 9d7138e: Font-load failures now route through the React `onError` prop and the Vue `error` event instead of the console, so you can forward them to your own error tracker; with no subscriber attached they fall back to `console.warn`. Adds `onFontError(callback)` to `@eigenpal/docx-editor-core/utils` for non-adapter hosts.
- 42ea72d: Track structural edits as OOXML revisions in suggesting mode. Paragraph-break insert/delete, paragraph-property changes, and table row/cell insert/delete/merge are now recorded, round-tripped through DOCX, and shown in the tracked-changes sidebar (React and Vue, localized). Adds `acceptChangeById(id)` / `rejectChangeById(id)`, and `acceptAllChanges` / `rejectAllChanges` now resolve every revision type rather than inline marks only. Fixes #614.
- 137d5de: Track inserted and deleted images as real tracked changes in suggesting mode. A picture added (or removed) while suggesting now carries the insertion/deletion mark, paints with a revision outline, shows a review card, and is accepted/rejected with the rest of the change — and round-trips to `<w:ins>`/`<w:del>` like Word. The mechanism is generic to inline atom nodes, so other elements (shapes, …) plug in the same way.

### Patch Changes

- 7e77654: Track list/numbering changes made in suggesting mode so rejecting them reverts cleanly. Applying a list to a paragraph now records a tracked paragraph-property change (`w:pPrChange`, matching Word), and rejecting the suggestion removes both the typed items and the numbering instead of stranding an empty list item. Fixes #634
- 30c1931: Handle DOCX tables with fully covered vertical-merge rows without creating invalid empty table rows.
- ebb85a5: Tolerate a stray unescaped `&` in DOCX XML parts (document, headers, footers, comments) instead of failing the whole parse with "Invalid character in entity name". Stray ampersands are escaped before parsing, and any remaining parse error now includes a snippet of the bytes around the offending column.
- e5e0997: Header/footer editing now matches the body: click, drag, multi-click, selection, right-click, image select, hyperlinks, table row/column/edge resize, and PAGE/NUMPAGES field inserts all behave the same as in the document body. Fixes #468.

## 1.0.3

### Patch Changes

- 24b31a4: Numbered paragraphs whose direct `w:ind` has a first-line indent but no hanging slot (e.g. `<w:ind w:left="0" w:firstLine="720"/>`) now render the marker inline with the first body line at the firstLine position, matching Word/LibreOffice. Previously the painter wrapped the marker into a separate row above the text and the layout engine didn't reserve space for that row — the last line of the first fragment spilled below its container and the continuation fragment rendered on top of it (fixes #483).
- ec36a50: Footnote references authored inside table cells (and text boxes) are now collected by the page-reservation pass. Previously `collectFootnoteRefs` walked only top-level blocks and skipped tables entirely, so nested refs never reached `mapFootnotesToPages` and the per-page footnote area silently dropped them while the body still rendered the in-line superscript marker. Fixes #584.
- 143c31e: Numbered paragraphs that write a neutral `w:hanging="0"` direct indent now keep the numbering level's hanging indent, mirroring the fix already in place for `w:firstLine="0"`. Per ECMA-376 §17.3.1.12, both are no-op values and shouldn't suppress the level-defined indent.
- d91357e: Render text boxes in headers and footers. Headers and footers now flow through the same block-content parser as the document body, so text boxes (and bullet-glyph conversion) are parsed everywhere a Word user can place them. The header/footer page painter also now draws `textBox` and `image` blocks, which it previously measured but never painted — so a header/footer text box that only appeared in the inline editor now also shows in the page view.
- bdd7f50: Preserve numbered paragraph hanging indents when DOCX paragraphs include a neutral first-line indent override.
- 6d56181: Vue now renders documents with stacked floating objects identically to React. Previously, the Vue composable ran a simplified measurement pipeline without floating-zone awareness, so anchored images / floating textboxes / floating tables would not push body text below them in Vue. The float-extraction and per-block orchestration is now shared from `@eigenpal/docx-editor-core/layout-bridge` (`measureBlocksWithFloats`); both adapters call it with their own per-block measure callback.
- e80093d: Body text now flows around stacked floating objects correctly. Documents with a side-anchored textbox plus an image floating to the right, or with a floating table whose width fills the page, used to render body paragraphs at full content width on top of the floats, push tables to the page top, or collapse the first paragraph to a single glyph per line. All three cases now match Word's layout.

## 1.0.2

### Patch Changes

- 4e73af5: Fix paragraph text wrapping onto an extra line when a right (`end`) or center tab stop is used (for example a header with a logo, a right tab, then text).

  The line measurer and the page painter each had their own tab-stop code. The measurer ignored the stop's alignment and the left indent, and used a coarse default-tab grid, so right-tabbed content was measured too wide and wrapped even though the painter laid it out on one line. Both now share one tab-stop model (`calculateTabWidth`): the same stop grid, indent handling, and `end`/`center`/`bar` alignment, so measurement and paint agree.

## 1.0.1

### Patch Changes

- 8d60d65: Extract WPS text-box drawings wrapped in `<mc:AlternateContent>` so floating text boxes from real Word docs (org-chart cards, callouts, etc.) round-trip through the parser instead of being silently dropped. The parser now walks both the direct `<w:drawing>` child of `<w:r>` and the `<mc:Choice>` / `<mc:Fallback>` branches of an `<mc:AlternateContent>` wrapper (preferring `Choice`).
- 7806b78: Clamp floating table and image wrap margins when they exceed the content width, fixing collapsed single-glyph line layout after near-full-width floating tables. Same fix applied at both wrap-zone sites: `rectsToFloatingZones` (page paint) and the React adapter's `extractFloatingZones` (pre-measurement scan).
- a193caa: Render TOC entries with Word fidelity: preserve tabs inside `<w:hyperlink>` (dot leaders no longer collapse) and inherit the TOCx paragraph color instead of the Hyperlink character style's blue + underline. Right-aligned tabs at line edges promote the line to flex layout so trailing page numbers land flush against the right margin without canvas-vs-DOM measurement drift.

  Also adjusts hyperlink anchor styling: anchors now inherit color and underline from the wrapping span (which `applyRunStyles` already styles from `run.color` / `run.underline`). The Word-default blue + underline fallback only fires when neither is resolved on the run. Documents with hyperlinks that explicitly set a non-default color or remove the underline will now reflect that, where previously the painter overrode them.

- fe4cb94: Add per-locale subpath imports to `@eigenpal/docx-editor-i18n` so dynamic
  locale loading can code-split a single locale instead of bundling the whole
  set:

  ```ts
  // Static — bundler ships only this locale's strings
  import pl from '@eigenpal/docx-editor-i18n/pl';

  // Dynamic — splits into its own chunk, loaded on demand
  const pl = (await import('@eigenpal/docx-editor-i18n/pl')).default;
  ```

  Subpaths ship for every locale: `/en`, `/de`, `/he`, `/pl`, `/pt-BR`, `/tr`,
  `/zh-CN`. The named exports on the package root still work — pick the
  ergonomic path for static lists, the subpath for runtime locale switching.

  Also re-export `createEmptyDocument`, `createDocumentWithText`, and
  `CreateEmptyDocumentOptions` from `@eigenpal/docx-editor-react` and
  `@eigenpal/docx-editor-vue` so the common "spawn a blank editor"
  affordance no longer requires installing `-core` alongside the adapter.

  Surface `Comment`, `CommentRangeStart`, `CommentRangeEnd`,
  `TrackedChangeInfo`, `TrackedRunChange`, `Insertion`, `Deletion`,
  `MoveFrom`, `MoveTo`, and `ParagraphContent` from the main
  `@eigenpal/docx-editor-core` entry. They were already public via
  `@eigenpal/docx-editor-core/headless`; the main entry just hadn't been
  re-exporting them.

## 1.0.0

### Major Changes

- 6272b32: # 1.0.0

  First multi-package, multi-framework release. The monolithic `@eigenpal/docx-js-editor` is split into a framework-agnostic core and per-framework adapters, Vue 3 ships as a first-class adapter alongside React, and the license moves to Apache 2.0 across all packages.

  ## Package restructure (breaking)

  | Old import                                 | New import                                |
  | ------------------------------------------ | ----------------------------------------- |
  | `@eigenpal/docx-js-editor`                 | `@eigenpal/docx-editor-react`             |
  | `@eigenpal/docx-js-editor/react`           | `@eigenpal/docx-editor-react`             |
  | `@eigenpal/docx-editor-react/core`         | `@eigenpal/docx-editor-core`              |
  | `@eigenpal/docx-editor-react/headless`     | `@eigenpal/docx-editor-core/headless`     |
  | `@eigenpal/docx-editor-react/core-plugins` | `@eigenpal/docx-editor-core/core-plugins` |
  | `@eigenpal/docx-editor-react/mcp`          | `@eigenpal/docx-editor-agents/mcp`        |
  | `@eigenpal/docx-editor-react/i18n/*.json`  | `@eigenpal/docx-editor-i18n/*.json`       |

  The old `@eigenpal/docx-js-editor` package stays on 0.x for legacy maintenance — no 1.x compatibility shim ships. Framework-agnostic utilities (e.g. `createEmptyDocument`) move to core:

  ```diff
  - import { DocxEditor, createEmptyDocument } from '@eigenpal/docx-js-editor';
  + import { DocxEditor } from '@eigenpal/docx-editor-react';
  + import { createEmptyDocument } from '@eigenpal/docx-editor-core';
  ```

  ## Vue 3 adapter (`@eigenpal/docx-editor-vue`)

  The Vue package becomes a real adapter (previously a stub). Public API mirrors React:
  - `<DocxEditor>` with matching prop surface
  - `useDocxEditor` composable + `renderAsync` for the Node.js path
  - `/ui`, `/composables`, `/dialogs`, `/plugin-api`, `/styles` subpaths

  Parity gates cover insert-table, find/replace, page-setup, context menus, image overlay (resize/move/rotate/aspect-locked corners, dimension tooltip), advanced cell/row options (margins, height rule, text direction, no-wrap), menu-bar icons + shortcuts + carets, toolbar pickers, and the agent UI surface.

  ## Shared i18n package (`@eigenpal/docx-editor-i18n`)

  Locale strings move out of `@eigenpal/docx-editor-react` into a dedicated package consumed by both adapters from a single source.

  ```diff
  - import de from '@eigenpal/docx-editor-react/i18n/de.json';
  + import de from '@eigenpal/docx-editor-i18n/de.json';
  ```

  The `defaultLocale` value (English) is still re-exported from the adapter packages, unchanged.

  ## Agent UI relocation (breaking)

  `AgentPanel`, `AgentChatLog`, `AgentComposer`, `AgentSuggestionChip`, `AgentTimeline` no longer ship from `@eigenpal/docx-editor-react`. They live at:
  - `@eigenpal/docx-editor-agents/react` — React components + `useAgentChat`
  - `@eigenpal/docx-editor-agents/vue` — Vue 3 twins, plus `AIContextMenu` and `AIResponsePreview`
  - `@eigenpal/docx-editor-agents/ai-sdk/react` / `/ai-sdk/vue` — `@ai-sdk/*` adapters
  - `@eigenpal/docx-editor-agents/bridge` — React-free `createEditorBridge`, `agentTools`, `executeToolCall`, `getToolSchemas`, `createReviewerBridge`. Safe for headless / Vue / Node.

  ```diff
  - import { AgentPanel, AgentChatLog } from '@eigenpal/docx-editor-react';
  + import { AgentPanel, AgentChatLog } from '@eigenpal/docx-editor-agents/react';
  ```

  The agent components no longer call `useTranslation` directly — pass localized `*Label` props instead. `<DocxEditor>`'s built-in agent panel slot still forwards localized strings automatically.

  Accessibility polish on the agent surface: keyboard-operable resize handle, Escape-dismissable context menu, live-region chat log, WCAG AA contrast on response previews.

  ## Toolbar naming unified (breaking)

  The standalone formatting bar is `Toolbar` on both adapters. The old "classic" single-row `Toolbar` (with File/Format/Insert menus baked in) is removed — compose `EditorToolbar.MenuBar` + `EditorToolbar.Toolbar` for that layout.

  | Old (React)                    | New (React + Vue)       |
  | ------------------------------ | ----------------------- |
  | `FormattingBar`                | `Toolbar`               |
  | Classic `Toolbar` (with menus) | `EditorToolbar`         |
  | `EditorToolbar.FormattingBar`  | `EditorToolbar.Toolbar` |

  Vue: `BasicToolbar` / `FormattingBar` aliases removed; `EditorToolbar`'s `formatting-bar` slot is now `toolbar`. Vue's table border-color and cell-fill pickers now use the advanced color picker matching React. Vue `MenuDropdown`'s `showChevron` default flips from `true` to `false` — pass `:show-chevron="true"` explicitly to keep the caret.

  ## `showPrintButton` prop removed (breaking)

  Removed from `<DocxEditor>` and `<Toolbar>` on both adapters; the Vue `<Toolbar>` `print` event is gone with it. `onPrint` callback stays.

  ```diff
  - <DocxEditor showPrintButton onPrint={handlePrint} />
  + <DocxEditor onPrint={handlePrint} />
  ```

  To hide File > Print, omit `onPrint`. Programmatic print still works via `ref.current.print()` / `editorRef.value.print()`.

  ## License moves to Apache 2.0

  All published packages relicense to Apache 2.0. Notably: `@eigenpal/docx-editor-agents` was AGPL-3.0-or-later — the relicense lifts copyleft obligations on agent embedders.

### Minor Changes

- 76093f9: `@eigenpal/docx-editor-core` now ships an API Extractor snapshot for every published subpath (61 entries) under `packages/core/etc/`. CI fails on any undocumented drift to the public surface via `bun run api:check`. Adds rich TSDoc on the 21 most-imported types — `Document`, `DocumentBody`, `Paragraph`, `Run`, `Table`, `TableRow`, `TableCell`, `Image`, `Hyperlink`, `Comment`, `ColorValue`, `BorderSpec`, `ShadingProperties`, `TextFormatting`, `ParagraphFormatting`, `Style`, `Section`, `SectionProperties`, `ListLevel`, `ListRendering`, `AbstractNumbering`, `NumberingDefinitions` — each linked to its ECMA-376 reference.

  No runtime change; doc-only.

### Patch Changes

- c5125ff: Annotate every subpath barrel with `@packageDocumentation` + `@public` so API Extractor can extract them in the next phase. The exports map is unchanged; the published surface is unchanged. Doc-only.
- 348fa6b: Tag three subpath helpers as `@internal` in TSDoc: `managers/TableSelectionManager`, `prosemirror/utils/extractTrackedChanges`, `prosemirror/utils/visualLineNavigation`. The subpaths stay in `package.json` `exports` for back-compat (shipped in v1.0), but the snapshots in `etc/managers-TableSelectionManager.api.md`, `etc/prosemirror-utils-extractTrackedChanges.api.md`, and `etc/prosemirror-utils-visualLineNavigation.api.md` now mark every export `// @internal`.

  Consumers should reach for the adapter-side wrappers (`useTableSelection`, `useTrackedChanges`, `useVisualLineNavigation` in React/Vue) instead of these subpaths. The tag is a signal of intent — these subpaths are expected to move behind public surfaces in a future major.

- 0187af2: Emit consumer-friendly JSON docs at `docs/json/<pkg-slug>/<subpath>.json` for every `@public` export across the published packages. Companion to the existing `etc/<slug>.api.md` snapshots — same source of truth (API Extractor), different output shape: instead of human-readable Markdown, the JSON is structured for a docs site to render any layout it wants. Includes per-export source-link URLs into the GitHub source tree, type-reference canonical IDs for cross-page linking, and TSDoc summaries/remarks/examples parsed out of the source.

  New tooling: `bun run docs:json` regenerates, `bun run docs:check` (in CI) fails on drift. Contract documented in `CLAUDE.md` under `### Docs JSON`. No runtime change to any published package.

- 61983ca: Add `@packageDocumentation` blocks to every public subpath across the published packages, and a small post-build step (`scripts/inject-package-doc.mjs`) that re-prepends the source's head doc-block to the dist `.d.ts` after tsup runs. tsup's rollup-plugin-dts hoists transitive type imports above the file-head comment, which previously stripped the description from the published types. Consumers now see the package-level prose in their IDE hover and the API Extractor snapshots no longer flag "No @packageDocumentation comment for this package".
- b2230a3: Internal refactor: TableExtension closure split into per-domain modules under `prosemirror/extensions/nodes/TableExtension/commands/` (insert, delete, selection, borders, cellFormatting, sizing, tableStyle, helpers, activeCellPlugin). Schema-binding commands become `make*(schema)` factories called once per editor; schema-free commands become module-level `Command` constants. No public API change.
- 8836214: Stop shipping sourcemaps and declaration maps in published tarballs. They were dead weight: the `.js.map` files referenced source files that aren't in the tarball, and the `.d.ts.map` files pointed at `.ts` files consumers can't see either.

  Concrete changes:
  - `@eigenpal/docx-editor-core`: drop `sourcemap: !isProd` from both tsup builds (the build never ran with `NODE_ENV=production`, so 245 `.js.map` files / ~8.2 MB were shipping). Tarball: 2.5 MB → 0.7 MB. Unpacked: 11.0 MB → 2.7 MB.
  - `@eigenpal/docx-editor-vue`: pass `compilerOptions: { declarationMap: false }` to `vite-plugin-dts` to suppress the 63 `.d.ts.map` files.
  - `@eigenpal/docx-editor-agents`: same `declarationMap: false` for the Vue sub-build; also add the missing `sideEffects: ["*.css"]` so bundlers can tree-shake.

  Total unpacked footprint across all published packages: 14.8 MB → 6.3 MB.
