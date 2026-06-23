/**
 * Footnote / Endnote Serializer
 *
 * Serializes Footnote[] → word/footnotes.xml and Endnote[] → word/endnotes.xml.
 *
 * Unlike the comment serializer (which reimplements a minimal paragraph/run
 * emitter), note bodies are serialized with the SAME `serializeBlockContent`
 * the document body uses. That is deliberate: note bodies can carry the full
 * block model — paragraphs, tables, tracked-change wrappers (`w:ins`/`w:del`),
 * fields, run/paragraph properties — and reusing the body serializer preserves
 * all of it on round-trip rather than silently flattening it.
 *
 * Separator notes (`w:type="separator"` / `w:type="continuationSeparator"`) and
 * the in-body auto-number marks (`w:footnoteRef` / `w:endnoteRef`) survive
 * because the run model now carries them (see SeparatorContent / NoteRefMark-
 * Content in types/content/run.ts); no special-casing is needed here.
 *
 * OOXML Reference:
 * - Footnotes root: w:footnotes; each note: w:footnote[@w:id][@w:type]
 * - Endnotes root:  w:endnotes;  each note: w:endnote[@w:id][@w:type]
 */

import type { Footnote, Endnote } from '../../types/content';
import type { BlockContent, Paragraph, Run } from '../../types/document';
import { serializeBlockContent } from './documentSerializer';
import { OOXML_NAMESPACES, MC_IGNORABLE } from './xmlUtils';

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * Does any run in this paragraph already carry the auto-number mark
 * (`w:footnoteRef` / `w:endnoteRef`)? Runs can be wrapped in tracked-change
 * containers (`w:ins`/`w:del`), but the leading marker is virtually always a
 * bare run on the first paragraph — a shallow scan of direct + tracked-wrapper
 * runs matches Word's documents and SuperDoc's `paragraphHasFootnoteRef`.
 */
function paragraphHasNoteRefMark(
  para: Paragraph,
  markType: 'footnoteRefMark' | 'endnoteRefMark'
): boolean {
  const runHasMark = (run: Run): boolean => run.content.some((c) => c.type === markType);
  for (const item of para.content) {
    if (item.type === 'run' && runHasMark(item)) return true;
    if ((item.type === 'insertion' || item.type === 'deletion') && Array.isArray(item.content)) {
      for (const inner of item.content) {
        if (inner.type === 'run' && runHasMark(inner)) return true;
      }
    }
  }
  return false;
}

/**
 * Build the synthetic auto-number marker run. Mirrors SuperDoc's
 * `insertFootnoteRefIntoParagraph`: an `rStyle` of the reference char style
 * plus superscript `vertAlign`, carrying the bare `w:footnoteRef`/`w:endnoteRef`
 * element. The reference char style name is the Word default
 * (`FootnoteReference` / `EndnoteReference`).
 */
function makeNoteRefMarkRun(elementName: 'footnote' | 'endnote'): Run {
  return {
    type: 'run',
    formatting: {
      styleId: elementName === 'footnote' ? 'FootnoteReference' : 'EndnoteReference',
      vertAlign: 'superscript',
    },
    content: [{ type: elementName === 'footnote' ? 'footnoteRefMark' : 'endnoteRefMark' }],
  };
}

/**
 * Ensure a `normal` note's serialized content carries the leading auto-number
 * mark. The body block parser (`parseNoteBlockContent`) does not model
 * `w:footnoteRef`/`w:endnoteRef`, so a note edited in the live editor and
 * rebuilt from `content` can lack it — Word still renders such a note, but the
 * number is gone. Prepend a marker run to the first paragraph's content when
 * none of the note's paragraphs already has one (skip-if-present, matching
 * SuperDoc). Returns a shallow-cloned content array; the input is left pure.
 */
function withNoteRefMark(
  elementName: 'footnote' | 'endnote',
  content: readonly BlockContent[]
): BlockContent[] {
  const markType = elementName === 'footnote' ? 'footnoteRefMark' : 'endnoteRefMark';
  const firstParaIndex = content.findIndex((b) => b.type === 'paragraph');

  // Already present anywhere in the note → leave untouched.
  for (const block of content) {
    if (block.type === 'paragraph' && paragraphHasNoteRefMark(block, markType)) {
      return content as BlockContent[];
    }
  }

  const markRun = makeNoteRefMarkRun(elementName);
  if (firstParaIndex === -1) {
    // No paragraph at all (degenerate edited note) — synthesize one so the
    // marker has a home, matching SuperDoc's unshift-a-paragraph fallback.
    return [{ type: 'paragraph', content: [markRun] } as Paragraph, ...(content as BlockContent[])];
  }

  const firstPara = content[firstParaIndex] as Paragraph;
  const patched: Paragraph = { ...firstPara, content: [markRun, ...firstPara.content] };
  const out = [...(content as BlockContent[])];
  out[firstParaIndex] = patched;
  return out;
}

/** Serialize one note element (footnote or endnote share identical structure). */
function serializeNote(elementName: 'footnote' | 'endnote', note: Footnote | Endnote): string {
  // Verbatim gate (#646 F3): when the note body carried a block-level construct
  // the model can't represent — note-level bookmarks or w:customXml — the parser
  // stored the original `<w:footnote>`/`<w:endnote>` bytes verbatim. Re-emit them
  // as-is rather than rebuilding from `content` (which would drop the unmodeled
  // block). This restores pre-#646 fidelity for these notes. Block-level w:sdt is
  // NOT gated: it now round-trips through the model (BlockSdt), so notes whose
  // only "exotic" content is a content control stay fully editable.
  //
  // KNOWN LIMITATION (residual edge): a note that is BOTH edited in the editor
  // AND carries a bookmark / customXml can't be both verbatim-copied and
  // re-serialized from the edited model. We prefer correctness of the
  // structure: the verbatim bytes win, so an edit to such a note does NOT
  // persist. Notes built from modeled blocks (paragraphs, tables, content
  // controls) are unaffected and remain fully editable.
  if (note.verbatimXml) {
    return note.verbatimXml;
  }

  const attrs: string[] = [];
  // Word emits w:type before w:id on separator notes; mirror that ordering.
  if (note.noteType && note.noteType !== 'normal') {
    attrs.push(`w:type="${note.noteType}"`);
  }
  attrs.push(`w:id="${note.id}"`);

  // Re-insert the leading auto-number mark for normal notes that lost it on the
  // body-parse round-trip (the parser doesn't model w:footnoteRef/w:endnoteRef).
  // Separator / continuation notes keep their own special marks and must not
  // gain an auto-number — they're already excluded by the verbatim gate above
  // and by the noteType check here.
  const content =
    !note.noteType || note.noteType === 'normal'
      ? withNoteRefMark(elementName, note.content as BlockContent[])
      : (note.content as BlockContent[]);

  const body = content.map((block) => serializeBlockContent(block)).join('');

  return `<w:${elementName} ${attrs.join(' ')}>${body}</w:${elementName}>`;
}

/**
 * Serialize footnotes to word/footnotes.xml content.
 *
 * @param footnotes - All footnotes in document order, including separator notes.
 * @returns Complete footnotes.xml string.
 */
export function serializeFootnotes(footnotes: Footnote[]): string {
  const notes = footnotes.map((fn) => serializeNote('footnote', fn)).join('');
  return `${XML_DECL}<w:footnotes ${OOXML_NAMESPACES} ${MC_IGNORABLE}>${notes}</w:footnotes>`;
}

/**
 * Serialize endnotes to word/endnotes.xml content.
 *
 * @param endnotes - All endnotes in document order, including separator notes.
 * @returns Complete endnotes.xml string.
 */
export function serializeEndnotes(endnotes: Endnote[]): string {
  const notes = endnotes.map((en) => serializeNote('endnote', en)).join('');
  return `${XML_DECL}<w:endnotes ${OOXML_NAMESPACES} ${MC_IGNORABLE}>${notes}</w:endnotes>`;
}
