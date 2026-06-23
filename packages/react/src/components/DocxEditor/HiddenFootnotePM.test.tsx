/**
 * Focused test for HiddenFootnotePM (footnote-editing unification, Step 2).
 *
 * Mounts the component with an `activeFootnoteId`, dispatches an insert-text
 * transaction on its EditorView, and asserts the writeback path:
 *   - `Document.package.footnotes[id].content` reflects the new text
 *   - the footnote's `verbatimXml` is cleared on a doc change
 *
 * This exercises the real ProseMirror EditorView under happy-dom, the same way
 * the HF / body PM components are tested.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, test } from 'bun:test';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

import { createRef } from 'react';
import { cleanup, render } from '@testing-library/react';
import { TextSelection } from 'prosemirror-state';

import { HiddenFootnotePM, type HiddenFootnotePMRef } from './HiddenFootnotePM';
import type { Document, Footnote, Paragraph } from '@eigenpal/docx-editor-core/types/document';

afterEach(() => {
  cleanup();
});

function paragraph(text: string): Paragraph {
  return {
    type: 'paragraph',
    content: [{ type: 'run', content: [{ type: 'text', text }] }],
  };
}

function makeDocWithFootnote(fn: Footnote): Document {
  return {
    package: {
      footnotes: [fn],
      settings: {},
    },
  } as unknown as Document;
}

function plainText(content: Footnote['content']): string {
  return content
    .filter((b): b is Paragraph => b.type === 'paragraph')
    .map((p) =>
      (p.content ?? [])
        .flatMap((run) =>
          run.type === 'run'
            ? (run.content ?? []).map((t) => (t.type === 'text' ? t.text : ''))
            : []
        )
        .join('')
    )
    .join('\n');
}

describe('HiddenFootnotePM — writeback to Document.package.footnotes', () => {
  test('insert-text transaction updates footnote content and clears verbatimXml', () => {
    const footnote: Footnote = {
      type: 'footnote',
      id: 3,
      noteType: 'normal',
      content: [paragraph('hello')],
      verbatimXml: '<w:footnote/>',
    };
    const doc = makeDocWithFootnote(footnote);

    const ref = createRef<HiddenFootnotePMRef>();
    render(<HiddenFootnotePM ref={ref} activeFootnoteId={3} document={doc} />);

    const view = ref.current?.getView();
    expect(view).toBeTruthy();
    if (!view) return;

    // Doc seeded from the footnote content.
    expect(view.state.doc.textContent).toBe('hello');

    // Insert " world" at the end of the paragraph text.
    const insertPos = view.state.doc.content.size - 1; // inside the paragraph
    const tr = view.state.tr.insertText(' world', insertPos);
    tr.setSelection(TextSelection.create(tr.doc, insertPos));
    view.dispatch(tr);

    // Writeback landed on the model.
    expect(view.state.doc.textContent).toBe('hello world');
    expect(plainText(footnote.content)).toBe('hello world');
    // verbatimXml cleared once the footnote was edited.
    expect(footnote.verbatimXml).toBeUndefined();
  });

  test('null activeFootnoteId holds no view', () => {
    const footnote: Footnote = {
      type: 'footnote',
      id: 1,
      noteType: 'normal',
      content: [paragraph('x')],
    };
    const doc = makeDocWithFootnote(footnote);

    const ref = createRef<HiddenFootnotePMRef>();
    render(<HiddenFootnotePM ref={ref} activeFootnoteId={null} document={doc} />);

    expect(ref.current?.getView()).toBeNull();
  });
});
