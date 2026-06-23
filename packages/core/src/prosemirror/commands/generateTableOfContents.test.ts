import { describe, test, expect } from 'bun:test';
import { EditorState, TextSelection, type Command, type Transaction } from 'prosemirror-state';
import { schema } from '../schema';
import { generateTOC, generateTableOfContents } from './paragraph';

/** Doc: Intro paragraph + Heading 1/2/3. */
function docWithHeadings() {
  return schema.node('doc', { defaultTabStopTwips: null, watermark: null }, [
    schema.node('paragraph', {}, [schema.text('Intro')]),
    schema.node('paragraph', { styleId: 'Heading1' }, [schema.text('One')]),
    schema.node('paragraph', { styleId: 'Heading2' }, [schema.text('Two')]),
    schema.node('paragraph', { styleId: 'Heading3' }, [schema.text('Three')]),
  ]);
}

/** Run `cmd` with the cursor at the end of the doc; return the resulting state. */
function run(cmd: Command) {
  let state = EditorState.create({ doc: docWithHeadings() });
  state = state.apply(
    state.tr.setSelection(TextSelection.near(state.doc.resolve(state.doc.content.size), -1))
  );
  const dispatch = (tr: Transaction) => {
    state = state.apply(tr);
  };
  cmd(state, dispatch);
  return state;
}

function tocEntries(state: EditorState) {
  const entries: Array<{ style: string; text: string; href: string | null }> = [];
  state.doc.descendants((n) => {
    if (n.type.name === 'paragraph' && /^TOC\d$/.test((n.attrs.styleId as string) ?? '')) {
      let href: string | null = null;
      n.descendants((c) => {
        for (const m of c.marks) if (m.type.name === 'hyperlink') href = m.attrs.href as string;
        return true;
      });
      entries.push({ style: n.attrs.styleId as string, text: n.textContent, href });
    }
    return true;
  });
  return entries;
}

function tocTitle(state: EditorState): string | null {
  let title: string | null = null;
  state.doc.descendants((n) => {
    if (n.type.name === 'paragraph' && n.attrs.styleId === 'TOCHeading') title = n.textContent;
    return true;
  });
  return title;
}

describe('generateTableOfContents — options', () => {
  test('default (no options) matches generateTOC: every heading, default title, hyperlinked', () => {
    const state = run(generateTableOfContents());
    const entries = tocEntries(state);
    expect(entries.map((e) => e.text)).toEqual(['One', 'Two', 'Three']);
    expect(entries.map((e) => e.style)).toEqual(['TOC1', 'TOC2', 'TOC3']);
    expect(tocTitle(state)).toBe('Table of Contents');
    expect(entries.every((e) => e.href?.startsWith('#_Toc'))).toBe(true);
  });

  test('bare generateTOC export is unchanged', () => {
    const entries = tocEntries(run(generateTOC));
    expect(entries.map((e) => e.text)).toEqual(['One', 'Two', 'Three']);
  });

  test('maxLevel scopes out deeper headings', () => {
    const entries = tocEntries(run(generateTableOfContents({ maxLevel: 2 })));
    expect(entries.map((e) => e.text)).toEqual(['One', 'Two']);
  });

  test('minLevel scopes out shallower headings', () => {
    const entries = tocEntries(run(generateTableOfContents({ minLevel: 2 })));
    expect(entries.map((e) => e.text)).toEqual(['Two', 'Three']);
  });

  test('custom title is used', () => {
    expect(tocTitle(run(generateTableOfContents({ title: 'Contents' })))).toBe('Contents');
  });

  test('null title omits the title paragraph', () => {
    expect(tocTitle(run(generateTableOfContents({ title: null })))).toBeNull();
  });

  test('empty-string title also omits the title paragraph', () => {
    expect(tocTitle(run(generateTableOfContents({ title: '' })))).toBeNull();
  });

  test('out-of-range levels are clamped into 1–9', () => {
    // minLevel 0 (below Heading 1) and an oversized maxLevel both clamp, so
    // every heading is still included rather than silently filtered out.
    const entries = tocEntries(run(generateTableOfContents({ minLevel: 0, maxLevel: 99 })));
    expect(entries.map((e) => e.text)).toEqual(['One', 'Two', 'Three']);
  });

  test('an inverted range (min > max) is ordered, not emptied', () => {
    const entries = tocEntries(run(generateTableOfContents({ minLevel: 3, maxLevel: 1 })));
    expect(entries.map((e) => e.text)).toEqual(['One', 'Two', 'Three']);
  });

  test('includeHyperlinks:false yields plain-text entries', () => {
    const entries = tocEntries(run(generateTableOfContents({ includeHyperlinks: false })));
    expect(entries.map((e) => e.text)).toEqual(['One', 'Two', 'Three']);
    expect(entries.every((e) => e.href === null)).toBe(true);
  });
});
