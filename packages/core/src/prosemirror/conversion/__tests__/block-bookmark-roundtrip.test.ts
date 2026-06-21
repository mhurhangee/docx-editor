/**
 * Regression — block-level bookmark markers (`w:bookmarkStart`/`w:bookmarkEnd`
 * sitting BETWEEN block siblings, not inside a paragraph's runs) must survive
 * the editor round trip: Document → toProseDoc → fromProseDoc → Document.
 *
 * These markers ride as opaque attrs (`leadingBlockMarkers`/`trailingBlockMarkers`)
 * on the adjacent paragraph / table / block-SDT. Before the fix the PM
 * conversion neither read nor wrote them, so every block bookmark vanished the
 * moment a doc passed through the `<DocxEditor>` (parse → PM → serialize).
 */

import { describe, test, expect } from 'bun:test';
import type {
  Document,
  Paragraph,
  Table,
  BlockSdt,
  BookmarkStart,
  BookmarkEnd,
} from '../../../types/document';
import { toProseDoc } from '../toProseDoc';
import { fromProseDoc } from '../fromProseDoc';

function wrapDoc(content: Document['package']['document']['content']): Document {
  return { package: { document: { content } } };
}

function bmStart(id: number, name: string): BookmarkStart {
  return { type: 'bookmarkStart', id, name };
}
function bmEnd(id: number): BookmarkEnd {
  return { type: 'bookmarkEnd', id };
}

describe('block-level bookmark markers round-trip (editor save)', () => {
  test('paragraph leading + trailing markers survive PM round trip', () => {
    const para: Paragraph = {
      type: 'paragraph',
      content: [{ type: 'run', content: [{ type: 'text', text: 'Hello' }] }],
      leadingBlockMarkers: [bmStart(10, 'block_bm')],
      trailingBlockMarkers: [bmEnd(10)],
    };

    const pmDoc = toProseDoc(wrapDoc([para]));
    const out = fromProseDoc(pmDoc);

    const p = out.package.document.content.find((b) => b.type === 'paragraph') as Paragraph;
    expect(p.leadingBlockMarkers).toEqual([bmStart(10, 'block_bm')]);
    expect(p.trailingBlockMarkers).toEqual([bmEnd(10)]);
  });

  test('block markers do NOT leak into the paragraph inline content', () => {
    const para: Paragraph = {
      type: 'paragraph',
      content: [{ type: 'run', content: [{ type: 'text', text: 'Body' }] }],
      leadingBlockMarkers: [bmStart(1, 'wrap')],
      trailingBlockMarkers: [bmEnd(1)],
    };

    const out = fromProseDoc(toProseDoc(wrapDoc([para])));
    const p = out.package.document.content.find((b) => b.type === 'paragraph') as Paragraph;

    // No bookmarkStart/End should have been folded into the run content.
    const hasInlineBookmark = p.content.some(
      (c) => c.type === 'bookmarkStart' || c.type === 'bookmarkEnd'
    );
    expect(hasInlineBookmark).toBe(false);
  });

  test('table leading + trailing markers survive PM round trip', () => {
    const table: Table = {
      type: 'table',
      columnWidths: [5000],
      rows: [
        {
          type: 'tableRow',
          cells: [
            {
              type: 'tableCell',
              content: [
                {
                  type: 'paragraph',
                  content: [{ type: 'run', content: [{ type: 'text', text: 'x' }] }],
                },
              ],
            },
          ],
        },
      ],
      leadingBlockMarkers: [bmStart(20, 'table_bm')],
      trailingBlockMarkers: [bmEnd(20)],
    };

    const out = fromProseDoc(toProseDoc(wrapDoc([table])));
    const t = out.package.document.content.find((b) => b.type === 'table') as Table;
    expect(t.leadingBlockMarkers).toEqual([bmStart(20, 'table_bm')]);
    expect(t.trailingBlockMarkers).toEqual([bmEnd(20)]);
  });

  test('block-SDT leading + trailing markers survive PM round trip', () => {
    const blockSdt: BlockSdt = {
      type: 'blockSdt',
      properties: { sdtType: 'richText', tag: 'ctrl' },
      content: [
        {
          type: 'paragraph',
          content: [{ type: 'run', content: [{ type: 'text', text: 'in ctrl' }] }],
        },
      ],
      leadingBlockMarkers: [bmStart(30, 'sdt_bm')],
      trailingBlockMarkers: [bmEnd(30)],
    };

    const out = fromProseDoc(toProseDoc(wrapDoc([blockSdt])));
    const s = out.package.document.content.find((b) => b.type === 'blockSdt') as BlockSdt;
    expect(s.leadingBlockMarkers).toEqual([bmStart(30, 'sdt_bm')]);
    expect(s.trailingBlockMarkers).toEqual([bmEnd(30)]);
  });

  test('paragraph without block markers stays clean', () => {
    const para: Paragraph = {
      type: 'paragraph',
      content: [{ type: 'run', content: [{ type: 'text', text: 'plain' }] }],
    };
    const out = fromProseDoc(toProseDoc(wrapDoc([para])));
    const p = out.package.document.content.find((b) => b.type === 'paragraph') as Paragraph;
    expect(p.leadingBlockMarkers).toBeUndefined();
    expect(p.trailingBlockMarkers).toBeUndefined();
  });

  test('split markers — start on one block, end on the next — both survive', () => {
    // A bookmark spanning two paragraphs: <bm:start/><p/><p/><bm:end/>
    const first: Paragraph = {
      type: 'paragraph',
      content: [{ type: 'run', content: [{ type: 'text', text: 'first' }] }],
      leadingBlockMarkers: [bmStart(40, 'span')],
    };
    const second: Paragraph = {
      type: 'paragraph',
      content: [{ type: 'run', content: [{ type: 'text', text: 'second' }] }],
      trailingBlockMarkers: [bmEnd(40)],
    };

    const out = fromProseDoc(toProseDoc(wrapDoc([first, second])));
    const paras = out.package.document.content.filter((b) => b.type === 'paragraph') as Paragraph[];
    expect(paras[0].leadingBlockMarkers).toEqual([bmStart(40, 'span')]);
    expect(paras[0].trailingBlockMarkers).toBeUndefined();
    expect(paras[1].leadingBlockMarkers).toBeUndefined();
    expect(paras[1].trailingBlockMarkers).toEqual([bmEnd(40)]);
  });
});

/**
 * Straddling bookmarks — one end is inline, the other is a block marker (or the
 * pair spans two paragraphs inline). The PM `bookmarks` attr fabricates a
 * balanced inline pair per inline start, while block markers + lone inline ends
 * ride as separate carriers; without a rebalance an id can over- or under-emit.
 * Every id must end up with matching start/end counts document-wide.
 */
describe('straddling bookmarks rebalance (no orphan markers)', () => {
  function bookmarkCounts(doc: Document): {
    starts: Map<number, number>;
    ends: Map<number, number>;
  } {
    const starts = new Map<number, number>();
    const ends = new Map<number, number>();
    const bump = (m: Map<number, number>, id: number) => m.set(id, (m.get(id) ?? 0) + 1);
    for (const block of doc.package.document.content) {
      const wm = block as {
        leadingBlockMarkers?: (BookmarkStart | BookmarkEnd)[];
        trailingBlockMarkers?: (BookmarkStart | BookmarkEnd)[];
      };
      for (const m of [...(wm.leadingBlockMarkers ?? []), ...(wm.trailingBlockMarkers ?? [])]) {
        bump(m.type === 'bookmarkStart' ? starts : ends, m.id);
      }
      if (block.type === 'paragraph') {
        for (const c of block.content) {
          if (c.type === 'bookmarkStart') bump(starts, c.id);
          else if (c.type === 'bookmarkEnd') bump(ends, c.id);
        }
      }
    }
    return { starts, ends };
  }

  function assertBalanced(doc: Document): void {
    const { starts, ends } = bookmarkCounts(doc);
    for (const id of new Set([...starts.keys(), ...ends.keys()])) {
      expect(`id=${id} start=${starts.get(id) ?? 0}`).toBe(`id=${id} start=${ends.get(id) ?? 0}`);
    }
  }

  test('inline start + block-level end (iS+bE) stays balanced — no double end', () => {
    // Inline start fabricates a balanced pair; the carried block end is the
    // duplicate the rebalance must drop.
    const para: Paragraph = {
      type: 'paragraph',
      content: [
        { type: 'bookmarkStart', id: 5, name: 'straddle' },
        { type: 'run', content: [{ type: 'text', text: 'body' }] },
      ],
      trailingBlockMarkers: [bmEnd(5)],
    };
    const out = fromProseDoc(toProseDoc(wrapDoc([para])));
    assertBalanced(out);
  });

  test('block-level start + inline end (bS+iE) preserves both — no dropped end', () => {
    // The inline end is a "lone" end carried through PM; the block start is
    // genuine. Both must survive and balance.
    const para: Paragraph = {
      type: 'paragraph',
      leadingBlockMarkers: [bmStart(6, 'straddle2')],
      content: [
        { type: 'run', content: [{ type: 'text', text: 'body' }] },
        { type: 'bookmarkEnd', id: 6 },
      ],
    };
    const out = fromProseDoc(toProseDoc(wrapDoc([para])));
    assertBalanced(out);
    const { starts, ends } = bookmarkCounts(out);
    expect(starts.get(6)).toBe(1);
    expect(ends.get(6)).toBe(1);
  });

  test('inline start in para A + inline end in para B (relocated) stays balanced', () => {
    const first: Paragraph = {
      type: 'paragraph',
      content: [
        { type: 'bookmarkStart', id: 7, name: 'crosspar' },
        { type: 'run', content: [{ type: 'text', text: 'first' }] },
      ],
    };
    const second: Paragraph = {
      type: 'paragraph',
      content: [
        { type: 'run', content: [{ type: 'text', text: 'second' }] },
        { type: 'bookmarkEnd', id: 7 },
      ],
    };
    const out = fromProseDoc(toProseDoc(wrapDoc([first, second])));
    assertBalanced(out);
  });
});
