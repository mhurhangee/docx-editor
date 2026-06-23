/**
 * Step 1 (footnote-edit seam): `convertFootnoteToContent` must prefer a live
 * footnote ProseMirror doc supplied via `getFootnotePmDoc` over the footnote's
 * static `content`. When the callback is absent or returns null/undefined the
 * output is byte-identical to the static-content path (backward compatible).
 */

import { describe, expect, test } from 'bun:test';
import { convertFootnoteToContent } from '../footnoteLayout';
import type { MeasureBlocksFn } from '../footnoteLayout';
import { footnoteToProseDoc } from '../../prosemirror/conversion/toProseDoc';
import type { Footnote, Paragraph } from '../../types/document';
import type { FlowBlock, Measure, ParagraphBlock } from '../../layout-engine/types';

function paragraph(text: string): Paragraph {
  return {
    type: 'paragraph',
    content: [{ type: 'run', content: [{ type: 'text', text }] }],
  };
}

function footnote(id: number, text: string): Footnote {
  return { type: 'footnote', id, noteType: 'normal', content: [paragraph(text)] };
}

/** Collect the concatenated text of every text run in the produced blocks. */
function blocksText(blocks: FlowBlock[]): string {
  const parts: string[] = [];
  for (const block of blocks) {
    if (block.kind !== 'paragraph') continue;
    for (const run of (block as ParagraphBlock).runs) {
      if (run.kind === 'text') parts.push(run.text);
    }
  }
  return parts.join('');
}

/** Minimal measure mock: one fixed-height paragraph measure per paragraph block. */
const measureBlocks: MeasureBlocksFn = (blocks): Measure[] =>
  blocks.map(
    () =>
      ({
        kind: 'paragraph',
        totalHeight: 10,
        lines: [{ height: 10 }],
      }) as unknown as Measure
  );

describe('convertFootnoteToContent — live PM doc seam', () => {
  test('uses the live doc text when getFootnotePmDoc returns a doc', () => {
    const fn = footnote(7, 'static text');
    const liveDoc = footnoteToProseDoc([paragraph('live text')]);

    const content = convertFootnoteToContent(fn, 1, 200, {
      measureBlocks,
      getFootnotePmDoc: (footnoteId) => (footnoteId === 7 ? liveDoc : null),
    });

    expect(content.id).toBe(7);
    expect(blocksText(content.blocks)).toContain('live text');
    expect(blocksText(content.blocks)).not.toContain('static text');
  });

  test('falls back to footnote.content when getFootnotePmDoc returns null', () => {
    const fn = footnote(7, 'static text');

    const withNull = convertFootnoteToContent(fn, 1, 200, {
      measureBlocks,
      getFootnotePmDoc: () => null,
    });
    const withoutCallback = convertFootnoteToContent(fn, 1, 200, { measureBlocks });

    expect(blocksText(withNull.blocks)).toContain('static text');
    // Absent callback behaves identically to the null-returning callback: both
    // re-parse `footnote.content`. (Block ids come from a module-level counter,
    // so compare the meaningful payload — text + block kinds — not raw ids.)
    expect(blocksText(withNull.blocks)).toBe(blocksText(withoutCallback.blocks));
    expect(withNull.blocks.map((b) => b.kind)).toEqual(withoutCallback.blocks.map((b) => b.kind));
  });
});
