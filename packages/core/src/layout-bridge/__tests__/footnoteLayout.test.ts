import { describe, expect, test } from 'bun:test';
import {
  calculateFootnoteReservedHeights,
  collectFootnoteRefs,
  FOOTNOTE_SEPARATOR_HEIGHT,
} from '../footnoteLayout';
import type {
  FlowBlock,
  ParagraphBlock,
  TableBlock,
  TextBoxBlock,
} from '../../layout-engine/types';

function paragraphWithFootnote(id: string, footnoteId: number, pmStart: number): ParagraphBlock {
  return {
    kind: 'paragraph',
    id,
    runs: [
      {
        kind: 'text',
        text: 'x',
        footnoteRefId: footnoteId,
        pmStart,
      },
    ],
  };
}

describe('footnote layout reservation', () => {
  test('adds the shared separator height to each page reservation', () => {
    const reserved = calculateFootnoteReservedHeights(
      new Map([
        [1, [10, 11]],
        [3, [12]],
      ]),
      new Map([
        [10, { height: 14 }],
        [11, { height: 18 }],
        [12, { height: 9 }],
      ])
    );

    expect(reserved.get(1)).toBe(14 + 18 + FOOTNOTE_SEPARATOR_HEIGHT);
    expect(reserved.get(3)).toBe(9 + FOOTNOTE_SEPARATOR_HEIGHT);
  });
});

describe('collectFootnoteRefs', () => {
  test('collects refs from top-level paragraphs', () => {
    const blocks: FlowBlock[] = [
      paragraphWithFootnote('p1', 1, 10),
      paragraphWithFootnote('p2', 2, 20),
    ];

    expect(collectFootnoteRefs(blocks)).toEqual([
      { footnoteId: 1, pmPos: 10 },
      { footnoteId: 2, pmPos: 20 },
    ]);
  });

  test('recurses into table cells so cell-authored refs reach the page-reservation pass', () => {
    // Regression: previously the collector iterated only top-level blocks and
    // skipped `kind: "table"` entirely, so any footnote authored inside a
    // table cell never made it into pageFootnoteMap. The body still rendered
    // the in-line ref marker, but the per-page footnote area dropped the
    // entry — leaving readers with a dangling superscript number.
    const table: TableBlock = {
      kind: 'table',
      id: 't1',
      rows: [
        {
          id: 'r1',
          cells: [
            {
              id: 'c1',
              blocks: [paragraphWithFootnote('cell-p1', 7, 100)],
            },
            {
              id: 'c2',
              blocks: [
                {
                  kind: 'table',
                  id: 't-nested',
                  rows: [
                    {
                      id: 'r-nested',
                      cells: [
                        {
                          id: 'c-nested',
                          blocks: [paragraphWithFootnote('nested-p', 8, 200)],
                        },
                      ],
                    },
                  ],
                } as TableBlock,
              ],
            },
          ],
        },
      ],
    };

    const blocks: FlowBlock[] = [
      paragraphWithFootnote('body-p', 1, 10),
      table,
      paragraphWithFootnote('trailing-p', 2, 300),
    ];

    expect(collectFootnoteRefs(blocks)).toEqual([
      { footnoteId: 1, pmPos: 10 },
      { footnoteId: 7, pmPos: 100 },
      { footnoteId: 8, pmPos: 200 },
      { footnoteId: 2, pmPos: 300 },
    ]);
  });

  test('recurses into text-box content blocks', () => {
    const textBox: TextBoxBlock = {
      kind: 'textBox',
      id: 'tb1',
      width: 100,
      content: [paragraphWithFootnote('tb-p', 9, 50)],
    };

    expect(collectFootnoteRefs([textBox])).toEqual([{ footnoteId: 9, pmPos: 50 }]);
  });
});
