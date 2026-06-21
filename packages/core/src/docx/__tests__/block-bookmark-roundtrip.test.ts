import { describe, test, expect } from 'bun:test';
import { parseDocumentBody } from '../documentParser';
import { serializeDocumentBody } from '../serializer/documentSerializer';
import { parseHeaderFooter } from '../headerFooterParser';
import { serializeHeaderFooter } from '../serializer/headerFooterSerializer';
import type { DocumentBody, Paragraph, Table } from '../../types/document';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';

function body(inner: string): DocumentBody {
  return parseDocumentBody(`<w:document ${W} ${W15}><w:body>${inner}</w:body></w:document>`);
}

function bodyRoundtrip(inner: string): string {
  return serializeDocumentBody(body(inner));
}

function hf(inner: string, isHeader: boolean): string {
  const root = isHeader ? 'w:hdr' : 'w:ftr';
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><${root} ${W} ${W15}>${inner}</${root}>`;
  return serializeHeaderFooter(parseHeaderFooter(xml, isHeader));
}

function countMatches(xml: string, re: RegExp): number {
  return (xml.match(re) ?? []).length;
}

const BOOKMARK_START = /<w:bookmarkStart\b/g;
const BOOKMARK_END = /<w:bookmarkEnd\b/g;

/**
 * Block-level bookmark markers (`w:bookmarkStart`/`w:bookmarkEnd`) can sit as
 * direct children of a block container BETWEEN paragraphs/tables/SDTs (e.g.
 * `</w:p><w:bookmarkEnd/><w:p>`). The block content model only carries
 * paragraphs/tables/SDTs, so without a carrier these markers are dropped.
 * They now ride the adjacent block via leading/trailingBlockMarkers and are
 * re-emitted in their original position on save.
 */
describe('block-level bookmark markers (parsing)', () => {
  test('a bookmarkEnd between two paragraphs attaches as trailing on the preceding para', () => {
    const doc = body(`
      <w:p><w:r><w:t>one</w:t></w:r></w:p>
      <w:bookmarkEnd w:id="0"/>
      <w:p><w:r><w:t>two</w:t></w:r></w:p>
    `);

    expect(doc.content.map((b) => b.type)).toEqual(['paragraph', 'paragraph']);
    const first = doc.content[0] as Paragraph;
    const second = doc.content[1] as Paragraph;
    // The marker sits between the two paragraphs; that position is identical
    // whether it leads para two or trails para one, so it rides para two —
    // keeping contiguous start/end runs in document order.
    expect(first.trailingBlockMarkers).toBeUndefined();
    expect(second.leadingBlockMarkers).toHaveLength(1);
    expect(second.leadingBlockMarkers?.[0]).toMatchObject({ type: 'bookmarkEnd', id: 0 });
  });

  test('a bookmarkStart between two paragraphs attaches as leading on the following para', () => {
    const doc = body(`
      <w:p><w:r><w:t>one</w:t></w:r></w:p>
      <w:bookmarkStart w:id="5" w:name="anchor"/>
      <w:p><w:r><w:t>two</w:t></w:r></w:p>
    `);

    const second = doc.content[1] as Paragraph;
    // The start opens a range over following content → leading on para two.
    expect(second.leadingBlockMarkers).toHaveLength(1);
    expect(second.leadingBlockMarkers?.[0]).toMatchObject({
      type: 'bookmarkStart',
      id: 5,
      name: 'anchor',
    });
  });
});

describe('block-level bookmark markers (round-trip)', () => {
  test('(1) bookmarkEnd between two w:p round-trips in its original position', () => {
    const out = bodyRoundtrip(`
      <w:p>
        <w:bookmarkStart w:id="0" w:name="rangeAcross"/>
        <w:r><w:t>one</w:t></w:r>
      </w:p>
      <w:bookmarkEnd w:id="0"/>
      <w:p><w:r><w:t>two</w:t></w:r></w:p>
    `);

    // Exactly one start and one end, matching id, no orphans / no duplicates.
    expect(countMatches(out, BOOKMARK_START)).toBe(1);
    expect(countMatches(out, BOOKMARK_END)).toBe(1);
    expect(out).toContain('w:name="rangeAcross"');
    // Reproduces the </w:p><w:bookmarkEnd/><w:p> shape: the end sits between
    // the two paragraphs, after the first paragraph closes.
    expect(out).toMatch(/<\/w:p><w:bookmarkEnd w:id="0"\/><w:p>/);
  });

  test('(2) cross-paragraph bookmark (inline start, block-level end) survives with no orphan', () => {
    const out = bodyRoundtrip(`
      <w:p>
        <w:bookmarkStart w:id="7" w:name="x"/>
        <w:r><w:t>start here</w:t></w:r>
      </w:p>
      <w:p><w:r><w:t>middle</w:t></w:r></w:p>
      <w:bookmarkEnd w:id="7"/>
      <w:p><w:r><w:t>after</w:t></w:r></w:p>
    `);

    expect(countMatches(out, BOOKMARK_START)).toBe(1);
    expect(countMatches(out, BOOKMARK_END)).toBe(1);
    // The end trails the "middle" paragraph (the block it closed after).
    expect(out).toMatch(/middle<\/w:t><\/w:r><\/w:p><w:bookmarkEnd w:id="7"\/>/);
  });

  test('(3a) block bookmark markers in a HEADER round-trip', () => {
    const out = hf(
      `
      <w:p><w:bookmarkStart w:id="3" w:name="h"/><w:r><w:t>hdr</w:t></w:r></w:p>
      <w:bookmarkEnd w:id="3"/>
      <w:p><w:r><w:t>next</w:t></w:r></w:p>
    `,
      true
    );

    expect(countMatches(out, BOOKMARK_START)).toBe(1);
    expect(countMatches(out, BOOKMARK_END)).toBe(1);
    expect(out).toContain('w:name="h"');
  });

  test('(3b) block bookmark markers in a FOOTER round-trip', () => {
    const out = hf(
      `
      <w:p><w:bookmarkStart w:id="4" w:name="f"/><w:r><w:t>ftr</w:t></w:r></w:p>
      <w:bookmarkEnd w:id="4"/>
      <w:p><w:r><w:t>next</w:t></w:r></w:p>
    `,
      false
    );

    expect(countMatches(out, BOOKMARK_START)).toBe(1);
    expect(countMatches(out, BOOKMARK_END)).toBe(1);
    expect(out).toContain('w:name="f"');
  });

  test('(4) a block-level bookmark marker inside a table cell survives', () => {
    const out = bodyRoundtrip(`
      <w:tbl>
        <w:tr>
          <w:tc>
            <w:p><w:bookmarkStart w:id="9" w:name="c"/><w:r><w:t>a</w:t></w:r></w:p>
            <w:bookmarkEnd w:id="9"/>
            <w:p><w:r><w:t>b</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
    `);

    expect(countMatches(out, BOOKMARK_START)).toBe(1);
    expect(countMatches(out, BOOKMARK_END)).toBe(1);
    expect(out).toContain('w:name="c"');
    // The end lands between the two cell paragraphs.
    expect(out).toMatch(/<\/w:p><w:bookmarkEnd w:id="9"\/><w:p>/);
  });

  test('(4b) a table cell containing ONLY bookmark markers (no w:p) keeps them', () => {
    // A point bookmark in an otherwise-empty cell: the cell has no paragraph,
    // so parseCellContent inserts the Word-required empty paragraph. The markers
    // must ride that inserted paragraph rather than being dropped.
    const out = bodyRoundtrip(`
      <w:tbl>
        <w:tr>
          <w:tc>
            <w:bookmarkStart w:id="13" w:name="emptycell"/>
            <w:bookmarkEnd w:id="13"/>
          </w:tc>
        </w:tr>
      </w:tbl>
    `);

    expect(countMatches(out, BOOKMARK_START)).toBe(1);
    expect(countMatches(out, BOOKMARK_END)).toBe(1);
    expect(out).toContain('w:name="emptycell"');
  });

  test('(5) bookmark brackets a table: start before the table, end after it', () => {
    const doc = body(`
      <w:bookmarkStart w:id="11" w:name="t"/>
      <w:tbl>
        <w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
      <w:bookmarkEnd w:id="11"/>
    `);

    const table = doc.content.find((b) => b.type === 'table') as Table;
    // Start before the table → leading on the table; end after it → trailing.
    expect(table.leadingBlockMarkers?.[0]).toMatchObject({ type: 'bookmarkStart', id: 11 });
    expect(table.trailingBlockMarkers?.[0]).toMatchObject({ type: 'bookmarkEnd', id: 11 });

    const out = serializeDocumentBody(doc);
    expect(countMatches(out, BOOKMARK_START)).toBe(1);
    expect(countMatches(out, BOOKMARK_END)).toBe(1);
    // Start immediately precedes <w:tbl>, end immediately follows </w:tbl>.
    expect(out).toMatch(/<w:bookmarkStart w:id="11" w:name="t"\/><w:tbl>/);
    expect(out).toMatch(/<\/w:tbl><w:bookmarkEnd w:id="11"\/>/);
  });

  test('(6) markers are emitted exactly once (no duplicate emission)', () => {
    const out = bodyRoundtrip(`
      <w:p><w:r><w:t>p1</w:t></w:r></w:p>
      <w:bookmarkStart w:id="1" w:name="a"/>
      <w:bookmarkEnd w:id="1"/>
      <w:p><w:r><w:t>p2</w:t></w:r></w:p>
    `);

    // One point bookmark (start+end) between two paragraphs, each once.
    expect(countMatches(out, BOOKMARK_START)).toBe(1);
    expect(countMatches(out, BOOKMARK_END)).toBe(1);
    expect(countMatches(out, /w:name="a"/g)).toBe(1);
  });

  test('(7) a block SDT whose content is ONLY bookmark markers keeps them', () => {
    // jedrazb's repro: an sdt wrapping only <w:bookmarkStart/><w:bookmarkEnd/>.
    // parseBlockContent must insert a placeholder paragraph for the markers to
    // ride on instead of dropping them when no block was ever pushed.
    const out = bodyRoundtrip(`
      <w:sdt>
        <w:sdtPr><w:tag w:val="ctrl"/></w:sdtPr>
        <w:sdtContent>
          <w:bookmarkStart w:id="14" w:name="sdtonly"/>
          <w:bookmarkEnd w:id="14"/>
        </w:sdtContent>
      </w:sdt>
    `);

    expect(countMatches(out, BOOKMARK_START)).toBe(1);
    expect(countMatches(out, BOOKMARK_END)).toBe(1);
    expect(out).toContain('w:name="sdtonly"');
  });

  test('(7b) a body whose only block-level children are bookmark markers keeps them', () => {
    const out = bodyRoundtrip(`
      <w:bookmarkStart w:id="15" w:name="bodyonly"/>
      <w:bookmarkEnd w:id="15"/>
    `);

    expect(countMatches(out, BOOKMARK_START)).toBe(1);
    expect(countMatches(out, BOOKMARK_END)).toBe(1);
    expect(out).toContain('w:name="bodyonly"');
  });

  test('a block SDT between paragraphs carries its bracketing markers', () => {
    const out = bodyRoundtrip(`
      <w:bookmarkStart w:id="2" w:name="s"/>
      <w:sdt>
        <w:sdtPr><w:tag w:val="ctrl"/></w:sdtPr>
        <w:sdtContent><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:sdtContent>
      </w:sdt>
      <w:bookmarkEnd w:id="2"/>
    `);

    expect(countMatches(out, BOOKMARK_START)).toBe(1);
    expect(countMatches(out, BOOKMARK_END)).toBe(1);
    expect(out).toMatch(/<w:bookmarkStart w:id="2" w:name="s"\/><w:sdt>/);
    expect(out).toMatch(/<\/w:sdt><w:bookmarkEnd w:id="2"\/>/);
  });

  // Regression: real templates (EU & Global IT) cluster a run of block-level
  // bookmarkStart markers immediately followed by a run of block-level
  // bookmarkEnd markers between the SAME two blocks. Splitting the run across
  // two carriers (ends→trailing-prev, starts→leading-next) inverts order and
  // emits every end BEFORE its start → end-only orphans. The whole run must
  // ride one carrier in document order so starts still precede their ends.
  test('a cluster of starts then a cluster of ends between two blocks keeps start-before-end order', () => {
    const out = bodyRoundtrip(`
      <w:p/>
      <w:bookmarkStart w:id="1" w:name="m1"/>
      <w:bookmarkStart w:id="2" w:name="m2"/>
      <w:bookmarkStart w:id="3" w:name="m3"/>
      <w:bookmarkEnd w:id="1"/>
      <w:bookmarkEnd w:id="2"/>
      <w:bookmarkEnd w:id="3"/>
      <w:p><w:r><w:t>body</w:t></w:r></w:p>
    `);

    // Each marker emitted exactly once, no orphans/dupes.
    expect(countMatches(out, BOOKMARK_START)).toBe(3);
    expect(countMatches(out, BOOKMARK_END)).toBe(3);

    // Every bookmarkStart must precede its matching bookmarkEnd in the output.
    for (const id of [1, 2, 3]) {
      const startPos = out.indexOf(`<w:bookmarkStart w:id="${id}"`);
      const endPos = out.indexOf(`<w:bookmarkEnd w:id="${id}"/>`);
      expect(startPos).toBeGreaterThanOrEqual(0);
      expect(endPos).toBeGreaterThanOrEqual(0);
      expect(startPos).toBeLessThan(endPos);
    }
  });
});
