/**
 * Regression — a text box anchored from a run inside a TABLE CELL must survive a
 * parse → save round-trip. The text-box enrichment pass historically ran only on
 * block-level paragraphs (body / header / footer / SDT content), never on cell
 * paragraphs, so a box anchored in a table cell was dropped at parse and its text
 * vanished on save. (The serializer also had to emit a text box's content based on
 * text presence rather than `textBox` geometry, since the parser stores all boxes
 * as `rect`.)
 *
 * Fixture: e2e/fixtures/textbox-in-table-cell.docx — a fully synthetic 1-cell
 * table whose cell anchors a text box reading "Internal use only".
 */
import { describe, test, expect } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { parseDocx } from '../parser';
import { repackDocx } from '../rezip';
import type { Document, BlockContent, Paragraph } from '../../types/document';

const FIXTURE = join(import.meta.dir, '../../../../../e2e/fixtures/textbox-in-table-cell.docx');
const BOX_TEXT = 'Internal use only';

function loadFixture(): Promise<Document> {
  const buf = readFileSync(FIXTURE);
  return parseDocx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
}

/** Count text-box shapes whose text body contains `text`, walking into tables. */
function countCellTextBoxes(blocks: BlockContent[], text: string): number {
  let n = 0;
  const visitParagraph = (p: Paragraph) => {
    for (const item of p.content) {
      if (item.type !== 'run') continue;
      for (const rc of item.content) {
        if (rc.type === 'shape' && rc.shape.textBody) {
          if (JSON.stringify(rc.shape.textBody.content).includes(text)) n++;
        }
      }
    }
  };
  const walk = (bs: BlockContent[], inCell: boolean) => {
    for (const b of bs) {
      if (b.type === 'paragraph' && inCell) visitParagraph(b);
      else if (b.type === 'table')
        for (const r of b.rows) for (const c of r.cells) walk(c.content, true);
      else if (b.type === 'blockSdt') walk(b.content, inCell);
    }
  };
  walk(blocks, false);
  return n;
}

describe('text box anchored inside a table cell — round-trip', () => {
  test('parse recovers the cell-anchored text box into the model', async () => {
    const doc = await loadFixture();
    expect(countCellTextBoxes(doc.package.document.content, BOX_TEXT)).toBe(1);
  });

  test('save preserves the cell text box (parse → repack → re-parse)', async () => {
    const out = await repackDocx(await loadFixture());
    const documentXml = await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('text');
    // The text box content survives as a w:txbxContent carrying its text.
    expect(documentXml).toContain('<w:txbxContent>');
    expect(documentXml).toContain(BOX_TEXT);
    // And it re-parses back into the cell.
    const reparsed = await parseDocx(out);
    expect(countCellTextBoxes(reparsed.package.document.content, BOX_TEXT)).toBe(1);
  });
});
