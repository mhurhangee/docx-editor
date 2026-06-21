/**
 * Regression — anchored wps:wsp text boxes whose run is nested inside an inline
 * wrapper (`w:sdt`, `w:hyperlink`, tracked-change wrappers) must reach the
 * text-box enrichment pass, not just text boxes hanging off a direct `w:r`
 * child of the paragraph.
 *
 * Word commonly anchors a banner/notice text box from a run that lives inside a
 * content control (`w:sdt > w:sdtContent > w:r`) — e.g. a page-number SDT in a
 * footer. The old enrichment loop only visited direct `w:r` children, so the
 * text box's text was silently dropped at parse time. The fix walks runs
 * through inline wrappers; these tests pin that behaviour for body, header, and
 * footer, in both content-control and hyperlink wrappers, and guard the
 * direct-run and no-textbox cases against regression.
 */

import { describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../documentParser';
import { parseHeader, parseFooter } from '../headerFooterParser';
import { serializeDocumentBody } from '../serializer/documentSerializer';
import { serializeHeaderFooter } from '../serializer/headerFooterSerializer';
import type { BlockContent, Paragraph } from '../../types/document';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

/** A wps:wsp text box wrapped in mc:AlternateContent (Choice Requires="wps"). */
function textBoxDrawing(label: string): string {
  return `<mc:AlternateContent>
    <mc:Choice Requires="wps">
      <w:drawing>
        <wp:anchor distT="45720" distB="45720" distL="114300" distR="114300"
          simplePos="0" relativeHeight="251695104" behindDoc="0" locked="0"
          layoutInCell="1" allowOverlap="1">
          <wp:simplePos x="0" y="0"/>
          <wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>
          <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
          <wp:extent cx="1390650" cy="278130"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:wrapNone/>
          <wp:docPr id="1" name="Text Box 1"/>
          <wp:cNvGraphicFramePr/>
          <a:graphic>
            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:wsp>
                <wps:cNvSpPr txBox="1"/>
                <wps:spPr>
                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="1390650" cy="278130"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                </wps:spPr>
                <wps:txbx>
                  <w:txbxContent>
                    <w:p><w:r><w:t>${label}</w:t></w:r></w:p>
                  </w:txbxContent>
                </wps:txbx>
                <wps:bodyPr/>
              </wps:wsp>
            </a:graphicData>
          </a:graphic>
        </wp:anchor>
      </w:drawing>
    </mc:Choice>
    <mc:Fallback><w:pict/></mc:Fallback>
  </mc:AlternateContent>`;
}

/** Wrap a text-box run inside an inline content control (w:sdt). */
function sdtWrappedTextBox(label: string): string {
  return `<w:sdt>
    <w:sdtPr><w:id w:val="-1118213369"/></w:sdtPr>
    <w:sdtContent>
      <w:r>${textBoxDrawing(label)}</w:r>
    </w:sdtContent>
  </w:sdt>`;
}

/** Wrap a text-box run inside an inline hyperlink. */
function hyperlinkWrappedTextBox(label: string): string {
  return `<w:hyperlink w:anchor="x">
    <w:r>${textBoxDrawing(label)}</w:r>
  </w:hyperlink>`;
}

/** Wrap a text-box run inside a tracked-move wrapper (w:moveTo). */
function moveToWrappedTextBox(label: string): string {
  return `<w:moveTo w:id="7" w:author="A" w:date="2020-01-01T00:00:00Z">
    <w:r>${textBoxDrawing(label)}</w:r>
  </w:moveTo>`;
}

function bodyXml(paragraphInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document ${NS}><w:body><w:p>${paragraphInner}</w:p></w:body></w:document>`;
}

function hfXml(root: 'hdr' | 'ftr', paragraphInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:${root} ${NS}><w:p>${paragraphInner}</w:p></w:${root}>`;
}

/**
 * Pull every text-box's first inner text out of a list of blocks. Descends
 * into inline wrappers (inlineSdt/hyperlink/...) because the shape may attach to
 * a run nested inside a wrapper when the paragraph has no top-level run.
 */
function shapeTexts(blocks: BlockContent[]): string[] {
  const texts: string[] = [];
  function visitRunLike(items: unknown[]): void {
    for (const item of items as Array<{
      type?: string;
      content?: unknown[];
      children?: unknown[];
    }>) {
      if (item.type === 'run' && Array.isArray(item.content)) {
        for (const rc of item.content as Array<{ type?: string; shape?: unknown }>) {
          if (rc.type === 'shape') {
            const shape = rc as {
              shape: { textBody?: { content: Array<{ type: string; content: unknown[] }> } };
            };
            const innerPara = shape.shape.textBody?.content[0];
            if (innerPara?.type !== 'paragraph') throw new Error('expected paragraph');
            const innerRun = innerPara.content[0] as { type: string; content: unknown[] };
            if (innerRun?.type !== 'run') throw new Error('expected run');
            const innerText = innerRun.content[0] as { type: string; text: string };
            if (innerText?.type !== 'text') throw new Error('expected text');
            texts.push(innerText.text);
          }
        }
      } else {
        // Wrappers use `content` (sdt/fields) or `children` (hyperlink).
        const nested = item.content ?? item.children;
        if (Array.isArray(nested)) visitRunLike(nested);
      }
    }
  }
  for (const block of blocks) {
    if (block.type === 'paragraph') visitRunLike((block as Paragraph).content);
  }
  return texts;
}

describe('wrapper-nested text boxes — parse + serialize round-trip', () => {
  test('body: text box anchored from a run inside an inline w:sdt survives', () => {
    const body = parseDocumentBody(bodyXml(sdtWrappedTextBox('SDT Box')));
    expect(shapeTexts(body.content)).toEqual(['SDT Box']);
    expect(serializeDocumentBody(body)).toContain('SDT Box');
  });

  test('body: text box anchored from a run inside a w:hyperlink survives', () => {
    const body = parseDocumentBody(bodyXml(hyperlinkWrappedTextBox('Link Box')));
    expect(shapeTexts(body.content)).toEqual(['Link Box']);
    expect(serializeDocumentBody(body)).toContain('Link Box');
  });

  test('body: text box anchored from a run inside a w:moveTo wrapper survives', () => {
    const body = parseDocumentBody(bodyXml(moveToWrappedTextBox('Move Box')));
    expect(shapeTexts(body.content)).toEqual(['Move Box']);
    expect(serializeDocumentBody(body)).toContain('Move Box');
  });

  test('header: sdt-nested text box survives parse + serialize', () => {
    const header = parseHeader(hfXml('hdr', sdtWrappedTextBox('Header SDT Box')));
    expect(shapeTexts(header.content)).toEqual(['Header SDT Box']);
    expect(serializeHeaderFooter(header)).toContain('Header SDT Box');
  });

  test('footer: sdt-nested text box survives parse + serialize', () => {
    const footer = parseFooter(hfXml('ftr', sdtWrappedTextBox('Footer SDT Box')));
    expect(shapeTexts(footer.content)).toEqual(['Footer SDT Box']);
    expect(serializeHeaderFooter(footer)).toContain('Footer SDT Box');
  });

  test('footer: a leading direct-run text box plus an sdt-nested one both survive', () => {
    // A footer with a direct-run banner text box followed by an inline
    // page-number content control that also anchors a notice box.
    const inner = `<w:r>${textBoxDrawing('Direct Box')}</w:r>${sdtWrappedTextBox('Nested Box')}`;
    const footer = parseFooter(hfXml('ftr', inner));
    expect(shapeTexts(footer.content).sort()).toEqual(['Direct Box', 'Nested Box']);
    const xml = serializeHeaderFooter(footer);
    expect(xml).toContain('Direct Box');
    expect(xml).toContain('Nested Box');
  });

  test('regression: direct-run text box still round-trips', () => {
    const body = parseDocumentBody(bodyXml(`<w:r>${textBoxDrawing('Direct Box')}</w:r>`));
    expect(shapeTexts(body.content)).toEqual(['Direct Box']);
    expect(serializeDocumentBody(body)).toContain('Direct Box');
  });

  test('regression: inline w:sdt with plain runs (no text box) is unaffected', () => {
    const inner = `<w:r><w:t>before </w:t></w:r>
      <w:sdt>
        <w:sdtPr><w:id w:val="42"/></w:sdtPr>
        <w:sdtContent><w:r><w:t>inside sdt</w:t></w:r></w:sdtContent>
      </w:sdt>
      <w:r><w:t> after</w:t></w:r>`;
    const body = parseDocumentBody(bodyXml(inner));
    // No shapes injected.
    expect(shapeTexts(body.content)).toEqual([]);
    // Plain text content survives the round trip.
    const xml = serializeDocumentBody(body);
    expect(xml).toContain('before');
    expect(xml).toContain('inside sdt');
    expect(xml).toContain('after');
  });
});
