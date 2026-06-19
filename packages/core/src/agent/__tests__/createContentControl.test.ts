/**
 * createContentControl — wrap existing content in a new content control.
 *
 * Inline (text-span) wrapping is the must-have: it splits runs at the span
 * boundaries, preserves formatting, includes interior fields/tabs wholesale,
 * and errors on overlap / a boundary inside a non-run item. Block wrapping is
 * the symmetric addition and rejects table-cell targets.
 */

import { describe, expect, test } from 'bun:test';

import { createContentControl, ContentControlCreateError } from '../createContentControl';
import {
  findContentControl,
  findContentControls,
  type ContentControlAddress,
} from '../contentControls';
import { setContentControlValue } from '../contentControlValues';
import { getParagraphText } from '../text-utils';
import { parseDocx } from '../../docx/parser';
import { createDocx } from '../../docx/rezip';
import type { Document } from '../../types/document';

/* eslint-disable @typescript-eslint/no-explicit-any */
const mkRun = (text: string, formatting?: Record<string, unknown>): any => ({
  type: 'run',
  content: [{ type: 'text', text }],
  ...(formatting ? { formatting } : {}),
});
const mkPara = (paraId: string | undefined, ...content: any[]): any => ({
  type: 'paragraph',
  ...(paraId ? { paraId } : {}),
  content,
});
const mkBodyDoc = (content: any[]): Document =>
  ({ package: { document: { content } } }) as unknown as Document;
const bodyBlocks = (doc: Document): any[] => doc.package.document.content as any[];
const blockAddr = (...index: number[]): ContentControlAddress => ({
  location: { part: 'body' },
  steps: index.map((i) => ({ kind: 'block', index: i })),
});
/* eslint-enable @typescript-eslint/no-explicit-any */

describe('createContentControl — inline text wrap', () => {
  test('wraps a single-run span located by paraId, returns the control info', () => {
    const doc = mkBodyDoc([mkPara('P1', mkRun('Dear '), mkRun('CLIENT'), mkRun(' party'))]);
    const { doc: next, control } = createContentControl(
      doc,
      { kind: 'text', paraId: 'P1', text: 'CLIENT' },
      { tag: 'client', alias: 'Client' }
    );

    expect(control.kind).toBe('inline');
    expect(control.tag).toBe('client');
    expect(control.alias).toBe('Client');
    expect(control.id).toBeGreaterThan(0);
    expect(control.text).toBe('CLIENT');
    // resolvable by tag immediately
    expect(findContentControl(next, { tag: 'client' })!.text).toBe('CLIENT');

    // paragraph is now: run "Dear " | sdt(CLIENT) | run " party"
    const para = bodyBlocks(next)[0];
    expect(para.content.map((n: { type: string }) => n.type)).toEqual(['run', 'inlineSdt', 'run']);
    expect(para.content[0].content[0].text).toBe('Dear ');
    expect(para.content[2].content[0].text).toBe(' party');
  });

  test('wraps a span crossing run boundaries, preserving run formatting', () => {
    const doc = mkBodyDoc([
      mkPara(
        'P',
        mkRun('Hello ', { bold: true }),
        mkRun('brave '),
        mkRun('new world', { italic: true })
      ),
    ]);
    const { doc: next } = createContentControl(
      doc,
      { kind: 'text', paraId: 'P', text: 'brave new' },
      { tag: 'span' }
    );
    expect(findContentControl(next, { tag: 'span' })!.text).toBe('brave new');

    const para = bodyBlocks(next)[0];
    // Hello (bold, before) | sdt[ "brave ", "new" ] | " world" (italic, after)
    expect(para.content[0].formatting).toEqual({ bold: true });
    const sdt = para.content.find((n: { type: string }) => n.type === 'inlineSdt');
    expect(sdt.content.map((r: { content: { text: string }[] }) => r.content[0].text)).toEqual([
      'brave ',
      'new',
    ]);
    // the split "new" run keeps its source formatting
    expect(sdt.content[1].formatting).toEqual({ italic: true });
    // trailing " world" run also keeps the italic
    const after = para.content[para.content.length - 1];
    expect(after.content[0].text).toBe(' world');
    expect(after.formatting).toEqual({ italic: true });
  });

  test('includes an interior field wholesale in the wrapped span', () => {
    const field = {
      type: 'simpleField',
      instruction: 'TITLE',
      content: [{ type: 'run', content: [{ type: 'text', text: 'X' }] }],
    };
    const doc = mkBodyDoc([
      mkPara('P', mkRun('Ref '), mkRun('AAA'), field, mkRun('BBB'), mkRun(' end')),
    ]);
    const { doc: next, control } = createContentControl(
      doc,
      { kind: 'text', paraId: 'P', text: 'AAABBB' },
      { tag: 'f' }
    );
    // the control spans AAA + <field> + BBB; its text flattens the field result.
    expect(control.text).toBe('AAAXBBB');
    const sdt = bodyBlocks(next)[0].content.find((n: { type: string }) => n.type === 'inlineSdt');
    expect(sdt.content.map((n: { type: string }) => n.type)).toEqual(['run', 'simpleField', 'run']);
  });

  test('wraps a span inside a table cell (inline; address carries a cell step)', () => {
    const doc = mkBodyDoc([
      {
        type: 'table',
        rows: [
          {
            type: 'tableRow',
            cells: [
              { type: 'tableCell', content: [mkPara('L', mkRun('Amount: '), mkRun('[VALUE]'))] },
            ],
          },
        ],
      },
    ]);
    const { doc: next, control } = createContentControl(
      doc,
      { kind: 'text', paraId: 'L', text: '[VALUE]' },
      { tag: 'amt', sdtType: 'plainText' }
    );
    expect(control.kind).toBe('inline');
    expect(control.address.steps.some((s) => s.kind === 'cell')).toBe(true);
    expect(findContentControl(next, { tag: 'amt' })!.text).toBe('[VALUE]');
  });

  test('wraps a span located by structural paragraph address', () => {
    const doc = mkBodyDoc([
      mkPara(undefined, mkRun('skip')),
      mkPara(undefined, mkRun('wrap ME here')),
    ]);
    const { doc: next } = createContentControl(
      doc,
      { kind: 'text', paragraph: blockAddr(1), text: 'ME' },
      { tag: 'm' }
    );
    expect(findContentControl(next, { tag: 'm' })!.text).toBe('ME');
  });

  test('selects the requested occurrence when the text repeats', () => {
    const doc = mkBodyDoc([mkPara('P', mkRun('x [Y] z [Y] w'))]);
    const { doc: next } = createContentControl(
      doc,
      { kind: 'text', paraId: 'P', text: '[Y]', occurrence: 2 },
      { tag: 'y2' }
    );
    const para = bodyBlocks(next)[0];
    const sdtIndex = para.content.findIndex((n: { type: string }) => n.type === 'inlineSdt');
    // the run before the control contains the FIRST [Y]; the control is the 2nd.
    expect(para.content[sdtIndex - 1].content[0].text).toContain('[Y]');
    expect(findContentControl(next, { tag: 'y2' })!.text).toBe('[Y]');
  });

  test('is pure — the input document is not mutated', () => {
    const doc = mkBodyDoc([mkPara('P', mkRun('Dear CLIENT'))]);
    createContentControl(doc, { kind: 'text', paraId: 'P', text: 'CLIENT' }, { tag: 'c' });
    const para = bodyBlocks(doc)[0];
    expect(para.content).toHaveLength(1); // still a single run, no sdt spliced in
    expect(getParagraphText(para)).toBe('Dear CLIENT');
  });

  test('auto-assigns a unique w:id above any existing control', () => {
    const existing = {
      type: 'inlineSdt',
      properties: { sdtType: 'richText', tag: 'old', id: 100 },
      content: [mkRun('OLD')],
    };
    const doc = mkBodyDoc([mkPara('P', existing, mkRun(' and NEW'))]);
    const { control } = createContentControl(
      doc,
      { kind: 'text', paraId: 'P', text: 'NEW' },
      {
        tag: 'new',
      }
    );
    expect(control.id).toBe(101);
  });
});

describe('createContentControl — errors', () => {
  const doc = () => mkBodyDoc([mkPara('P', mkRun('ab ab cd'))]);

  test('throws when the text is not found', () => {
    expect(() =>
      createContentControl(doc(), { kind: 'text', paraId: 'P', text: 'ZZZ' }, {})
    ).toThrow(ContentControlCreateError);
  });

  test('throws when the requested occurrence does not exist', () => {
    expect(() =>
      createContentControl(doc(), { kind: 'text', paraId: 'P', text: 'ab', occurrence: 3 }, {})
    ).toThrow(ContentControlCreateError);
  });

  test('throws when the paraId is not found', () => {
    expect(() =>
      createContentControl(doc(), { kind: 'text', paraId: 'NOPE', text: 'ab' }, {})
    ).toThrow(ContentControlCreateError);
  });

  test('throws when the span overlaps an existing inline control', () => {
    const existing = {
      type: 'inlineSdt',
      properties: { sdtType: 'richText', tag: 'x', id: 1 },
      content: [mkRun('MID')],
    };
    // getParagraphText ignores the control's own text → "Hello  world" (two spaces).
    const d = mkBodyDoc([mkPara('P', mkRun('Hello '), existing, mkRun(' world'))]);
    expect(() =>
      createContentControl(d, { kind: 'text', paraId: 'P', text: 'Hello  world' }, { tag: 'over' })
    ).toThrow(ContentControlCreateError);
  });

  test('throws when a boundary falls inside a hyperlink', () => {
    const link = { type: 'hyperlink', href: 'https://x', children: [mkRun('LINK')] };
    const d = mkBodyDoc([mkPara('P', mkRun('see '), link, mkRun(' here'))]);
    expect(() =>
      createContentControl(d, { kind: 'text', paraId: 'P', text: 'LIN' }, { tag: 'b' })
    ).toThrow(ContentControlCreateError);
  });

  test('throws when the span encloses content that cannot live inside an inline control', () => {
    // A bookmark cannot sit inside an inline SDT; a span that swallows one is rejected.
    const bookmark = { type: 'bookmarkStart', id: 1, name: 'bm' };
    const d = mkBodyDoc([mkPara('P', mkRun('a '), mkRun('XY'), bookmark, mkRun('Z'), mkRun(' b'))]);
    expect(() =>
      createContentControl(d, { kind: 'text', paraId: 'P', text: 'XYZ' }, { tag: 'bm' })
    ).toThrow(ContentControlCreateError);
  });

  test('rejects a caller-supplied id that collides with an existing control', () => {
    const existing = {
      type: 'inlineSdt',
      properties: { sdtType: 'richText', tag: 'old', id: 7 },
      content: [mkRun('OLD')],
    };
    const d = mkBodyDoc([mkPara('P', existing, mkRun(' NEW'))]);
    expect(() =>
      createContentControl(d, { kind: 'text', paraId: 'P', text: 'NEW' }, { tag: 'n', id: 7 })
    ).toThrow(ContentControlCreateError);
  });

  test('with { force: true } a colliding id is auto-reassigned and the new control is returned', () => {
    const existing = {
      type: 'inlineSdt',
      properties: { sdtType: 'richText', tag: 'old', id: 7 },
      content: [mkRun('OLD')],
    };
    const d = mkBodyDoc([mkPara('P', existing, mkRun(' NEW'))]);
    const { doc: next, control } = createContentControl(
      d,
      { kind: 'text', paraId: 'P', text: 'NEW' },
      { tag: 'n', id: 7 },
      { force: true }
    );
    // the returned info is the NEW control (tag 'n'), not the pre-existing id-7 one
    expect(control.tag).toBe('n');
    expect(control.id).not.toBe(7); // reassigned to a fresh unique id
    expect(control.text).toBe('NEW');
    // both controls coexist, each resolvable by its own tag
    expect(findContentControl(next, { tag: 'old' })!.text).toBe('OLD');
    expect(findContentControl(next, { tag: 'n' })!.text).toBe('NEW');
  });
});

describe('createContentControl — block wrap', () => {
  test('wraps a contiguous block range in a block control', () => {
    const doc = mkBodyDoc([
      mkPara(undefined, mkRun('A')),
      mkPara(undefined, mkRun('B')),
      mkPara(undefined, mkRun('C')),
    ]);
    const { doc: next, control } = createContentControl(
      doc,
      { kind: 'blocks', from: blockAddr(0), to: blockAddr(1) },
      { tag: 'sec', alias: 'Section' }
    );
    expect(control.kind).toBe('block');
    expect(control.text).toBe('A\nB');

    const blocks = bodyBlocks(next);
    expect(blocks).toHaveLength(2);
    expect(blocks[0].type).toBe('blockSdt');
    expect(blocks[1].type).toBe('paragraph'); // "C" untouched
    expect(findContentControl(next, { tag: 'sec' })!.text).toBe('A\nB');
  });

  test('wraps a single block when `to` is omitted', () => {
    const doc = mkBodyDoc([mkPara(undefined, mkRun('only'))]);
    const { control } = createContentControl(
      doc,
      { kind: 'blocks', from: blockAddr(0) },
      {
        tag: 's',
      }
    );
    expect(control.kind).toBe('block');
    expect(control.text).toBe('only');
  });

  test('rejects a block target inside a table cell with a clear error', () => {
    const doc = mkBodyDoc([
      {
        type: 'table',
        rows: [
          {
            type: 'tableRow',
            cells: [{ type: 'tableCell', content: [mkPara(undefined, mkRun('cell'))] }],
          },
        ],
      },
    ]);
    const cellFrom: ContentControlAddress = {
      location: { part: 'body' },
      steps: [
        { kind: 'block', index: 0 },
        { kind: 'cell', row: 0, col: 0 },
        { kind: 'block', index: 0 },
      ],
    };
    expect(() =>
      createContentControl(doc, { kind: 'blocks', from: cellFrom }, { tag: 'x' })
    ).toThrow(ContentControlCreateError);
  });

  test('rejects a range that crosses container boundaries', () => {
    const doc = mkBodyDoc([mkPara(undefined, mkRun('A')), mkPara(undefined, mkRun('B'))]);
    const nestedTo: ContentControlAddress = {
      location: { part: 'body' },
      steps: [
        { kind: 'block', index: 0 },
        { kind: 'block', index: 0 },
      ],
    };
    expect(() =>
      createContentControl(doc, { kind: 'blocks', from: blockAddr(0), to: nestedTo }, {})
    ).toThrow(ContentControlCreateError);
  });
});

describe('createContentControl — sdtPr synthesis + round-trip', () => {
  test('a date control writes its format to <w:dateFormat>, not @w:fullDate', () => {
    const doc = mkBodyDoc([mkPara('P', mkRun('Date: '), mkRun('[DATE]'))]);
    const { doc: created } = createContentControl(
      doc,
      { kind: 'text', paraId: 'P', text: '[DATE]' },
      { tag: 'eff', sdtType: 'date', dateFormat: 'MMMM d, yyyy' }
    );
    const sdt = bodyBlocks(created)[0].content.find(
      (n: { type: string }) => n.type === 'inlineSdt'
    );
    const raw: string = sdt.properties.rawPropertiesXml;
    expect(raw).toContain('<w:date><w:dateFormat w:val="MMMM d, yyyy"/></w:date>');
    expect(raw).not.toContain('w:fullDate');
    expect(raw).toContain('<w:id w:val=');
  });

  test('created control round-trips and stays resolvable by tag', async () => {
    const doc = mkBodyDoc([mkPara('P', mkRun('Hello '), mkRun('NAME'), mkRun('!'))]);
    const { doc: created } = createContentControl(
      doc,
      { kind: 'text', paraId: 'P', text: 'NAME' },
      { tag: 'who', alias: 'Who', sdtType: 'richText' }
    );
    const reparsed = await parseDocx(await createDocx(created));
    const info = findContentControl(reparsed, { tag: 'who' })!;
    expect(info).toBeDefined();
    expect(info.kind).toBe('inline');
    expect(info.alias).toBe('Who');
    expect(info.text).toBe('NAME');
    expect(findContentControls(reparsed)).toHaveLength(1);
  });

  test('a created date control can then be filled via setContentControlValue (round-trips)', async () => {
    const doc = mkBodyDoc([mkPara('P', mkRun('Effective: '), mkRun('[DATE]'))]);
    const { doc: created } = createContentControl(
      doc,
      { kind: 'text', paraId: 'P', text: '[DATE]' },
      { tag: 'eff', sdtType: 'date', dateFormat: 'MMMM d, yyyy' }
    );
    // format survives the round trip (proves the date-bug fix end to end)
    const reparsed = await parseDocx(await createDocx(created));
    expect(findContentControl(reparsed, { tag: 'eff' })!.dateFormat).toBe('MMMM d, yyyy');

    // and the typed setter fills both visible text and structured state
    const filled = setContentControlValue(
      created,
      { tag: 'eff' },
      { kind: 'date', date: '2026-06-01' }
    );
    expect(findContentControl(filled, { tag: 'eff' })!.text).toBe('June 1, 2026');
    const reparsed2 = await parseDocx(await createDocx(filled));
    expect(findContentControl(reparsed2, { tag: 'eff' })!.text).toBe('June 1, 2026');
  });

  test('a created dropDownList can be filled via setContentControlValue', async () => {
    const doc = mkBodyDoc([mkPara('P', mkRun('Type: '), mkRun('[CHOICE]'))]);
    const { doc: created } = createContentControl(
      doc,
      { kind: 'text', paraId: 'P', text: '[CHOICE]' },
      {
        tag: 'choice',
        sdtType: 'dropDownList',
        listItems: [
          { displayText: 'Goods', value: 'g' },
          { displayText: 'Services', value: 's' },
        ],
      }
    );
    const filled = setContentControlValue(
      created,
      { tag: 'choice' },
      {
        kind: 'dropdown',
        value: 's',
      }
    );
    expect(findContentControl(filled, { tag: 'choice' })!.text).toBe('Services');
    // the structured selection (w:lastValue) is populated too, not just the text
    const sdt = bodyBlocks(filled)[0].content.find((n: { type: string }) => n.type === 'inlineSdt');
    expect(sdt.properties.rawPropertiesXml).toContain('w:lastValue="s"');
    const reparsed = await parseDocx(await createDocx(filled));
    expect(findContentControl(reparsed, { tag: 'choice' })!.text).toBe('Services');
  });
});
