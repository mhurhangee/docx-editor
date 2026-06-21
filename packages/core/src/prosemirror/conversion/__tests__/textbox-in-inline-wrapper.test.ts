/**
 * CHANGE 2 regression — a text box anchored from a run NESTED inside an inline
 * wrapper (content control / hyperlink / field / tracked-change) must be pulled
 * out into a sibling `textBox` PM node, just like a top-level anchored box.
 * Before the fix `extractTextBoxesFromParagraph` only scanned top-level runs, so
 * the editor save path silently dropped any wrapper-nested text box.
 */

import { describe, expect, test } from 'bun:test';
import type {
  Document,
  Paragraph,
  ShapeContent,
  InlineSdt,
  Hyperlink,
} from '../../../types/document';
import { toProseDoc } from '../toProseDoc';

function textBoxShape(text: string): ShapeContent {
  return {
    type: 'shape',
    shape: {
      type: 'shape',
      shapeType: 'rect',
      size: { width: 3600000, height: 360000 },
      wrap: { type: 'square' },
      textBody: {
        content: [
          { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text }] }] },
        ],
      },
    },
  };
}

function wrapDoc(content: Document['package']['document']['content']): Document {
  return { package: { document: { content } } };
}

function pmNodeTypes(doc: Document): string[] {
  const pm = toProseDoc(doc);
  const types: string[] = [];
  pm.forEach((n) => types.push(n.type.name));
  return types;
}

function textBoxText(doc: Document): string[] {
  const pm = toProseDoc(doc);
  const texts: string[] = [];
  pm.forEach((n) => {
    if (n.type.name === 'textBox') texts.push(n.textContent);
  });
  return texts;
}

describe('CHANGE 2 — text box anchored inside an inline wrapper', () => {
  test('top-level anchored text box still extracts (unchanged behavior)', () => {
    const para: Paragraph = {
      type: 'paragraph',
      content: [{ type: 'run', content: [textBoxShape('TOP LEVEL BOX')] }],
    };
    expect(pmNodeTypes(wrapDoc([para]))).toContain('textBox');
    expect(textBoxText(wrapDoc([para]))).toContain('TOP LEVEL BOX');
  });

  test('text box nested inside an inline content control (w:sdt) is recovered', () => {
    const sdt: InlineSdt = {
      type: 'inlineSdt',
      properties: { sdtType: 'richText' },
      content: [{ type: 'run', content: [textBoxShape('SDT BOX')] }],
    };
    const para: Paragraph = { type: 'paragraph', content: [sdt] };
    expect(pmNodeTypes(wrapDoc([para]))).toContain('textBox');
    expect(textBoxText(wrapDoc([para]))).toContain('SDT BOX');
  });

  test('text box nested inside a hyperlink (children accessor) is recovered', () => {
    const hyperlink: Hyperlink = {
      type: 'hyperlink',
      anchor: 'x',
      children: [{ type: 'run', content: [textBoxShape('LINK BOX')] }],
    };
    const para: Paragraph = { type: 'paragraph', content: [hyperlink] };
    expect(textBoxText(wrapDoc([para]))).toContain('LINK BOX');
  });

  test('text box nested inside a tracked insertion is recovered', () => {
    const para: Paragraph = {
      type: 'paragraph',
      content: [
        {
          type: 'insertion',
          info: { id: 1, author: 'a' },
          content: [{ type: 'run', content: [textBoxShape('INS BOX')] }],
        },
      ],
    };
    expect(textBoxText(wrapDoc([para]))).toContain('INS BOX');
  });
});
