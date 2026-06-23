import { describe, test, expect } from 'bun:test';
import { parseFootnotes, parseEndnotes } from '../footnoteParser';
import { serializeFootnotes, serializeEndnotes } from './noteSerializer';
import type { Footnote, Endnote } from '../../types/content';

const ENDNOTES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
  <w:endnote w:id="1"><w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> First endnote.</w:t></w:r></w:p></w:endnote>
  <w:endnote w:id="2"><w:p><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r><w:ins w:id="5" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"><w:r><w:t>inserted text</w:t></w:r></w:ins></w:p></w:endnote>
</w:endnotes>`;

const FOOTNOTES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> A footnote.</w:t></w:r></w:p></w:footnote>
</w:footnotes>`;

describe('noteSerializer — endnotes', () => {
  test('round-trips ids, types, and content (deep-equal through reparse)', () => {
    const original = parseEndnotes(ENDNOTES_XML);
    const xml = serializeEndnotes(original.endnotes);
    const reparsed = parseEndnotes(xml);
    expect(reparsed.endnotes).toEqual(original.endnotes);
  });

  test('preserves separator and continuationSeparator markers in output XML', () => {
    const xml = serializeEndnotes(parseEndnotes(ENDNOTES_XML).endnotes);
    expect(xml).toContain('<w:separator/>');
    expect(xml).toContain('<w:continuationSeparator/>');
    expect(xml).toContain('w:type="separator"');
    expect(xml).toContain('w:type="continuationSeparator"');
  });

  test('preserves the endnoteRef auto-number mark in output XML', () => {
    const xml = serializeEndnotes(parseEndnotes(ENDNOTES_XML).endnotes);
    expect(xml).toContain('<w:endnoteRef/>');
  });

  test('preserves tracked insertions (w:ins) inside a note body', () => {
    const xml = serializeEndnotes(parseEndnotes(ENDNOTES_XML).endnotes);
    expect(xml).toContain('<w:ins');
    expect(xml).toContain('w:author="Reviewer"');
    expect(xml).toContain('inserted text');
  });

  test('emits a well-formed w:endnotes root with namespace', () => {
    const xml = serializeEndnotes(parseEndnotes(ENDNOTES_XML).endnotes);
    expect(xml).toContain('<w:endnotes');
    expect(xml).toContain('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"');
    expect(xml.trim().endsWith('</w:endnotes>')).toBe(true);
  });

  test('preserves negative separator id distinct from continuationSeparator id=0', () => {
    const reparsed = parseEndnotes(serializeEndnotes(parseEndnotes(ENDNOTES_XML).endnotes));
    expect(reparsed.getEndnote(-1)?.noteType).toBe('separator');
    expect(reparsed.getEndnote(0)?.noteType).toBe('continuationSeparator');
  });
});

describe('noteSerializer — footnotes', () => {
  test('round-trips ids, types, and content (deep-equal through reparse)', () => {
    const original = parseFootnotes(FOOTNOTES_XML);
    const xml = serializeFootnotes(original.footnotes);
    const reparsed = parseFootnotes(xml);
    expect(reparsed.footnotes).toEqual(original.footnotes);
  });

  test('preserves separators and the footnoteRef mark in output XML', () => {
    const xml = serializeFootnotes(parseFootnotes(FOOTNOTES_XML).footnotes);
    expect(xml).toContain('<w:separator/>');
    expect(xml).toContain('<w:continuationSeparator/>');
    expect(xml).toContain('<w:footnoteRef/>');
    expect(xml).toContain('<w:footnotes');
  });
});

describe('noteSerializer — auto-insert the ref marker on edited notes', () => {
  // A note edited in the live editor (Phase 2) routes its content through the
  // body block parser, which does NOT model `w:footnoteRef`/`w:endnoteRef`. So a
  // `normal` note rebuilt from `content` can lack the leading auto-number mark.
  // The serializer must re-insert exactly one marker into the first paragraph,
  // mirroring SuperDoc's `insertFootnoteRefIntoParagraph` (skip if present).

  test('prepends a single w:footnoteRef when a normal footnote lacks one', () => {
    const fn: Footnote = {
      type: 'footnote',
      id: 3,
      noteType: 'normal',
      content: [
        {
          type: 'paragraph',
          content: [{ type: 'run', content: [{ type: 'text', text: 'Edited footnote.' }] }],
        },
      ],
    };
    const xml = serializeFootnotes([fn]);

    expect((xml.match(/<w:footnoteRef\/>/g) ?? []).length).toBe(1);
    // The marker sits in the first paragraph, before the authored text.
    expect(xml.indexOf('<w:footnoteRef/>')).toBeLessThan(xml.indexOf('Edited footnote.'));
    // Well-formed: reparses without throwing and keeps the note.
    expect(parseFootnotes(xml).getFootnote(3)?.noteType).toBe('normal');
  });

  test('does not add a second w:footnoteRef when one already exists', () => {
    const xml = serializeFootnotes(parseFootnotes(FOOTNOTES_XML).footnotes);
    expect((xml.match(/<w:footnoteRef\/>/g) ?? []).length).toBe(1);
  });

  test('prepends a single w:endnoteRef when a normal endnote lacks one', () => {
    const en: Endnote = {
      type: 'endnote',
      id: 4,
      noteType: 'normal',
      content: [
        {
          type: 'paragraph',
          content: [{ type: 'run', content: [{ type: 'text', text: 'Edited endnote.' }] }],
        },
      ],
    };
    const xml = serializeEndnotes([en]);

    expect((xml.match(/<w:endnoteRef\/>/g) ?? []).length).toBe(1);
    expect(xml.indexOf('<w:endnoteRef/>')).toBeLessThan(xml.indexOf('Edited endnote.'));
  });

  test('does not inject a marker into separator / verbatim notes', () => {
    const sep: Footnote = {
      type: 'footnote',
      id: -1,
      noteType: 'separator',
      content: [],
      verbatimXml:
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
    };
    const xml = serializeFootnotes([sep]);
    // Byte-identical verbatim passthrough; no footnoteRef injected.
    expect(xml).toContain(sep.verbatimXml!);
    expect(xml).not.toContain('<w:footnoteRef/>');
  });
});
