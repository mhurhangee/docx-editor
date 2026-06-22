import { describe, test, expect } from 'bun:test';
import JSZip from 'jszip';
import {
  unzipDocx,
  assertDecompressionBudget,
  MAX_ENTRIES,
  MAX_ENTRY_UNCOMPRESSED_BYTES,
  MAX_TOTAL_UNCOMPRESSED_BYTES,
} from '../unzip';

describe('assertDecompressionBudget', () => {
  test('allows entries within budget', () => {
    expect(() => assertDecompressionBudget('a.xml', 1024, 0)).not.toThrow();
    expect(() =>
      assertDecompressionBudget('a.xml', 1024, MAX_TOTAL_UNCOMPRESSED_BYTES - 2048)
    ).not.toThrow();
  });

  test('rejects a single oversized entry before inflation', () => {
    expect(() =>
      assertDecompressionBudget('bomb.xml', MAX_ENTRY_UNCOMPRESSED_BYTES + 1, 0)
    ).toThrow(/exceeds size limit/);
  });

  test('rejects when the cumulative total would be exceeded', () => {
    expect(() => assertDecompressionBudget('a.xml', 1024, MAX_TOTAL_UNCOMPRESSED_BYTES)).toThrow(
      /total uncompressed size exceeds limit/
    );
  });
});

describe('unzipDocx entry-count guard', () => {
  test('rejects an archive with too many entries', async () => {
    const zip = new JSZip();
    for (let i = 0; i <= MAX_ENTRIES; i++) {
      zip.file(`f${i}.bin`, '');
    }
    const buffer = await zip.generateAsync({ type: 'arraybuffer' });
    await expect(unzipDocx(buffer)).rejects.toThrow(/too many entries/);
  });

  test('accepts a small well-formed archive', async () => {
    const zip = new JSZip();
    zip.file(
      'word/document.xml',
      '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>'
    );
    const buffer = await zip.generateAsync({ type: 'arraybuffer' });
    const content = await unzipDocx(buffer);
    expect(content.documentXml).toContain('w:document');
  });
});
