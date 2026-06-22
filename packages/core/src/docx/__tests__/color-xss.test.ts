import { describe, test, expect } from 'bun:test';
import { sanitizeRgbHex, parseColorElement, resolveColorValueToHex } from '../drawingUtils';
import { parseXml } from '../xmlParser';
import type { XmlElement } from '../xmlParser';

/** Parse a DrawingML color from the children of a wrapper element. */
function parseColor(innerXml: string) {
  const doc = parseXml(
    `<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${innerXml}</a:solidFill>`
  );
  return parseColorElement((doc.elements as XmlElement[])[0]);
}

describe('sanitizeRgbHex', () => {
  test('accepts and uppercases valid 6-digit hex', () => {
    expect(sanitizeRgbHex('ff0000')).toBe('FF0000');
    expect(sanitizeRgbHex('00FF00')).toBe('00FF00');
  });

  test('accepts 3- and 8-digit hex, with or without leading #', () => {
    expect(sanitizeRgbHex('abc')).toBe('ABC');
    expect(sanitizeRgbHex('#a1b2c3d4')).toBe('A1B2C3D4');
  });

  test('rejects anything that is not bare hex', () => {
    expect(sanitizeRgbHex('red')).toBeUndefined();
    expect(sanitizeRgbHex('')).toBeUndefined();
    expect(sanitizeRgbHex(null)).toBeUndefined();
    expect(sanitizeRgbHex(undefined)).toBeUndefined();
    // The XSS payload the audit flagged: breaking out of the SVG style/tag.
    expect(sanitizeRgbHex('"><script>alert(1)</script>')).toBeUndefined();
    expect(sanitizeRgbHex('FF0000"/><img src=x onerror=alert(1)>')).toBeUndefined();
    expect(sanitizeRgbHex('url(javascript:alert(1))')).toBeUndefined();
  });
});

describe('parseColorElement drops malicious srgbClr/sysClr values', () => {
  test('valid srgbClr is preserved', () => {
    expect(parseColor('<a:srgbClr val="1F497D"/>')).toEqual({ rgb: '1F497D' });
  });

  test('non-hex srgbClr val yields no color rather than passing it through', () => {
    // An attacker-controlled val that is not clean hex must not reach the DOM.
    expect(parseColor('<a:srgbClr val="zzzzzz"/>')).toBeUndefined();
  });

  test('non-hex sysClr lastClr falls back to black', () => {
    expect(parseColor('<a:sysClr val="windowText" lastClr="not-a-color"/>')).toEqual({
      rgb: '000000',
    });
  });
});

describe('resolveColorValueToHex defensive guard', () => {
  test('emits valid hex with a leading #', () => {
    expect(resolveColorValueToHex({ rgb: 'FF0000' })).toBe('#FF0000');
  });

  test('never emits an unvalidated rgb into CSS', () => {
    // Even if a bad rgb is constructed directly (bypassing the parse boundary).
    expect(resolveColorValueToHex({ rgb: '"><script>' })).toBeUndefined();
  });
});
