import { describe, test, expect } from 'bun:test';
import { parseStyleAttr } from '../vmlWatermarkParser';

describe('parseStyleAttr', () => {
  test('parses normal declarations', () => {
    const out = parseStyleAttr('width:415.2pt;height:20pt;rotation:315');
    expect(out.width).toBe('415.2pt');
    expect(out.height).toBe('20pt');
    expect(out.rotation).toBe('315');
  });

  test('does not pollute Object.prototype via __proto__/constructor keys', () => {
    // The style string is attacker-controlled in an untrusted document.
    parseStyleAttr('__proto__:polluted;constructor:x;prototype:y');
    expect(({} as Record<string, unknown>).polluted).toBeUndefined();
    // The dangerous keys are dropped from the result entirely.
    const out = parseStyleAttr('__proto__:a;width:10pt');
    expect(out.width).toBe('10pt');
    expect(Object.prototype.hasOwnProperty.call(out, '__proto__')).toBe(false);
  });
});
