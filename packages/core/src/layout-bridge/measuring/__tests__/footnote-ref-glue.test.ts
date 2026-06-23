import { describe, test, expect } from 'bun:test';

// bun:test has no DOM; give measureTextWidth a deterministic canvas stub
// (8px per character) before the lazy getCanvasContext() first runs — the same
// stub pattern the sibling intrinsic-width / inline-image-wrap-height tests use.
// NOTE: another test file in the suite may install its own global document
// first, in which case this guard is a no-op and measureTextWidth uses that
// stub. So the thresholds below are derived from measureTextWidth itself rather
// than hardcoded pixels — the test stays self-consistent under whichever stub
// is active (a hardcoded width broke when run after a different stub loaded).
if (typeof document === 'undefined') {
  (globalThis as Record<string, unknown>).document = {
    createElement: () => ({
      getContext: () => ({
        font: '',
        measureText: (text: string) => ({ width: text.length * 8 }),
      }),
    }),
  };
}

import { measureParagraph } from '../measureParagraph';
import { measureTextWidth } from '../measureContainer';
import type { ParagraphBlock, Run } from '../../../layout-engine/types';

// runToFontStyle defaults to Calibri / 11pt; match it so the widths we compute
// here line up with what the line-breaker measures for these runs.
const STYLE = { fontFamily: 'Calibri', fontSize: 11 };
const w = (t: string): number => measureTextWidth(t, STYLE);

function para(runs: Run[]): ParagraphBlock {
  return { kind: 'paragraph', id: 'p1', runs } as ParagraphBlock;
}

/** True if a measured line begins at the given run's char 0. */
function lineStartsAtRun(
  lines: { fromRun: number; fromChar: number }[],
  runIndex: number
): boolean {
  return lines.some((l) => l.fromRun === runIndex && l.fromChar === 0);
}

describe('measureParagraph — cross-run glue (footnote reference must not split)', () => {
  // run0 ends with "beta." (no trailing space); run1 is the footnote ref "1"
  // (a separate superscript run, no space before it); run2 begins with a space.
  test('footnote-ref run never starts its own line (no space before it)', () => {
    const runs: Run[] = [
      { kind: 'text', text: 'alpha beta.' },
      { kind: 'text', text: '1', superscript: true },
      { kind: 'text', text: ' gamma' },
    ] as Run[];
    // Width fits "alpha beta." but not the glued "alpha beta.1" (overflow by
    // 0.6px, clear of the 0.5px wrap tolerance). The breaker must move the whole
    // "beta.1" cluster to the next line rather than splitting the ref off.
    const maxWidth = w('alpha beta.') + w('1') - 0.6;
    const { lines } = measureParagraph(para(runs), maxWidth);
    expect(lineStartsAtRun(lines, 1)).toBe(false);
  });

  test('a space before the ref still allows a normal wrap', () => {
    // Same widths, but run0 now ends with a space → "1" is free to wrap; the
    // glue logic must not suppress a legitimate break opportunity.
    const runs: Run[] = [
      { kind: 'text', text: 'alpha beta ' },
      { kind: 'text', text: '1', superscript: true },
      { kind: 'text', text: ' gamma' },
    ] as Run[];
    const maxWidth = w('alpha beta ') + w('1') - 0.6;
    const { lines } = measureParagraph(para(runs), maxWidth);
    expect(lineStartsAtRun(lines, 1)).toBe(true);
  });
});
