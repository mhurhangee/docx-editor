/**
 * Word parity — a footnote/endnote anchor is superscript ONLY when its
 * character style says so.
 *
 * In Word the raised appearance of a reference mark comes entirely from the
 * FootnoteReference / EndnoteReference character style (`w:vertAlign`); the
 * `<w:footnoteReference>` element carries no intrinsic superscript. A bare
 * anchor run (no rStyle — e.g. Pandoc's `<w:r><w:footnoteReference/></w:r>`)
 * renders at the baseline in Word.
 *
 * docx-editor previously force-superscripted every anchor in the layout bridge,
 * which "corrected" bare anchors and so diverged from Word. The implicit hack is
 * gone: superscript now flows solely from the `superscript` mark, which the
 * conversion derives from the run's own vertAlign or resolves from the
 * character-style chain. These tests lock both directions.
 */

import { test, expect } from 'bun:test';
import { convertRun } from '../toProseDoc/runs';
import type { Run } from '../../../types/content/run';
import type { StyleResolver } from '../../styles';

// Minimal StyleResolver stub: hard-codes FootnoteReference's superscript.
// In production `getRunStyleOwnProperties` reads `style.rPr` from a style whose
// `w:basedOn` chain styleParser has ALREADY merged in (resolveStyleInheritance,
// styleParser.ts), so a `vertAlign:superscript` inherited from a basedOn parent
// is honored too — not just one declared directly on FootnoteReference. This
// stub just returns the own-property case the tests below exercise.
const resolver = {
  getRunStyleOwnProperties(styleId?: string | null) {
    if (styleId === 'FootnoteReference' || styleId === 'EndnoteReference') {
      return { vertAlign: 'superscript' };
    }
    return undefined;
  },
} as unknown as StyleResolver;

function markNames(run: Run): string[] {
  const nodes = convertRun(run, undefined, resolver);
  return nodes[0].marks.map((m) => m.type.name);
}

test('styled footnote anchor (rStyle=FootnoteReference) is superscript via the style chain', () => {
  const run = {
    formatting: { styleId: 'FootnoteReference' },
    content: [{ type: 'footnoteRef', id: 1 }],
  } as Run;
  const marks = markNames(run);
  expect(marks).toContain('footnoteRef');
  expect(marks).toContain('superscript');
});

test('bare footnote anchor (no rStyle) is NOT superscript — matches Word', () => {
  const run = {
    content: [{ type: 'footnoteRef', id: 1 }],
  } as Run;
  const marks = markNames(run);
  expect(marks).toContain('footnoteRef');
  expect(marks).not.toContain('superscript');
});

test('bare endnote anchor (no rStyle) is NOT superscript — matches Word', () => {
  const run = {
    content: [{ type: 'endnoteRef', id: 2 }],
  } as Run;
  const marks = markNames(run);
  // `endnoteRef` content converts to the same `footnoteRef` PM mark, tagged
  // with noteType:'endnote' — so asserting `footnoteRef` here is intentional.
  expect(marks).toContain('footnoteRef');
  expect(marks).not.toContain('superscript');
});
