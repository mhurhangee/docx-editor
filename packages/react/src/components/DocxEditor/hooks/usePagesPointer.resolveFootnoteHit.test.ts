/**
 * Focused test for `resolveFootnoteHit` (footnote-editing unification, Step 3).
 *
 * The pure helper that footnote click-routing depends on: given a click
 * target, resolve which painted footnote (`data-footnote-id`) it landed in and
 * return the scoped container. Scoping is load-bearing — footnote PM positions
 * collide across footnotes, so callers must snap spans within the returned
 * container only, never the whole page.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

import { resolveFootnoteHit } from './usePagesPointer';

function buildFootnoteDom(): { area: HTMLElement; spanIn2: HTMLElement; outside: HTMLElement } {
  const area = document.createElement('div');
  area.className = 'layout-footnote-area';

  const fn1 = document.createElement('div');
  fn1.className = 'layout-footnote-content';
  fn1.dataset.footnoteId = '1';
  const span1 = document.createElement('span');
  span1.dataset.pmStart = '1';
  span1.dataset.pmEnd = '5';
  span1.textContent = 'one';
  fn1.appendChild(span1);

  const fn2 = document.createElement('div');
  fn2.className = 'layout-footnote-content';
  fn2.dataset.footnoteId = '2';
  const span2 = document.createElement('span');
  span2.dataset.pmStart = '1';
  span2.dataset.pmEnd = '5';
  span2.textContent = 'two';
  fn2.appendChild(span2);

  area.appendChild(fn1);
  area.appendChild(fn2);

  const outside = document.createElement('div');
  outside.className = 'layout-page-content';

  return { area, spanIn2: span2, outside };
}

describe('resolveFootnoteHit', () => {
  test('resolves the footnote id + scoped container for a click on its span', () => {
    const { spanIn2 } = buildFootnoteDom();
    const hit = resolveFootnoteHit(spanIn2);
    expect(hit).not.toBeNull();
    expect(hit?.id).toBe(2);
    // The container is the footnote-2 element, not footnote-1 or the area.
    expect(hit?.container.dataset.footnoteId).toBe('2');
    expect(hit?.container.contains(spanIn2)).toBe(true);
  });

  test('returns null when the target is outside any footnote', () => {
    const { outside } = buildFootnoteDom();
    expect(resolveFootnoteHit(outside)).toBeNull();
  });

  test('returns null for a footnote-content element with no data-footnote-id', () => {
    const el = document.createElement('div');
    el.className = 'layout-footnote-content';
    expect(resolveFootnoteHit(el)).toBeNull();
  });
});
