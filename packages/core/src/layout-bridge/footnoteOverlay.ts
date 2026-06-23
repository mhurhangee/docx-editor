/**
 * Footnote caret + selection overlay helpers.
 *
 * The footnote-editing unification (Step 4) mirrors the header/footer caret
 * overlay (`computeHfCaretRectFromView` / `computeHfSelectionRectsFromView` in
 * `headerFooterLayout.ts`): the painter is the sole visible renderer, and the
 * actively-edited footnote's hidden PM EditorView holds the selection — so the
 * user has no visible caret in the painted footnote unless the host draws one.
 *
 * These helpers turn the footnote view's selection into viewport-relative
 * rects by walking the painted `span[data-pm-start][data-pm-end]` markers.
 *
 * **Footnote-specific scoping constraint.** Each footnote is its own PM doc, so
 * footnote PM positions COLLIDE across footnotes (every footnote starts at 0).
 * The DOM walk MUST be scoped to the ACTIVE footnote's container only — the
 * caller passes in the resolved `.layout-footnote-content[data-footnote-id]`
 * element. Querying `span[data-pm-start]` page-wide would snap the caret to a
 * different footnote's span at the same numeric position. This is why these are
 * separate from the HF helpers (which cache + pick a host by section, and have
 * no notion of a per-document container); the section-keyed HF cache cannot
 * express "this one footnote container".
 *
 * Framework-free: DOM + EditorView only, matching the HF helpers' signatures
 * and return types so the React/Vue adapters can share one implementation.
 */

import type { EditorView } from 'prosemirror-view';

/**
 * Viewport-relative caret rect for a footnote EditorView's collapsed selection
 * head, resolved against the painted spans inside `container` (the active
 * `.layout-footnote-content`). Returns null when the selection is non-empty or
 * no painted span/anchor brackets the PM head.
 *
 * @public
 */
export function computeFootnoteCaretRectFromView(
  view: EditorView,
  container: HTMLElement
): { top: number; left: number; height: number } | null {
  const sel = view.state.selection;
  if (!sel.empty) return null;
  const pmPos = sel.head;
  const spans = Array.from(
    container.querySelectorAll<HTMLElement>('span[data-pm-start][data-pm-end]')
  );
  for (const span of spans) {
    const start = Number(span.dataset.pmStart);
    const end = Number(span.dataset.pmEnd);
    if (!Number.isFinite(start) || !Number.isFinite(end)) continue;
    if (pmPos >= start && pmPos <= end) {
      const range = container.ownerDocument.createRange();
      const walker = container.ownerDocument.createTreeWalker(span, NodeFilter.SHOW_TEXT);
      let remaining = pmPos - start;
      let textNode = walker.nextNode() as Text | null;
      while (textNode) {
        const len = textNode.data.length;
        if (remaining <= len) {
          try {
            range.setStart(textNode, remaining);
            range.setEnd(textNode, remaining);
            const rect = range.getClientRects()[0] ?? range.getBoundingClientRect();
            if (rect && rect.height > 0) {
              return { top: rect.top, left: rect.left, height: rect.height };
            }
          } catch {
            // fall through
          }
          break;
        }
        remaining -= len;
        textNode = walker.nextNode() as Text | null;
      }
      const spanRect = span.getBoundingClientRect();
      const ratio = (pmPos - start) / Math.max(1, end - start);
      return {
        top: spanRect.top,
        left: spanRect.left + spanRect.width * ratio,
        height: spanRect.height,
      };
    }
  }

  // Exact paragraph/line anchor at `pmPos` (when the painter emits one).
  const anchor = container.querySelector<HTMLElement>(`[data-pm-start="${pmPos}"]`);
  if (anchor) {
    const rect = anchor.getBoundingClientRect();
    return { top: rect.top, left: rect.left + 1, height: rect.height || 16 };
  }

  // Fallback for empty paragraphs / line-ends: walk every painted element that
  // carries `[data-pm-start][data-pm-end]` and find the tightest one whose
  // range brackets `pmPos`. Left edge for the paragraph start, right edge when
  // the cursor sits at the paragraph end.
  const ranged = Array.from(
    container.querySelectorAll<HTMLElement>('[data-pm-start][data-pm-end]')
  );
  let bestEl: HTMLElement | null = null;
  let bestSpan = Infinity;
  for (const el of ranged) {
    const start = Number(el.dataset.pmStart);
    const end = Number(el.dataset.pmEnd);
    if (!Number.isFinite(start) || !Number.isFinite(end)) continue;
    if (pmPos < start || pmPos > end) continue;
    const span = end - start;
    if (span < bestSpan) {
      bestSpan = span;
      bestEl = el;
    }
  }
  if (bestEl) {
    const rect = bestEl.getBoundingClientRect();
    const end = Number(bestEl.dataset.pmEnd);
    const atEnd = pmPos >= end;
    return {
      top: rect.top,
      left: atEnd ? rect.right : rect.left + 1,
      height: rect.height || 16,
    };
  }

  // Cursor past every painted range (typically end of last paragraph): snap to
  // the trailing edge of the painted element with the largest `pmStart <= pmPos`.
  let trailingEl: HTMLElement | null = null;
  let trailingStart = -Infinity;
  for (const el of ranged) {
    const start = Number(el.dataset.pmStart);
    if (!Number.isFinite(start)) continue;
    if (start > pmPos) continue;
    if (start > trailingStart) {
      trailingStart = start;
      trailingEl = el;
    }
  }
  if (trailingEl) {
    const rect = trailingEl.getBoundingClientRect();
    return { top: rect.top, left: rect.right, height: rect.height || 16 };
  }

  // Last resort: anchor at the container's top-left so the caret is at least
  // visible while editing. Better than disappearing.
  const containerRect = container.getBoundingClientRect();
  return {
    top: containerRect.top + 2,
    left: containerRect.left + 2,
    height: 16,
  };
}

/**
 * Selection-rect set for a non-empty footnote selection, projected against the
 * painted spans inside `container`. Mirror of `computeHfSelectionRectsFromView`
 * but scoped to a single footnote container (see scoping constraint above).
 *
 * Returns viewport-relative `{top, left, width, height}` rects. Empty array
 * when the selection is collapsed or no painted span overlaps the range.
 *
 * @public
 */
export function computeFootnoteSelectionRectsFromView(
  view: EditorView,
  container: HTMLElement
): Array<{ top: number; left: number; width: number; height: number }> {
  const sel = view.state.selection;
  if (sel.empty) return [];
  const from = sel.from;
  const to = sel.to;
  const out: Array<{ top: number; left: number; width: number; height: number }> = [];

  const spans = Array.from(
    container.querySelectorAll<HTMLElement>('span[data-pm-start][data-pm-end]')
  );
  for (const spanEl of spans) {
    const pmStart = Number(spanEl.dataset.pmStart);
    const pmEnd = Number(spanEl.dataset.pmEnd);
    if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) continue;
    if (pmEnd <= from || pmStart >= to) continue;

    // Tab spans: full-span highlight.
    if (spanEl.classList.contains('layout-run-tab')) {
      const rect = spanEl.getBoundingClientRect();
      out.push({ top: rect.top, left: rect.left, width: rect.width, height: rect.height });
      continue;
    }

    let textNode: Text | null = null;
    if (spanEl.firstChild?.nodeType === Node.TEXT_NODE) {
      textNode = spanEl.firstChild as Text;
    } else if (
      spanEl.firstChild?.nodeType === Node.ELEMENT_NODE &&
      (spanEl.firstChild as HTMLElement).tagName === 'A' &&
      spanEl.firstChild.firstChild?.nodeType === Node.TEXT_NODE
    ) {
      textNode = spanEl.firstChild.firstChild as Text;
    }
    if (!textNode) continue;

    const startChar = Math.max(0, from - pmStart);
    const endChar = Math.min(textNode.length, to - pmStart);
    if (startChar >= endChar) continue;

    const range = container.ownerDocument.createRange();
    range.setStart(textNode, startChar);
    range.setEnd(textNode, endChar);
    for (const rect of Array.from(range.getClientRects())) {
      out.push({ top: rect.top, left: rect.left, width: rect.width, height: rect.height });
    }
  }

  return out;
}
