/**
 * useFootnoteOverlay — paints the caret + selection rects over the
 * actively-edited painted footnote (Vue parity with React's footnote overlay).
 *
 * The actively-edited footnote's hidden PM holds the selection but the painter
 * is the sole visible renderer, so the user has no caret unless we draw one
 * over the painted `.layout-footnote-content[data-footnote-id]`. The DOM walk
 * MUST be scoped to the active container by id (footnote PM positions collide
 * across footnotes), via the footnote-scoped core helpers. Coords are
 * viewport-relative (position: fixed), matching the HF overlay.
 *
 * Extracted from DocxEditor.vue so that component stays under the max-lines cap
 * (mirror of how the HF overlay is wired inline there). The template binds the
 * returned `footnoteCaretRect` / `footnoteSelectionRects`.
 */

import { onBeforeUnmount, onMounted, ref, watch, type Ref } from 'vue';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import {
  computeFootnoteCaretRectFromView,
  computeFootnoteSelectionRectsFromView,
} from '@eigenpal/docx-editor-core/layout-bridge';

export interface UseFootnoteOverlayOptions {
  /** Active footnote id (null = none). */
  footnoteEditId: Ref<number | null>;
  /** Resolve the live footnote EditorView. */
  getFootnoteView: () => EditorView | null;
  /** Subscribe to footnote-PM transactions (for overlay + toolbar sync). */
  setFootnoteTransactionListener: (
    cb: ((id: number, view: EditorView, docChanged: boolean) => void) | null
  ) => void;
  /** Body editor view ref — collapsed + blurred while a footnote is active. */
  editorView: Ref<EditorView | null>;
  /** Painted pages container — `painter:painted` fires on it. */
  pagesRef: Ref<HTMLElement | null>;
  /** Bump to refresh toolbar selection state against the footnote view. */
  stateTick: Ref<number>;
  /** Force the body selection overlay to re-render (hide/show body caret). */
  updateSelectionOverlay: () => void;
}

export interface UseFootnoteOverlayReturn {
  footnoteCaretRect: Ref<{ top: number; left: number; height: number } | null>;
  footnoteSelectionRects: Ref<Array<{ top: number; left: number; width: number; height: number }>>;
}

export function useFootnoteOverlay(opts: UseFootnoteOverlayOptions): UseFootnoteOverlayReturn {
  const {
    footnoteEditId,
    getFootnoteView,
    setFootnoteTransactionListener,
    editorView,
    pagesRef,
    stateTick,
    updateSelectionOverlay,
  } = opts;

  const footnoteCaretRect = ref<{ top: number; left: number; height: number } | null>(null);
  const footnoteSelectionRects = ref<
    Array<{ top: number; left: number; width: number; height: number }>
  >([]);

  function applyFootnoteOverlay(view: EditorView, id: number) {
    const container = window.document.querySelector<HTMLElement>(
      `.layout-footnote-content[data-footnote-id="${id}"]`
    );
    if (!container) {
      footnoteCaretRect.value = null;
      footnoteSelectionRects.value = [];
      return;
    }
    footnoteCaretRect.value = computeFootnoteCaretRectFromView(view, container);
    footnoteSelectionRects.value = computeFootnoteSelectionRectsFromView(view, container);
  }
  function clearFootnoteOverlay() {
    footnoteCaretRect.value = null;
    footnoteSelectionRects.value = [];
  }

  onMounted(() => {
    // Recompute the painted-footnote overlay after each footnote transaction.
    // Wait for the painter's `painter:painted` (with a one-shot rAF fallback
    // for selection-only transactions that skip a layout pass) so the
    // measurement sees the fresh `data-pm-start` spans. Footnote-edit mode
    // keeps the body PM read-only, so the toolbar selection state tracks the
    // footnote view while it's active.
    setFootnoteTransactionListener((id, view) => {
      stateTick.value++;
      const pagesEl = pagesRef.value;
      let painted = false;
      const apply = () => {
        if (painted) return;
        painted = true;
        if (footnoteEditId.value === id) applyFootnoteOverlay(view, id);
      };
      pagesEl?.addEventListener('painter:painted', apply, { once: true });
      requestAnimationFrame(() => {
        if (!painted) {
          pagesEl?.removeEventListener('painter:painted', apply);
          apply();
        }
      });
    });

    // On engage the hidden PM's click-placed selection is already set but no
    // transaction may fire, so measure once the painter has repainted (rAF).
    // On exit, clear the overlay and let the body caret reappear. Also collapse
    // + blur the body view while editing a footnote so the user doesn't see two
    // carets and stray keystrokes can't land in the body.
    watch(footnoteEditId, (id) => {
      if (id == null) {
        clearFootnoteOverlay();
        updateSelectionOverlay();
        return;
      }
      const body = editorView.value;
      if (body) {
        try {
          body.dispatch(body.state.tr.setSelection(TextSelection.create(body.state.doc, 0)));
        } catch {
          // selection may be invalid mid-transition; body caret is gated off so
          // it stays hidden anyway.
        }
        (body.dom as HTMLElement).blur?.();
      }
      updateSelectionOverlay();
      requestAnimationFrame(() => {
        const view = getFootnoteView();
        if (view && footnoteEditId.value === id) applyFootnoteOverlay(view, id);
      });
    });

    // Footnote caret uses position:fixed — recompute on scroll/resize.
    let fnRafScroll = 0;
    function onFootnoteScroll() {
      const id = footnoteEditId.value;
      if (id == null || fnRafScroll) return;
      fnRafScroll = requestAnimationFrame(() => {
        fnRafScroll = 0;
        const view = getFootnoteView();
        if (view) applyFootnoteOverlay(view, id);
      });
    }
    window.addEventListener('scroll', onFootnoteScroll, true);
    window.addEventListener('resize', onFootnoteScroll);
    onBeforeUnmount(() => {
      if (fnRafScroll) cancelAnimationFrame(fnRafScroll);
      window.removeEventListener('scroll', onFootnoteScroll, true);
      window.removeEventListener('resize', onFootnoteScroll);
    });
  });

  return { footnoteCaretRect, footnoteSelectionRects };
}
