import type { EditorView } from 'prosemirror-view';

export interface VisibleCaretViewportRect {
  left: number;
  top: number;
}

export interface SyncImeCaretAnchorOptions {
  hiddenHost: HTMLElement | null | undefined;
  editorView: EditorView | null | undefined;
  visibleCaret: VisibleCaretViewportRect | null | undefined;
}

function resetHostTransform(hiddenHost: HTMLElement | null | undefined): void {
  if (hiddenHost) hiddenHost.style.transform = '';
}

function isFiniteCoord(value: number): boolean {
  return Number.isFinite(value);
}

/**
 * Keep the real off-screen ProseMirror caret aligned with the painted caret.
 *
 * Native IME candidate windows use the focused contenteditable caret geometry,
 * not the painted caret overlay. The paged editors keep ProseMirror off-screen
 * and paint the document separately, so CJK IME panels otherwise appear at the
 * off-screen PM location. Translating the hidden host preserves the split
 * rendering model while making browser/OS IME geometry land on the visible
 * insertion point.
 */
export function syncImeCaretAnchor({
  hiddenHost,
  editorView,
  visibleCaret,
}: SyncImeCaretAnchorOptions): boolean {
  if (!hiddenHost) return false;

  if (!editorView || !visibleCaret || !editorView.hasFocus()) {
    resetHostTransform(hiddenHost);
    return false;
  }

  // Do not move the focused composition target while the IME is open.
  if (editorView.composing) return false;

  const selection = editorView.state.selection;
  if (!selection.empty) {
    resetHostTransform(hiddenHost);
    return false;
  }

  let hiddenCaret: { left: number; top: number };
  const previousTransform = hiddenHost.style.transform;
  try {
    // coordsAtPos includes CSS transforms. Measure the off-screen baseline so
    // repeated selection updates keep the same anchor instead of cancelling the
    // previous translation back to zero.
    hiddenHost.style.transform = '';
    hiddenCaret = editorView.coordsAtPos(selection.head);
  } catch {
    hiddenHost.style.transform = previousTransform;
    return false;
  }

  if (
    !isFiniteCoord(hiddenCaret.left) ||
    !isFiniteCoord(hiddenCaret.top) ||
    !isFiniteCoord(visibleCaret.left) ||
    !isFiniteCoord(visibleCaret.top)
  ) {
    resetHostTransform(hiddenHost);
    return false;
  }

  const dx = Math.round(visibleCaret.left - hiddenCaret.left);
  const dy = Math.round(visibleCaret.top - hiddenCaret.top);
  hiddenHost.style.transform = `translate3d(${dx}px, ${dy}px, 0)`;
  return true;
}

export function resetImeCaretAnchor(hiddenHost: HTMLElement | null | undefined): void {
  resetHostTransform(hiddenHost);
}
