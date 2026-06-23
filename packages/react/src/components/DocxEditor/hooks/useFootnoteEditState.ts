/**
 * Painter PM-doc resolvers + footnote-editing state for PagedEditor.
 *
 * Owns the lazy single hidden-footnote-PM slot — footnote-editing
 * unification (Step 2). At most one footnote is editable at a time;
 * `footnoteEditId` names it (null = none). Click routing / blur logic
 * lands in Step 3 — for now the plumbing exists so the painter can read
 * the live footnote doc via the `getFootnotePmDoc` seam.
 *
 * Also owns the HF PM-doc lookup (`getHfPmDoc`) — the header/footer sibling
 * of `getFootnotePmDoc`. Both feed the layout pipeline so the painter can
 * consult the live hidden PM docs; they are grouped here as the painter's
 * PM-doc resolver family.
 *
 * The `<HiddenFootnotePM>` JSX mount stays in PagedEditor; this hook owns
 * the state, the EditorView ref, and the resolver callbacks it consumes.
 */

import { useCallback, useEffect, useRef, useState } from 'react';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';

import type { Document, HeaderFooter } from '@eigenpal/docx-editor-core/types/document';

import type { HiddenHeaderFooterPMsRef } from '../HiddenHeaderFooterPMs';
import type { HiddenFootnotePMRef } from '../HiddenFootnotePM';
import type { HiddenProseMirrorRef } from '../HiddenProseMirror';

export interface UseFootnoteEditStateOptions {
  /** Latest `document` prop — used to resolve a HeaderFooter's rId. */
  document: Document | null | undefined;
  /** Ref to the persistent hidden HF PM EditorViews (per rId). */
  hiddenHfPMsRef: React.RefObject<HiddenHeaderFooterPMsRef | null>;
  /** Which HF slot (if any) is being edited — selects the active HF rId. */
  hfEditMode?: 'header' | 'footer' | null;
  /** Ref to the body PM — restored on exit from footnote-edit mode. */
  bodyPmRef: React.RefObject<HiddenProseMirrorRef | null>;
  /** Mark the body PM focused (drives the body caret/overlay). */
  setIsFocused: (focused: boolean) => void;
  /**
   * Notified whenever footnote-edit mode changes (active footnote id, or null
   * on exit). DocxEditorPagedArea uses it to drive the painted-footnote caret
   * overlay — resolve the active `.layout-footnote-content` on engage, clear
   * the rects on exit. Mirrors how `hfEditMode` flows to the parent.
   */
  onFootnoteEditChange?: (id: number | null) => void;
}

export interface UseFootnoteEditStateResult {
  /**
   * Persistent hidden HF PM lookup — phase 1 of HF editing unification.
   * Walks `document.package.headers`/`footers` to find the rId for this
   * HeaderFooter instance, then asks the HiddenHeaderFooterPMs ref for its
   * current PM doc. Returns null when no PM is mounted (cold boot before the
   * effect runs) so the pipeline falls back to the Document model path.
   */
  getHfPmDoc: (hf: HeaderFooter) => PMNode | null;
  /**
   * Resolve the active HF EditorView for the current `hfEditMode` slot — the
   * default (or first-page) header/footer reference's rId. Returns null when
   * no HF is being edited or the view isn't mounted. usePagesPointer routes
   * gestures through it when `hfEditMode` is set. Lifted here from PagedEditor
   * to keep that file under its line cap.
   */
  getHfView: () => EditorView | null;
  /** Id of the footnote currently mounted in the lazy single hidden footnote PM, or null. */
  footnoteEditId: number | null;
  /** Setter for the actively-edited footnote id. */
  setFootnoteEditId: React.Dispatch<React.SetStateAction<number | null>>;
  /** Ref for the lazy single hidden footnote PM EditorView. */
  hiddenFootnotePMRef: React.RefObject<HiddenFootnotePMRef | null>;
  /** Resolve the live footnote EditorView (the lazy single slot). */
  getFootnoteView: () => EditorView | null;
  /**
   * Exit footnote-edit mode and hand focus back to the body, placing the body
   * caret at `pos` (or leaving the current selection when null). The actual
   * focus is deferred to an effect that runs AFTER the child HiddenFootnotePM
   * tears its view down — otherwise the footnote view's destroy-blur would
   * steal the focus we just set. Use this (not a bare `setFootnoteEditId(null)`)
   * whenever a click should resume body editing in the same gesture.
   */
  exitFootnoteToBody: (pos: number | null) => void;
  /**
   * Footnote-PM lookup for the painter — mirrors `getHfPmDoc`. Returns the
   * live PM doc only for the actively-edited footnote, so other footnotes
   * still render from their stored `Footnote.content`.
   */
  getFootnotePmDoc: (id: number) => PMNode | null;
  /**
   * Footnote-edit control surface for Step 3 (click routing / blur). Kept in
   * a ref so the click handlers added then can read a stable handle without
   * PagedEditor re-exposing public API.
   */
  footnoteEditApiRef: React.MutableRefObject<{
    setFootnoteEditId: React.Dispatch<React.SetStateAction<number | null>>;
    getFootnoteView: () => EditorView | null;
  }>;
}

export function useFootnoteEditState(
  opts: UseFootnoteEditStateOptions
): UseFootnoteEditStateResult {
  const { document, hiddenHfPMsRef, hfEditMode, bodyPmRef, setIsFocused, onFootnoteEditChange } =
    opts;

  // Persistent hidden HF PM lookup — phase 1 of HF editing unification.
  // Walks `document.package.headers`/`footers` to find the rId for this
  // HeaderFooter instance, then asks the HiddenHeaderFooterPMs ref for
  // its current PM doc. Returns null when no PM is mounted (cold boot
  // before the effect runs) so the pipeline falls back to the Document
  // model path. Stable identity per `document` so the layout pipeline
  // doesn't re-run on every render.
  const getHfPmDoc = useCallback(
    (hf: HeaderFooter): PMNode | null => {
      const ref = hiddenHfPMsRef.current;
      if (!ref) return null;
      const pkg = document?.package;
      if (!pkg) return null;
      const findRid = (bag?: Map<string, HeaderFooter>): string | null => {
        if (!bag) return null;
        for (const [rId, value] of bag) {
          if (value === hf) return rId;
        }
        return null;
      };
      const rId = findRid(pkg.headers) ?? findRid(pkg.footers);
      if (!rId) return null;
      return ref.getView(rId)?.state.doc ?? null;
    },
    [document, hiddenHfPMsRef]
  );

  // Resolve the active HF EditorView for the current `hfEditMode` slot so
  // usePagesPointer routes every gesture (click, drag, multi-click, image-
  // select, hyperlink, context menu) through the HF PM instead of the body PM.
  const getHfView = useCallback((): EditorView | null => {
    const hfRef = hiddenHfPMsRef.current;
    if (!hfRef) return null;
    const sp = document?.package?.document?.finalSectionProperties;
    const refs = hfEditMode === 'header' ? sp?.headerReferences : sp?.footerReferences;
    const refEntry =
      refs?.find((r) => r.type === 'default') ?? refs?.find((r) => r.type === 'first') ?? null;
    if (!refEntry?.rId) return null;
    return hfRef.getView(refEntry.rId);
  }, [hfEditMode, document, hiddenHfPMsRef]);

  /**
   * Lazy single hidden PM EditorView for the actively-edited footnote —
   * footnote-editing unification (Step 2). At most one footnote is editable
   * at a time; `footnoteEditId` names it (null = none). Click routing /
   * blur logic lands in Step 3 — for now the plumbing exists so the painter
   * can read the live footnote doc via the `getFootnotePmDoc` seam.
   */
  const hiddenFootnotePMRef = useRef<HiddenFootnotePMRef>(null);

  // Footnote-editing unification (Step 2): the id of the footnote currently
  // mounted in the lazy single hidden footnote PM, or null. Step 3 wires
  // click routing to call `setFootnoteEditId`; for now it stays null.
  const [footnoteEditId, setFootnoteEditId] = useState<number | null>(null);

  // Resolve the live footnote EditorView (the lazy single slot).
  const getFootnoteView = useCallback(
    (): EditorView | null => hiddenFootnotePMRef.current?.getView() ?? null,
    []
  );

  // Notify the parent of footnote-edit-mode changes so it can drive the
  // painted-footnote caret overlay (engage → resolve container, exit → clear).
  useEffect(() => {
    onFootnoteEditChange?.(footnoteEditId);
  }, [footnoteEditId, onFootnoteEditChange]);

  // Footnote-edit control surface for Step 3 (click routing / blur). Kept in
  // a ref so the click handlers added then can read a stable handle without
  // this component re-exposing public API. `getFootnoteView` is already
  // consumed by `getFootnotePmDoc` below; the setter is parked here until
  // Step 3 wires the gesture path.
  const footnoteEditApiRef = useRef({ setFootnoteEditId, getFootnoteView });
  footnoteEditApiRef.current.setFootnoteEditId = setFootnoteEditId;
  footnoteEditApiRef.current.getFootnoteView = getFootnoteView;

  // Footnote-PM lookup for the painter — mirrors `getHfPmDoc`. Returns the
  // live PM doc only for the actively-edited footnote, so other footnotes
  // still render from their stored `Footnote.content`.
  const getFootnotePmDoc = useCallback(
    (id: number): PMNode | null =>
      id === footnoteEditId ? (getFootnoteView()?.state.doc ?? null) : null,
    [footnoteEditId, getFootnoteView]
  );

  // Reset footnote-edit mode when the actively-edited footnote is no longer
  // present in the document — i.e. a new document was loaded while a footnote
  // was open. Without this, `footnoteEditId` would stay set across a load: the
  // body PM stays readOnly and the next writeback resolves the stale id against
  // the new document (Vue does the equivalent via `destroyFootnotePM()` in its
  // load path). Keyed on the footnote's *existence*, not on `document`
  // identity — `history.state` gets a fresh object on every body edit, so a
  // bare `[document]` reset would exit footnote mode mid-edit. The edited
  // footnote stays present while editing (and the body is readOnly then, so no
  // body edits churn identity), so this only fires on a genuine document swap.
  useEffect(() => {
    if (footnoteEditId == null) return;
    const exists = document?.package?.footnotes?.some(
      (fn) => fn.id === footnoteEditId && (fn.noteType === 'normal' || fn.noteType == null)
    );
    if (!exists) setFootnoteEditId(null);
  }, [document, footnoteEditId]);

  // Exit-to-body focus handoff. `undefined` = no pending request; `number` =
  // place the body caret there; `null` = focus body, keep its selection.
  const pendingBodyFocusRef = useRef<number | null | undefined>(undefined);
  const exitFootnoteToBody = useCallback((pos: number | null) => {
    pendingBodyFocusRef.current = pos;
    setFootnoteEditId(null);
  }, []);

  // When footnote-edit mode clears with a pending body-focus request, restore
  // the body caret + focus. This effect lives in the PARENT (PagedEditor) fiber,
  // so it runs AFTER the child `<HiddenFootnotePM>` has torn its view down on
  // the same commit — the teardown's destroy-blur therefore can't steal the
  // focus set here, and the body PM is already editable (its readOnly prop
  // dropped when footnoteEditId went null).
  useEffect(() => {
    if (footnoteEditId != null) return;
    const pos = pendingBodyFocusRef.current;
    if (pos === undefined) return;
    pendingBodyFocusRef.current = undefined;
    const body = bodyPmRef.current;
    if (!body) return;
    if (pos != null) body.setSelection(pos);
    body.focus();
    setIsFocused(true);
  }, [footnoteEditId, bodyPmRef, setIsFocused]);

  return {
    getHfPmDoc,
    getHfView,
    footnoteEditId,
    setFootnoteEditId,
    hiddenFootnotePMRef,
    getFootnoteView,
    exitFootnoteToBody,
    getFootnotePmDoc,
    footnoteEditApiRef,
  };
}
