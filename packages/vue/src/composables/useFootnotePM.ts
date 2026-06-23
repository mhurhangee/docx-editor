/**
 * useFootnotePM — lazy single hidden footnote ProseMirror view for the Vue
 * adapter. Vue parity for editable footnotes; mirrors React's HiddenFootnotePM
 * + useFootnoteEditState (a single off-screen EditorView for the actively-edited
 * footnote, keyed by `footnoteEditId`).
 *
 * At most ONE footnote is editable at a time; `footnoteEditId` names it
 * (null = none). The painter reads `view.state.doc` via the `getFootnotePmDoc`
 * seam on `computeLayout`, so footnote edits live-render without a second
 * visible PM. On each transaction the doc is serialized back into
 * `Document.package.footnotes[id].content` (and `verbatimXml` cleared) so
 * `save()` reads the latest footnote content.
 *
 * Extracted from `useDocxEditor` so that file stays under the `max-lines` cap;
 * the dependencies it needs (the document ref, the body editor state/view, the
 * layout pipeline, suggestion-mode inputs) are passed in.
 */

import { ref, unref, watch, type MaybeRef, type Ref } from 'vue';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import {
  footnoteToProseDoc,
  proseDocToBlocks,
} from '@eigenpal/docx-editor-core/prosemirror/conversion';
import { schema } from '@eigenpal/docx-editor-core/prosemirror';
import {
  createSuggestionModePlugin,
  setSuggestionMode,
  createDocumentStylesPlugin,
  createDocumentContextPlugin,
} from '@eigenpal/docx-editor-core/prosemirror/plugins';
import {
  ExtensionManager,
  createStarterKit,
} from '@eigenpal/docx-editor-core/prosemirror/extensions';
import type { Document, Footnote } from '@eigenpal/docx-editor-core/types/document';

export interface UseFootnotePMOptions {
  /** Live document ref — `package.footnotes` is the writeback target. */
  document: Ref<Document | null>;
  /** Body editor view ref — focus is restored to it on exit. */
  editorView: Ref<EditorView | null>;
  /** Latest body editor state ref — re-laid out when a footnote edit changes. */
  editorState: Ref<EditorState | null>;
  /** Re-run the layout pipeline so the painter repaints the live footnote doc. */
  runLayoutPipeline: (state: EditorState) => void;
  /** Editor mode (drives suggestion-mode active state on the footnote view). */
  editorMode?: MaybeRef<'editing' | 'suggesting' | 'viewing'>;
  /** Author name attached to tracked changes minted in suggesting mode. */
  author?: MaybeRef<string>;
}

export interface UseFootnotePMReturn {
  /** Id of the footnote in the lazy single hidden footnote PM, or null. */
  footnoteEditId: Ref<number | null>;
  /** Enter/switch/exit footnote-edit mode for a footnote id (null = exit). */
  setFootnoteEditId: (id: number | null) => void;
  /** Resolve the live footnote EditorView (the lazy single slot), or null. */
  getFootnoteView: () => EditorView | null;
  /**
   * Exit footnote-edit mode and restore body focus + caret at `pos` (deferred
   * past the footnote view teardown). Use on a body click that should resume
   * body editing in the same gesture.
   */
  exitFootnoteToBody: (pos: number | null) => void;
  /** Subscribe to every footnote-PM transaction (overlay + toolbar sync). */
  setFootnoteTransactionListener: (
    cb: ((id: number, view: EditorView, docChanged: boolean) => void) | null
  ) => void;
  /** Tear down the footnote view + host (call on document load / destroy). */
  destroyFootnotePM: () => void;
}

export function useFootnotePM(opts: UseFootnotePMOptions): UseFootnotePMReturn {
  const { document, editorView, editorState, runLayoutPipeline, editorMode, author } = opts;

  /** Id of the footnote mounted in the lazy single hidden footnote PM, or null. */
  const footnoteEditId = ref<number | null>(null);
  /** Off-screen host owning the footnote EditorView DOM. */
  const footnoteHostRef: { current: HTMLDivElement | null } = { current: null };
  /** The single mounted footnote EditorView, or null. */
  let footnoteView: EditorView | null = null;
  /** ExtensionManager owning the footnote view's plugins/commands. */
  let footnoteManager: ExtensionManager | null = null;

  function ensureFootnoteHost(): HTMLDivElement {
    if (footnoteHostRef.current && footnoteHostRef.current.isConnected) {
      return footnoteHostRef.current;
    }
    const host = window.document.createElement('div');
    host.dataset.footnoteHost = 'true';
    host.style.cssText =
      'position: fixed; left: -9999px; top: 0; opacity: 0; z-index: -1; pointer-events: none;';
    window.document.body.appendChild(host);
    footnoteHostRef.current = host;
    return host;
  }

  function resolveFootnote(id: number): Footnote | null {
    const footnotes = document.value?.package?.footnotes;
    if (!footnotes) return null;
    return (
      footnotes.find((f) => f.id === id && (f.noteType === 'normal' || f.noteType == null)) ?? null
    );
  }

  function teardownFootnotePM() {
    if (footnoteView) {
      footnoteView.destroy();
      footnoteView.dom.parentElement?.remove();
      footnoteView = null;
    }
    footnoteManager?.destroy();
    footnoteManager = null;
  }

  /** Resolve the live footnote EditorView (the lazy single slot). */
  function getFootnoteView(): EditorView | null {
    return footnoteView;
  }

  // Listener slot for footnote transactions (DocxEditor.vue subscribes here to
  // recompute the painted-footnote caret/selection overlay + toolbar state).
  const onFootnoteTransactionRef: {
    value: ((id: number, view: EditorView, docChanged: boolean) => void) | null;
  } = { value: null };

  /**
   * Build (or rebuild) the single hidden footnote view to match
   * `footnoteEditId`. Tears the old view down first. Mirrors React's
   * HiddenFootnotePM mount effect.
   */
  function syncFootnotePM() {
    teardownFootnotePM();
    const id = footnoteEditId.value;
    if (id == null) return;
    const fn = resolveFootnote(id);
    if (!fn) return;

    const pkg = document.value?.package;
    const styles = pkg?.styles ?? null;
    const theme = pkg?.theme ?? null;
    const defaultTabStopTwips = pkg?.settings?.defaultTabStop ?? null;
    const defaultTableStyleId = pkg?.settings?.defaultTableStyle ?? null;

    const host = ensureFootnoteHost();
    const node = window.document.createElement('div');
    node.dataset.footnotePmId = String(id);
    host.appendChild(node);

    const mgr = new ExtensionManager(createStarterKit());
    mgr.buildSchema();
    mgr.initializeRuntime();
    footnoteManager = mgr;

    const pmDoc = footnoteToProseDoc(fn.content, {
      styles: styles ?? undefined,
      theme,
      defaultTabStopTwips,
    });
    // Footnote paragraphs share the document's style table → style-aware
    // behavior (e.g. Enter after a heading → body text).
    const fnStyleResolverPlugin = createDocumentStylesPlugin(styles);
    const fnDocumentContextPlugin = createDocumentContextPlugin({ theme, defaultTableStyleId });
    const fnSuggestionPlugin = createSuggestionModePlugin(
      unref(editorMode) === 'suggesting',
      unref(author)
    );
    const state = EditorState.create({
      doc: pmDoc,
      schema,
      plugins: [
        fnSuggestionPlugin,
        ...mgr.getPlugins(),
        fnStyleResolverPlugin,
        fnDocumentContextPlugin,
      ],
    });
    const view: EditorView = new EditorView(node, {
      state,
      dispatchTransaction(tr) {
        const newState = view.state.apply(tr);
        view.updateState(newState);
        const activeId = footnoteEditId.value;
        // Writeback: serialize the footnote PM doc back into
        // `Document.package.footnotes[id].content` so `save()` reads the latest
        // content. Clear `verbatimXml` — once edited the captured original bytes
        // are stale, so the serializer must rebuild from `content`.
        if (tr.docChanged && activeId != null) {
          const target = resolveFootnote(activeId);
          if (target) {
            target.content = proseDocToBlocks(newState.doc);
            delete target.verbatimXml;
          }
          // Re-layout so the painter repaints the live footnote doc.
          if (editorState.value) runLayoutPipeline(editorState.value);
        }
        if (activeId != null) onFootnoteTransactionRef.value?.(activeId, view, tr.docChanged);
      },
    });
    footnoteView = view;
  }

  // Rebuild the footnote view whenever the active footnote id changes.
  watch(footnoteEditId, () => {
    syncFootnotePM();
  });

  // Keep footnote suggestion-mode in sync with editorMode/author (mirror HF).
  watch([() => unref(editorMode), () => unref(author)], ([mode, who]) => {
    if (footnoteView) {
      setSuggestionMode(mode === 'suggesting', footnoteView.state, footnoteView.dispatch, who);
    }
  });

  /** Enter/switch/exit footnote-edit mode for a given footnote id (null = exit). */
  function setFootnoteEditId(id: number | null) {
    footnoteEditId.value = id;
  }

  /**
   * Exit footnote-edit mode and hand focus back to the body, placing the body
   * caret at `pos` (or leaving the current selection when null). The footnote
   * view is torn down synchronously by the `footnoteEditId` watch; the body
   * focus is deferred a frame so the teardown's destroy-blur can't steal it.
   */
  function exitFootnoteToBody(pos: number | null) {
    footnoteEditId.value = null;
    requestAnimationFrame(() => {
      const view = editorView.value;
      if (!view) return;
      if (pos != null) {
        try {
          const sel = TextSelection.create(view.state.doc, pos);
          view.dispatch(view.state.tr.setSelection(sel));
        } catch {
          // pos may be invalid for the body doc; ignore and just focus.
        }
      }
      view.focus();
    });
  }

  function setFootnoteTransactionListener(
    cb: ((id: number, view: EditorView, docChanged: boolean) => void) | null
  ) {
    onFootnoteTransactionRef.value = cb;
  }

  function destroyFootnotePM() {
    teardownFootnotePM();
    footnoteEditId.value = null;
    if (footnoteHostRef.current) {
      footnoteHostRef.current.remove();
      footnoteHostRef.current = null;
    }
  }

  return {
    footnoteEditId,
    setFootnoteEditId,
    getFootnoteView,
    exitFootnoteToBody,
    setFootnoteTransactionListener,
    destroyFootnotePM,
  };
}
