/**
 * HiddenFootnotePM — a single off-screen ProseMirror EditorView for the
 * *actively-edited* footnote.
 *
 * Step 2 of footnote-editing unification (mirrors the HF model in
 * `HiddenHeaderFooterPMs.tsx`). Unlike the HF component, this mounts a LAZY
 * SINGLE slot: at most one EditorView exists at a time, for whichever footnote
 * `activeFootnoteId` names. When `activeFootnoteId` is null no view exists.
 *
 * The view is built from the footnote's content through `footnoteToProseDoc`
 * with the same plugin stack the HF views use (starterkit + suggestion-mode +
 * document styles/context). On each transaction the doc is serialized back into
 * `Document.package.footnotes[id].content` (and `verbatimXml` cleared on a doc
 * change) so `save()` reads the latest footnote content; the painter reads the
 * live doc through the `getFootnotePmDoc` seam on `computeLayout`.
 */

import { useRef, useEffect, useImperativeHandle, forwardRef, memo } from 'react';
import { EditorState } from 'prosemirror-state';
import type { EditorState as EditorStateT } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import {
  schema,
  createDocumentStylesPlugin,
  createDocumentContextPlugin,
} from '@eigenpal/docx-editor-core/prosemirror';
import {
  createSuggestionModePlugin,
  setSuggestionMode,
} from '@eigenpal/docx-editor-core/prosemirror/plugins';
import {
  footnoteToProseDoc,
  proseDocToBlocks,
} from '@eigenpal/docx-editor-core/prosemirror/conversion';
import { createStarterKit } from '@eigenpal/docx-editor-core/prosemirror/extensions';
import { ExtensionManager } from '@eigenpal/docx-editor-core/prosemirror/extensions';
import type {
  Document,
  Footnote,
  StyleDefinitions,
  Theme,
} from '@eigenpal/docx-editor-core/types/document';

import 'prosemirror-view/style/prosemirror.css';

export interface HiddenFootnotePMRef {
  /**
   * The persistent EditorView for the active footnote, or null when no
   * footnote is being edited (or its id is absent from the document).
   */
  getView(): EditorView | null;
}

export interface HiddenFootnotePMProps {
  /** Footnote id currently being edited, or null to hold no view. */
  activeFootnoteId: number | null;
  /** The loaded document — its `package.footnotes` is the writeback target. */
  document: Document | null;
  /** Document styles, threaded into `footnoteToProseDoc` for style resolution. */
  styles?: StyleDefinitions | null;
  /** Document theme, threaded for themed cell shading on initial PM build. */
  theme?: Theme | null;
  /** Suggestion mode active state. */
  isSuggesting?: boolean;
  /** Active author for suggestion mode. */
  author?: string;
  /** `defaultTabStop` from settings, threaded to the footnote PM doc. */
  defaultTabStopTwips?: number | null;
  /**
   * Called after every transaction lands on the footnote EditorView. Mirrors
   * the HF `onTransaction` — the host re-runs the layout pipeline (so the
   * painter repaints) on doc changes.
   */
  onTransaction?: (footnoteId: number, view: EditorView, docChanged: boolean) => void;
}

// ============================================================================
// IMPLEMENTATION
// ============================================================================

function resolveFootnote(doc: Document | null, id: number): Footnote | null {
  const footnotes = doc?.package?.footnotes;
  if (!footnotes) return null;
  return (
    footnotes.find((f) => f.id === id && (f.noteType === 'normal' || f.noteType == null)) ?? null
  );
}

function buildInitialState(
  fn: Footnote,
  styles: StyleDefinitions | null | undefined,
  theme: Theme | null | undefined,
  defaultTabStopTwips: number | null | undefined,
  defaultTableStyleId: string | null | undefined,
  mgr: ExtensionManager,
  isSuggesting: boolean,
  author: string
): EditorState {
  const pmDoc = footnoteToProseDoc(fn.content, {
    styles: styles ?? undefined,
    theme: theme ?? null,
    defaultTabStopTwips: defaultTabStopTwips ?? null,
  });
  // Footnote paragraphs share the document's style table, so they get the
  // same style-aware behavior (e.g. Enter after a heading → body text).
  const styleResolverPlugin = createDocumentStylesPlugin(styles);
  // Document context (theme + settings `w:defaultTableStyle`) so inserting a
  // table in a footnote adopts the default table style too.
  const documentContextPlugin = createDocumentContextPlugin({
    theme: theme ?? null,
    defaultTableStyleId: defaultTableStyleId ?? null,
  });
  const suggestionPlugin = createSuggestionModePlugin(isSuggesting, author);
  return EditorState.create({
    doc: pmDoc,
    schema,
    plugins: [suggestionPlugin, ...mgr.getPlugins(), styleResolverPlugin, documentContextPlugin],
  });
}

export const HiddenFootnotePM = memo(
  forwardRef<HiddenFootnotePMRef, HiddenFootnotePMProps>(function HiddenFootnotePM(
    {
      activeFootnoteId,
      document,
      styles,
      theme,
      isSuggesting = false,
      author = 'User',
      defaultTabStopTwips,
      onTransaction,
    },
    ref
  ) {
    // Keep the callback stable across renders — the EditorView's
    // `dispatchTransaction` closes over it, so going through a ref lets the
    // parent pass a fresh callback each render without recreating the view.
    const onTransactionRef = useRef(onTransaction);
    onTransactionRef.current = onTransaction;

    // styleId for newly inserted tables — primitive so safe in effect deps.
    const defaultTableStyleId = document?.package?.settings?.defaultTableStyle ?? null;

    const isSuggestingRef = useRef(isSuggesting);
    isSuggestingRef.current = isSuggesting;
    const authorRef = useRef(author);
    authorRef.current = author;

    // Latest `document` in a ref so the writeback closure inside
    // `dispatchTransaction` always targets the current footnotes array (its
    // identity changes on every body transaction).
    const documentRef = useRef(document);
    documentRef.current = document;

    const hostRef = useRef<HTMLDivElement>(null);
    const viewRef = useRef<EditorView | null>(null);
    const mountNodeRef = useRef<HTMLElement | null>(null);
    const managerRef = useRef<ExtensionManager | null>(null);
    const activeIdRef = useRef<number | null>(activeFootnoteId);
    activeIdRef.current = activeFootnoteId;

    // Writeback: serialize the footnote PM doc back into
    // `Document.package.footnotes[id].content` so `save()` reads the latest
    // content. Clear `verbatimXml` on a doc change — once the footnote is
    // edited the captured original bytes are stale, so the serializer must
    // rebuild from `content`.
    const syncBlocksToDocumentRef = useRef<((id: number, state: EditorStateT) => void) | null>(
      null
    );
    syncBlocksToDocumentRef.current = (id: number, state: EditorStateT) => {
      const fn = resolveFootnote(documentRef.current, id);
      if (!fn) return;
      fn.content = proseDocToBlocks(state.doc);
      delete fn.verbatimXml;
    };

    // Keep suggestion-mode in sync when isSuggesting/author change (mirror HF).
    useEffect(() => {
      const view = viewRef.current;
      if (!view) return;
      setSuggestionMode(isSuggesting, view.state, view.dispatch, author);
    }, [isSuggesting, author]);

    // Build / tear down the single view when the active footnote changes.
    useEffect(() => {
      const host = hostRef.current;
      if (!host) return;

      // Tear down any existing view first.
      const teardown = () => {
        viewRef.current?.destroy();
        viewRef.current = null;
        mountNodeRef.current?.remove();
        mountNodeRef.current = null;
        managerRef.current?.destroy();
        managerRef.current = null;
      };

      if (activeFootnoteId == null) {
        teardown();
        return;
      }

      const fn = resolveFootnote(document, activeFootnoteId);
      if (!fn) {
        teardown();
        return;
      }

      teardown();

      const mgr = new ExtensionManager(createStarterKit());
      mgr.buildSchema();
      mgr.initializeRuntime();
      managerRef.current = mgr;

      // `document` in this closure is the Document model (DOCX), not the
      // browser DOM `Document`. Create the mount node via the host's owner.
      const node = host.ownerDocument.createElement('div');
      node.dataset.footnotePmId = String(activeFootnoteId);
      host.appendChild(node);
      mountNodeRef.current = node;

      const state = buildInitialState(
        fn,
        styles,
        theme,
        defaultTabStopTwips,
        defaultTableStyleId,
        mgr,
        isSuggestingRef.current,
        authorRef.current
      );
      const view: EditorView = new EditorView(node, {
        state,
        // The persistent footnote PM is the sole footnote editor. Every
        // transaction (typing, click → setSelection, undo/redo) needs to:
        //   1. Re-run the layout pipeline so the painter repaints (rides
        //      through `onTransaction`).
        //   2. Sync `view.state.doc` back into `Document.package.footnotes`
        //      so `save()` doesn't lose unsaved footnote edits (here).
        dispatchTransaction(tr) {
          const newState = view.state.apply(tr);
          view.updateState(newState);
          const id = activeIdRef.current;
          if (id != null && tr.docChanged) syncBlocksToDocumentRef.current?.(id, newState);
          if (id != null) onTransactionRef.current?.(id, view, tr.docChanged);
        },
      });
      viewRef.current = view;

      return teardown;
      // Note: `document` intentionally excluded from deps. The writeback
      // closure reads from `documentRef`; depending on `document` directly
      // would re-mount the footnote view on every body PM transaction (each
      // Document.applyTransaction returns a new identity), destroying IME
      // state and selection.
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [activeFootnoteId, styles, theme, defaultTabStopTwips, defaultTableStyleId]);

    useImperativeHandle(
      ref,
      () => ({
        getView(): EditorView | null {
          return viewRef.current;
        },
      }),
      []
    );

    // Off-screen host — positioned like `HiddenHeaderFooterPMs` so the view
    // retains focusability and keyboard routing while being visually absent.
    return (
      <div
        ref={hostRef}
        style={{
          position: 'fixed',
          left: -9999,
          top: 0,
          opacity: 0,
          zIndex: -1,
          pointerEvents: 'none',
        }}
      />
    );
  })
);

HiddenFootnotePM.displayName = 'HiddenFootnotePM';
