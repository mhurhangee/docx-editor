import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import type { EditorView } from 'prosemirror-view';

import { resetImeCaretAnchor, syncImeCaretAnchor } from '../imeCaretAnchor';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

function mockEditorView(options: {
  focused?: boolean;
  composing?: boolean;
  empty?: boolean;
  head?: number;
  coords?: { left: number; top: number };
}): EditorView {
  return {
    composing: options.composing ?? false,
    hasFocus: () => options.focused ?? true,
    state: {
      selection: {
        empty: options.empty ?? true,
        head: options.head ?? 5,
      },
    },
    coordsAtPos: () => options.coords ?? { left: -9999, top: 0 },
  } as unknown as EditorView;
}

describe('syncImeCaretAnchor', () => {
  test('translates the hidden ProseMirror host so its caret matches the painted caret', () => {
    const host = document.createElement('div');
    const view = mockEditorView({ coords: { left: -9999, top: 0 } });

    const anchored = syncImeCaretAnchor({
      hiddenHost: host,
      editorView: view,
      visibleCaret: { left: 320, top: 240 },
    });

    expect(anchored).toBe(true);
    expect(host.style.transform).toBe('translate3d(10319px, 240px, 0)');
  });

  test('measures from the untransformed baseline on repeated updates', () => {
    const host = document.createElement('div');
    host.style.transform = 'translate3d(10319px, 240px, 0)';
    const view = mockEditorView({ coords: { left: -9999, top: 0 } });

    syncImeCaretAnchor({
      hiddenHost: host,
      editorView: view,
      visibleCaret: { left: 321, top: 241 },
    });

    expect(host.style.transform).toBe('translate3d(10320px, 241px, 0)');
  });

  test('resets the anchor for range selections or unfocused editors', () => {
    const host = document.createElement('div');
    host.style.transform = 'translate3d(10px, 20px, 0)';

    const rangeAnchored = syncImeCaretAnchor({
      hiddenHost: host,
      editorView: mockEditorView({ empty: false }),
      visibleCaret: { left: 10, top: 20 },
    });

    expect(rangeAnchored).toBe(false);
    expect(host.style.transform).toBe('');

    host.style.transform = 'translate3d(10px, 20px, 0)';
    const focusAnchored = syncImeCaretAnchor({
      hiddenHost: host,
      editorView: mockEditorView({ focused: false }),
      visibleCaret: { left: 10, top: 20 },
    });

    expect(focusAnchored).toBe(false);
    expect(host.style.transform).toBe('');
  });

  test('leaves the current anchor untouched while composition is active', () => {
    const host = document.createElement('div');
    host.style.transform = 'translate3d(10px, 20px, 0)';

    const anchored = syncImeCaretAnchor({
      hiddenHost: host,
      editorView: mockEditorView({ composing: true }),
      visibleCaret: { left: 320, top: 240 },
    });

    expect(anchored).toBe(false);
    expect(host.style.transform).toBe('translate3d(10px, 20px, 0)');

    resetImeCaretAnchor(host);
    expect(host.style.transform).toBe('');
  });
});
