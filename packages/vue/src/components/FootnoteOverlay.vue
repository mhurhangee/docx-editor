<script setup lang="ts">
/**
 * FootnoteOverlay — presentational caret + selection rects for the
 * actively-edited painted footnote (Vue parity with React's footnote overlay).
 *
 * The caret is a thin black blinking bar (matches the body caret); selection rects are blue
 * highlights. Coords are viewport-relative (position: fixed), matching the HF
 * overlay. Caret and selection are mutually exclusive (the caret only shows
 * when no range is selected). Rendered only while a footnote is the active
 * edit surface (`active`).
 */
defineProps<{
  /** True while a footnote is the active edit surface. */
  active: boolean;
  /** Caret rect (viewport-relative) when the footnote selection is collapsed. */
  caretRect: { top: number; left: number; height: number } | null;
  /** Selection rects (viewport-relative) for a non-empty footnote selection. */
  selectionRects: Array<{ top: number; left: number; width: number; height: number }>;
}>();
</script>

<template>
  <template v-if="active">
    <div
      v-for="(rect, i) in selectionRects"
      :key="`fn-sel-${i}-${rect.top}-${rect.left}`"
      class="vue-footnote-sel-rect"
      aria-hidden="true"
      :style="{
        position: 'fixed',
        top: `${rect.top}px`,
        left: `${rect.left}px`,
        width: `${rect.width}px`,
        height: `${rect.height}px`,
        background: 'var(--doc-selection)',
        pointerEvents: 'none',
        zIndex: 9998,
      }"
    />
    <div
      v-if="caretRect && selectionRects.length === 0"
      aria-hidden="true"
      :style="{
        position: 'fixed',
        top: `${caretRect.top}px`,
        left: `${caretRect.left}px`,
        width: '2px',
        height: `${caretRect.height}px`,
        background: 'var(--doc-caret, #000)',
        pointerEvents: 'none',
        zIndex: 9999,
        animation: 'hf-caret-blink 1.06s steps(1) infinite',
      }"
    />
  </template>
</template>
