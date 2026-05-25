/**
 * Footnote Layout Utilities
 *
 * Footnote/endnote rendering pipeline plus page-mapping helpers:
 * - scanning FlowBlocks for footnote references and their PM positions
 * - mapping references to the page that ends up containing them
 * - converting a Footnote → FootnoteContent via the body pipeline
 *   (footnoteToProseDoc → toFlowBlocks → caller-supplied measureBlocks)
 * - reserving per-page footnote area heights for layout
 *
 * Everything that's pure OOXML / FlowBlock semantics lives here so the
 * React, Vue, and any future adapters can share the conversion logic
 * and just supply their own measurement function (which depends on
 * platform-specific Canvas/font metrics).
 */

import type {
  FlowBlock,
  ParagraphBlock,
  Measure,
  Page,
  Layout,
  FootnoteContent,
} from '../layout-engine/types';
import { layoutDocument, type LayoutOptions } from '../layout-engine';
import type { Document, Footnote, StyleDefinitions, Theme } from '../types/document';
import type { FootnoteRenderItem } from '../layout-painter';
import { footnoteToProseDoc } from '../prosemirror/conversion/toProseDoc';
import { toFlowBlocks } from './toFlowBlocks';
import { getFootnoteText } from '../docx/footnoteParser';

/** Separator line height + vertical padding in pixels. */
export const FOOTNOTE_SEPARATOR_HEIGHT = 12;

/**
 * Hard cap on the multi-pass footnote layout loop. Reserving footnote
 * space can move a reference to another page, so adapters keep remapping
 * until the page→height contract is stable. Dense layouts converge in
 * 2–3 passes in practice; 6 is a safe ceiling.
 */
export const MAX_FOOTNOTE_LAYOUT_PASSES = 6;

/**
 * Compare two per-page footnote reservation maps. Used by the React +
 * Vue adapters to detect when the multi-pass loop has converged.
 */
export function footnoteReservedHeightsEqual(
  a: Map<number, number>,
  b: Map<number, number>
): boolean {
  if (a.size !== b.size) return false;
  for (const [pageNumber, height] of a) {
    if (b.get(pageNumber) !== height) return false;
  }
  return true;
}

/**
 * Default footnote font size in points. Word's built-in "Footnote Text"
 * style sets 8pt; we apply this only when the footnote's runs don't
 * already specify a fontSize (avoids overriding authored sizes).
 *
 * TODO once the style cascade for paragraph styles is fully wired through
 * the bridge, footnotes should pick this up from the resolved
 * "FootnoteText" / "footnote text" style instead of hardcoding the value.
 */
const FOOTNOTE_FONT_SIZE_PT = 8;

// ============================================================================
// 1. Scan FlowBlocks for footnote references
// ============================================================================

/**
 * Scan FlowBlocks for runs with footnoteRefId set.
 * Returns a list of { footnoteId, pmPos } in document order.
 *
 * Recurses into container blocks (table cells, text boxes) so footnote
 * references authored anywhere in the body reach the page-reservation
 * pass. Without this, a `footnoteRefId` nested inside a table cell never
 * gets mapped to a page and the per-page `.layout-footnote-area` silently
 * drops that entry even though the body still renders the in-line ref
 * marker.
 */
export function collectFootnoteRefs(
  blocks: FlowBlock[]
): Array<{ footnoteId: number; pmPos: number }> {
  const refs: Array<{ footnoteId: number; pmPos: number }> = [];

  const walk = (input: FlowBlock[]): void => {
    for (const block of input) {
      if (block.kind === 'paragraph') {
        for (const run of block.runs) {
          if (run.kind === 'text' && run.footnoteRefId != null) {
            refs.push({
              footnoteId: run.footnoteRefId,
              pmPos: run.pmStart ?? 0,
            });
          }
        }
      } else if (block.kind === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            walk(cell.blocks);
          }
        }
      } else if (block.kind === 'textBox') {
        walk(block.content);
      }
    }
  };

  walk(blocks);

  return refs;
}

// ============================================================================
// 2. Map footnote references to pages
// ============================================================================

/**
 * After layout, determine which footnotes appear on which pages.
 * Checks each page's fragments to see if any footnoteRef PM positions fall within.
 *
 * Returns Map<pageNumber, footnoteId[]> in document order.
 */
export function mapFootnotesToPages(
  pages: Page[],
  footnoteRefs: Array<{ footnoteId: number; pmPos: number }>
): Map<number, number[]> {
  const pageFootnotes = new Map<number, number[]>();

  if (footnoteRefs.length === 0) return pageFootnotes;

  // For each footnote ref, find which page it lands on
  for (const ref of footnoteRefs) {
    for (const page of pages) {
      let found = false;
      for (const fragment of page.fragments) {
        const pmStart = fragment.pmStart ?? -1;
        const pmEnd = fragment.pmEnd ?? -1;
        if (pmStart >= 0 && pmEnd >= 0 && ref.pmPos >= pmStart && ref.pmPos < pmEnd) {
          const existing = pageFootnotes.get(page.number) ?? [];
          // Avoid duplicates (same footnote shouldn't appear twice on same page)
          if (!existing.includes(ref.footnoteId)) {
            existing.push(ref.footnoteId);
          }
          pageFootnotes.set(page.number, existing);
          found = true;
          break;
        }
      }
      if (found) break;
    }
  }

  return pageFootnotes;
}

// ============================================================================
// 3. Convert a footnote to renderable FlowBlocks (body-pipeline)
// ============================================================================

/**
 * Footnote-specific block normalization. Mirrors the spirit of
 * `normalizeHeaderFooterMeasureBlocks`: post-process the body-pipeline
 * output for a single footnote so it carries the correct visual prefix
 * (its display number, rendered as a superscript) and a default 8pt font
 * for any run that didn't specify a size.
 *
 * The displayNumber is prepended onto the FIRST paragraph as a fresh
 * superscript text run — visually matches Word's footnote numbering
 * without disturbing the authored runs.
 *
 * Exported for callers that want to compose their own conversion
 * pipeline; `convertFootnoteToContent` calls it as part of its flow.
 */
export function applyFootnotePresentation(blocks: FlowBlock[], displayNumber: number): FlowBlock[] {
  if (blocks.length === 0) {
    return [
      {
        kind: 'paragraph',
        id: `fn-empty-${displayNumber}`,
        runs: [
          {
            kind: 'text',
            text: `${displayNumber}  `,
            fontSize: FOOTNOTE_FONT_SIZE_PT,
            superscript: true,
          },
        ],
      } as ParagraphBlock,
    ];
  }

  // Apply default 8pt to every run that didn't specify a fontSize. Mutating
  // a copy keeps the input blocks pure for caching upstream.
  const out = blocks.map((b) => {
    if (b.kind !== 'paragraph') return b;
    const para = b as ParagraphBlock;
    return {
      ...para,
      runs: para.runs.map((r) => {
        if (r.kind === 'text' || r.kind === 'tab') {
          if (r.fontSize == null) {
            return { ...r, fontSize: FOOTNOTE_FONT_SIZE_PT };
          }
        }
        return r;
      }),
    } as ParagraphBlock;
  });

  // Prepend display number on the first paragraph.
  const first = out[0];
  if (first.kind === 'paragraph') {
    const numberRun = {
      kind: 'text' as const,
      text: `${displayNumber}  `,
      fontSize: FOOTNOTE_FONT_SIZE_PT,
      superscript: true,
    };
    out[0] = {
      ...(first as ParagraphBlock),
      runs: [numberRun, ...(first as ParagraphBlock).runs],
    } as ParagraphBlock;
  }

  return out;
}

/**
 * Adapter-supplied block measurement function. The caller (React /
 * Vue / etc.) supplies its platform's measure routine — at minimum
 * paragraph + table + image + textBox — so this core helper stays
 * Canvas-free.
 */
export type MeasureBlocksFn = (blocks: FlowBlock[], contentWidth: number) => Measure[];

/**
 * Options for {@link convertFootnoteToContent}.
 */
export type ConvertFootnoteOptions = {
  /** The document's parsed style definitions, threaded into the body pipeline. */
  styles?: StyleDefinitions | null;
  /** Theme for resolving themed fills / fonts inside the footnote. */
  theme?: Theme | null;
  /** Measure callback supplied by the rendering adapter. */
  measureBlocks: MeasureBlocksFn;
  /**
   * Doc-level `w:defaultTabStop` (twips) from the body so list markers
   * inside footnotes honor the same tab grid.
   */
  defaultTabStopTwips?: number | null;
};

/**
 * Convert a Footnote to renderable FootnoteContent via the body pipeline:
 * `footnoteToProseDoc → toFlowBlocks → applyFootnotePresentation →
 * measureBlocks`. Pre-PR (#378) this lived in a hand-rolled shadow stack
 * that silently dropped non-paragraph content; routing through the body
 * pipeline gives footnotes full block-kind support — paragraph + table
 * + image + textBox + fields.
 */
export function convertFootnoteToContent(
  footnote: Footnote,
  displayNumber: number,
  contentWidth: number,
  options: ConvertFootnoteOptions
): FootnoteContent {
  const pmDoc = footnoteToProseDoc(footnote.content, {
    styles: options.styles ?? undefined,
    theme: options.theme ?? null,
    defaultTabStopTwips: options.defaultTabStopTwips ?? null,
  });
  const rawBlocks = toFlowBlocks(pmDoc, { theme: options.theme ?? undefined });
  const blocks = applyFootnotePresentation(rawBlocks, displayNumber);

  const measures = options.measureBlocks(blocks, contentWidth);

  const totalHeight = measures.reduce((h, m) => {
    if (m.kind === 'paragraph') return h + m.totalHeight;
    if (m.kind === 'table') return h + m.totalHeight;
    if (m.kind === 'image') return h + m.height;
    if (m.kind === 'textBox') return h + m.height;
    return h;
  }, 0);

  return {
    id: footnote.id,
    displayNumber,
    blocks,
    measures,
    height: totalHeight,
  };
}

/**
 * Build footnote content for all footnotes referenced in the document.
 * Display numbers are assigned by first-appearance order (the same way
 * Word renders them).
 */
export function buildFootnoteContentMap(
  footnotes: Footnote[],
  footnoteRefs: Array<{ footnoteId: number }>,
  contentWidth: number,
  options: ConvertFootnoteOptions
): Map<number, FootnoteContent> {
  const contentMap = new Map<number, FootnoteContent>();
  const footnoteById = new Map<number, Footnote>();

  for (const fn of footnotes) {
    if (fn.noteType === 'normal' || fn.noteType == null) {
      footnoteById.set(fn.id, fn);
    }
  }

  let displayNumber = 1;
  const seen = new Set<number>();

  for (const ref of footnoteRefs) {
    if (seen.has(ref.footnoteId)) continue;
    seen.add(ref.footnoteId);

    const footnote = footnoteById.get(ref.footnoteId);
    if (!footnote) continue;

    contentMap.set(
      ref.footnoteId,
      convertFootnoteToContent(footnote, displayNumber, contentWidth, options)
    );
    displayNumber++;
  }

  return contentMap;
}

// ============================================================================
// 4. Per-page footnote area height reservation
// ============================================================================

/**
 * Calculate per-page footnote reserved heights.
 * Returns Map<pageNumber, reservedHeight>.
 */
export function calculateFootnoteReservedHeights(
  pageFootnoteMap: Map<number, number[]>,
  footnoteContentMap: Map<number, { height: number }>
): Map<number, number> {
  const reserved = new Map<number, number>();

  for (const [pageNumber, footnoteIds] of pageFootnoteMap) {
    let totalHeight = 0;

    for (const fnId of footnoteIds) {
      const content = footnoteContentMap.get(fnId);
      if (content) {
        totalHeight += content.height;
      }
    }

    if (totalHeight > 0) {
      // Add separator height
      totalHeight += FOOTNOTE_SEPARATOR_HEIGHT;
      reserved.set(pageNumber, totalHeight);
    }
  }

  return reserved;
}

// ============================================================================
// 4b. Multi-pass footnote layout convergence
// ============================================================================

export interface StabilizeFootnoteLayoutArgs {
  blocks: FlowBlock[];
  measures: Measure[];
  layoutOpts: LayoutOptions;
  footnoteRefs: Array<{ footnoteId: number; pmPos: number }>;
  footnoteContentMap: Map<number, FootnoteContent>;
  /** First-pass layout already computed by the caller without reserved heights. */
  initialLayout: Layout;
}

export interface StabilizeFootnoteLayoutResult {
  layout: Layout;
  pageFootnoteMap: Map<number, number[]>;
  /** True if the loop converged before hitting MAX_FOOTNOTE_LAYOUT_PASSES. */
  converged: boolean;
}

/**
 * Run the multi-pass footnote layout loop. Reserving footnote space on a
 * page can move a reference to another page, which changes the reservation,
 * which can move references again. Iterate until the page→height contract
 * is the same one used by the latest layout, or `MAX_FOOTNOTE_LAYOUT_PASSES`
 * passes have run.
 *
 * Lives in core so the React + Vue adapters call the same loop and stay in
 * lockstep on convergence behaviour. Writes `page.footnoteIds` onto each
 * page in the returned layout so renderers can paint footnote areas.
 */
export function stabilizeFootnoteLayout(
  args: StabilizeFootnoteLayoutArgs
): StabilizeFootnoteLayoutResult {
  const { blocks, measures, layoutOpts, footnoteRefs, footnoteContentMap, initialLayout } = args;

  let pageFootnoteMap = mapFootnotesToPages(initialLayout.pages, footnoteRefs);
  let footnoteReservedHeights = calculateFootnoteReservedHeights(
    pageFootnoteMap,
    footnoteContentMap
  );

  if (footnoteReservedHeights.size === 0) {
    return { layout: initialLayout, pageFootnoteMap, converged: true };
  }

  let newLayout = initialLayout;
  let converged = false;
  for (let pass = 0; pass < MAX_FOOTNOTE_LAYOUT_PASSES; pass++) {
    newLayout = layoutDocument(blocks, measures, {
      ...layoutOpts,
      footnoteReservedHeights,
    });

    const nextPageFootnoteMap = mapFootnotesToPages(newLayout.pages, footnoteRefs);
    const nextFootnoteReservedHeights = calculateFootnoteReservedHeights(
      nextPageFootnoteMap,
      footnoteContentMap
    );

    pageFootnoteMap = nextPageFootnoteMap;
    if (footnoteReservedHeightsEqual(footnoteReservedHeights, nextFootnoteReservedHeights)) {
      footnoteReservedHeights = nextFootnoteReservedHeights;
      converged = true;
      break;
    }
    footnoteReservedHeights = nextFootnoteReservedHeights;
  }

  if (!converged) {
    newLayout = layoutDocument(blocks, measures, {
      ...layoutOpts,
      footnoteReservedHeights,
    });
    pageFootnoteMap = mapFootnotesToPages(newLayout.pages, footnoteRefs);
    console.warn(
      `[docx-editor] footnote layout did not stabilize within ${MAX_FOOTNOTE_LAYOUT_PASSES} passes; ` +
        'settling with best-effort reservation. If footnotes appear misplaced, please file a bug with the document.'
    );
  }

  for (const [pageNum, fnIds] of pageFootnoteMap) {
    const page = newLayout.pages.find((p) => p.number === pageNum);
    if (page) page.footnoteIds = fnIds;
  }

  return { layout: newLayout, pageFootnoteMap, converged };
}

// ============================================================================
// 5. Build per-page render items
// ============================================================================

/**
 * Turn the page→footnote-id map into the per-page render payload that
 * `renderPages` consumes via `footnotesByPage`. Skips non-`normal` notes
 * (separators, continuation notices), reads the display number out of the
 * content map, and pulls plain text via `getFootnoteText`.
 *
 * Lives in core (not in either adapter) so React + Vue both call the
 * same helper — same rule as the rest of this module.
 */
export function buildFootnoteRenderItems(
  pageFootnoteMap: Map<number, number[]>,
  footnoteContentMap: Map<number, FootnoteContent>,
  doc: Document | null
): Map<number, FootnoteRenderItem[]> {
  const result = new Map<number, FootnoteRenderItem[]>();
  if (!doc?.package?.footnotes) return result;

  const fnLookup = new Map<number, Footnote>();
  for (const fn of doc.package.footnotes) {
    if (fn.noteType && fn.noteType !== 'normal') continue;
    fnLookup.set(fn.id, fn);
  }

  for (const [pageNumber, footnoteIds] of pageFootnoteMap) {
    const items: FootnoteRenderItem[] = [];
    for (const fnId of footnoteIds) {
      const fn = fnLookup.get(fnId);
      if (!fn) continue;
      const content = footnoteContentMap.get(fnId);
      const displayNum = content?.displayNumber ?? 0;
      items.push({
        displayNumber: String(displayNum),
        text: getFootnoteText(fn),
        content,
      });
    }
    if (items.length > 0) result.set(pageNumber, items);
  }

  return result;
}
