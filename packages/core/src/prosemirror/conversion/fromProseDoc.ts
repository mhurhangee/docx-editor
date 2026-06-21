/**
 * ProseMirror to Document Conversion
 *
 * Converts a ProseMirror document back to our Document type.
 * This enables round-trip editing: DOCX -> Document -> PM -> Document -> DOCX
 *
 * Key responsibilities:
 * - Coalesce consecutive text with same marks into single Runs
 * - Preserve paragraph attributes (paraId, textId, formatting)
 * - Handle marks -> TextFormatting conversion
 *
 * This file owns the top-level orchestrator (`fromProseDoc`) plus block
 * extraction and the page-break paragraph factory. Per-domain converters
 * live under ./fromProseDoc/ (marks, runs, paragraph, tables, textbox).
 * The deep import `@eigenpal/.../prosemirror/conversion/fromProseDoc` is
 * a tsup entry consumed by the Vue adapter — the barrel re-exports
 * preserve that surface.
 * @packageDocumentation
 * @public
 */

import type { Node as PMNode } from 'prosemirror-model';
import type {
  Document,
  DocumentBody,
  Paragraph,
  Run,
  BreakContent,
  BlockContent,
  BlockSdt,
  Watermark,
} from '../../types/document';
import { getDocumentWatermark, setDocumentWatermark } from '../../docx/watermarkApi';
import type { TextBoxAttrs } from '../extensions/nodes/TextBoxExtension';
import { shouldExportTextBoxInsideFollowingParagraph } from './textBoxAnchors';
import { sdtAttrsToProps } from './sdtAttrs';
import { convertPMParagraph } from './fromProseDoc/paragraph';
import { convertPMTable } from './fromProseDoc/tables';
import { convertPMTextBox, convertPMTextBoxRun } from './fromProseDoc/textbox';
import { collectNumberingFromPM } from './fromProseDoc/numbering';

/**
 * Convert a ProseMirror document to our Document type
 */
export function fromProseDoc(pmDoc: PMNode, baseDocument?: Document): Document {
  const blocks = extractBlocks(pmDoc);

  // Preserve section properties (margins, headers, footers) from base document
  const documentBody: DocumentBody = {
    content: blocks,
    finalSectionProperties: baseDocument?.package.document.finalSectionProperties,
    sections: baseDocument?.package.document.sections,
    comments: baseDocument?.package.document.comments,
  };

  // If we have a base document, preserve its package structure
  const result: Document = baseDocument
    ? {
        ...baseDocument,
        package: {
          ...baseDocument.package,
          document: documentBody,
        },
      }
    : { package: { document: documentBody, numbering: collectNumberingFromPM(pmDoc) } };

  // Rebalance bookmark markers so every id has matching start/end counts. The
  // PM `bookmarks` attr fabricates a balanced inline pair for each inline start,
  // and "lone" inline ends + block markers are carried separately; a straddling
  // bookmark can leave an id with more ends than starts (or vice versa). This
  // pass trims the excess so no orphan marker reaches the serializer.
  rebalanceBookmarkMarkers(result);

  // Sync the watermark doc attr → `HeaderFooter.watermark` so the serializer
  // and any model consumers see watermark applies/removes (incl. via undo).
  // Reference comparison skips the header-map clone on the common no-change
  // path — the same Watermark object rides PM attrs until explicitly changed.
  const attrWatermark = (pmDoc.attrs.watermark as Watermark | null) ?? null;
  const currentWatermark = getDocumentWatermark(result) ?? null;
  if (attrWatermark !== currentWatermark) {
    return setDocumentWatermark(result, attrWatermark);
  }
  return result;
}

/**
 * A block (paragraph / table / block-SDT) that may carry block-level bookmark
 * markers riding alongside it for round-trip. See {@link wrapBlockMarkers}.
 */
type BlockWithMarkers = {
  leadingBlockMarkers?: Paragraph['leadingBlockMarkers'];
  trailingBlockMarkers?: Paragraph['trailingBlockMarkers'];
};

/**
 * Make every bookmark id balanced (`bookmarkStart` count == `bookmarkEnd`
 * count) document-wide by trimming surplus ends. Three sources can each emit an
 * end for the same id:
 *   1. the `bookmarks` attr fabrication (a balanced inline pair per inline start),
 *   2. a carried "lone" inline end (`loneBookmarkEndIds` — a cross-paragraph or
 *      block-anchored close), and
 *   3. a carried block-level end marker (`leadingBlockMarkers`/`trailingBlockMarkers`).
 *
 * A straddling bookmark therefore over- or under-emits:
 *   • inline start + block end  → fabricated end + block end = 2 ends (trim 1);
 *   • inline start in para A + inline end in para B → fabricated end + carried
 *     lone end = 2 ends (trim 1);
 *   • block start + inline end → block start + carried lone end = balanced (keep).
 *
 * Genuinely block-level bookmarks (block start + block end, the FIX-B case) stay
 * balanced and are never touched. We only ever REMOVE surplus ends, so a real
 * start is never orphaned. Trim order prefers block-marker ends, then inline
 * content ends, so the most redundant carrier goes first.
 */
function rebalanceBookmarkMarkers(doc: Document): void {
  const startCount = new Map<number, number>();
  const endCount = new Map<number, number>();
  const bump = (m: Map<number, number>, id: number): void => {
    m.set(id, (m.get(id) ?? 0) + 1);
  };

  const countBlock = (block: BlockContent): void => {
    const withMarkers = block as BlockWithMarkers;
    for (const marker of withMarkers.leadingBlockMarkers ?? []) {
      bump(marker.type === 'bookmarkStart' ? startCount : endCount, marker.id);
    }
    for (const marker of withMarkers.trailingBlockMarkers ?? []) {
      bump(marker.type === 'bookmarkStart' ? startCount : endCount, marker.id);
    }
    if (block.type === 'paragraph') {
      for (const item of block.content) {
        if (item.type === 'bookmarkStart') bump(startCount, item.id);
        else if (item.type === 'bookmarkEnd') bump(endCount, item.id);
      }
    }
  };

  const walkBlocks = (
    blocks: BlockContent[] | undefined,
    visit: (b: BlockContent) => void
  ): void => {
    if (!blocks) return;
    for (const block of blocks) {
      visit(block);
      if (block.type === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) walkBlocks(cell.content as BlockContent[], visit);
        }
      } else if (block.type === 'blockSdt') {
        walkBlocks(block.content, visit);
      }
    }
  };
  const walkAll = (visit: (b: BlockContent) => void): void => {
    walkBlocks(doc.package.document.content, visit);
    for (const section of doc.package.document.sections ?? []) walkBlocks(section.content, visit);
    for (const hf of doc.package.headers?.values() ?? []) walkBlocks(hf.content, visit);
    for (const hf of doc.package.footers?.values() ?? []) walkBlocks(hf.content, visit);
  };

  // Pass 1 — tally starts and ends per id across the whole document.
  walkAll(countBlock);

  // Per-id budget of ends we must drop (ends beyond the matching start count).
  const surplus = new Map<number, number>();
  for (const [id, ends] of endCount) {
    const extra = ends - (startCount.get(id) ?? 0);
    if (extra > 0) surplus.set(id, extra);
  }
  if (surplus.size === 0) return;

  // Pass 2a — drop surplus block-marker ends first (the most redundant carrier).
  const filterMarkers = (
    markers: BlockWithMarkers['leadingBlockMarkers']
  ): BlockWithMarkers['leadingBlockMarkers'] | undefined => {
    if (!markers || markers.length === 0) return markers;
    const kept: NonNullable<BlockWithMarkers['leadingBlockMarkers']> = [];
    let changed = false;
    for (const marker of markers) {
      const budget = marker.type === 'bookmarkEnd' ? (surplus.get(marker.id) ?? 0) : 0;
      if (budget > 0) {
        surplus.set(marker.id, budget - 1);
        changed = true;
        continue;
      }
      kept.push(marker);
    }
    return changed ? kept : markers;
  };
  walkAll((block) => {
    const withMarkers = block as BlockWithMarkers;
    const leading = filterMarkers(withMarkers.leadingBlockMarkers);
    if (leading && leading.length > 0) withMarkers.leadingBlockMarkers = leading;
    else delete withMarkers.leadingBlockMarkers;
    const trailing = filterMarkers(withMarkers.trailingBlockMarkers);
    if (trailing && trailing.length > 0) withMarkers.trailingBlockMarkers = trailing;
    else delete withMarkers.trailingBlockMarkers;
  });

  // Pass 2b — drop any still-surplus inline content ends (relocated duplicates).
  let remaining = 0;
  for (const n of surplus.values()) remaining += n;
  if (remaining === 0) return;
  walkAll((block) => {
    if (block.type !== 'paragraph') return;
    let changed = false;
    const kept: typeof block.content = [];
    for (const item of block.content) {
      if (item.type === 'bookmarkEnd') {
        const budget = surplus.get(item.id) ?? 0;
        if (budget > 0) {
          surplus.set(item.id, budget - 1);
          changed = true;
          continue;
        }
      }
      kept.push(item);
    }
    if (changed) block.content = kept;
  });
}

/**
 * Extract blocks (paragraphs, tables, and block-level SDTs) from a
 * ProseMirror document or block-containing node.
 */
function extractBlocks(pmDoc: PMNode): BlockContent[] {
  const blocks: BlockContent[] = [];
  let pendingAnchoredTextBoxRuns: Run[] = [];

  const flushPendingTextBoxes = (): void => {
    for (const run of pendingAnchoredTextBoxRuns) {
      blocks.push({
        type: 'paragraph',
        content: [run],
      });
    }
    pendingAnchoredTextBoxRuns = [];
  };

  pmDoc.forEach((node) => {
    if (node.type.name === 'paragraph') {
      const paragraph = convertPMParagraph(node);
      if (pendingAnchoredTextBoxRuns.length > 0) {
        paragraph.content = [...pendingAnchoredTextBoxRuns, ...paragraph.content];
        pendingAnchoredTextBoxRuns = [];
      }
      blocks.push(paragraph);
    } else if (node.type.name === 'table') {
      flushPendingTextBoxes();
      blocks.push(convertPMTable(node));
    } else if (node.type.name === 'blockSdt') {
      flushPendingTextBoxes();
      blocks.push(convertPMBlockSdt(node));
    } else if (node.type.name === 'textBox') {
      const attrs = node.attrs as TextBoxAttrs;
      if (shouldExportTextBoxInsideFollowingParagraph(attrs)) {
        pendingAnchoredTextBoxRuns.push(convertPMTextBoxRun(node));
      } else {
        flushPendingTextBoxes();
        blocks.push(convertPMTextBox(node));
      }
    } else if (node.type.name === 'pageBreak') {
      flushPendingTextBoxes();
      // Convert page break node to a paragraph with a page break run
      blocks.push(createPageBreakParagraph());
    }
  });

  flushPendingTextBoxes();

  return blocks;
}

/**
 * Reconstruct a {@link BlockSdt} model node from a `blockSdt` PM node:
 * project the attrs back to {@link SdtProperties} (the captured raw `sdtPr`
 * rides along for lossless serialization) and recurse into the children.
 */
function convertPMBlockSdt(node: PMNode): BlockSdt {
  const attrs = node.attrs as Record<string, unknown>;
  const blockSdt: BlockSdt = {
    type: 'blockSdt',
    properties: sdtAttrsToProps(attrs),
    content: extractBlocks(node),
  };
  // Round-trip block-level bookmark markers carried as opaque attrs. The
  // serializer re-emits them around this control's `w:sdt` via
  // `wrapBlockMarkers`.
  const leading = attrs.leadingBlockMarkers as BlockSdt['leadingBlockMarkers'];
  const trailing = attrs.trailingBlockMarkers as BlockSdt['trailingBlockMarkers'];
  if (leading && leading.length > 0) blockSdt.leadingBlockMarkers = leading;
  if (trailing && trailing.length > 0) blockSdt.trailingBlockMarkers = trailing;
  return blockSdt;
}

/**
 * Create a paragraph containing only a page break run (for DOCX serialization)
 */
function createPageBreakParagraph(): Paragraph {
  const breakContent: BreakContent = { type: 'break', breakType: 'page' };
  const run: Run = { type: 'run', content: [breakContent] };
  return {
    type: 'paragraph',
    content: [run],
  };
}

/**
 * Update a Document with content from a ProseMirror document
 * Preserves all non-content parts of the original document
 */
export function updateDocumentContent(originalDocument: Document, pmDoc: PMNode): Document {
  return fromProseDoc(pmDoc, originalDocument);
}

/**
 * Convert a ProseMirror document back to an array of Paragraph/Table blocks.
 * Used for converting edited header/footer PM content back to the document model.
 */
export function proseDocToBlocks(pmDoc: PMNode): BlockContent[] {
  return extractBlocks(pmDoc);
}
