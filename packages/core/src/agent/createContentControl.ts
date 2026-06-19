/**
 * Create a new content control (`w:sdt`) by wrapping existing content.
 *
 * Complements the discovery/edit functions in {@link ./contentControls}: where
 * those find and mutate *existing* controls, this wraps a text span (inline) or
 * a contiguous block range (block-level) in a *new* control with a synthesized,
 * Word-correct `w:sdtPr`. Pure — the input {@link Document} is not mutated.
 */

import type {
  Document,
  BlockContent,
  BlockSdt,
  InlineSdt,
  Paragraph,
  Run,
  RunContent,
  SdtType,
  SdtProperties,
} from '../types/document';
import { getParagraphText, getRunText, getHyperlinkText } from './text-utils';
import { synthesizeSdtPr } from '../docx/serializer/paragraphSerializer/content';
import {
  rebuild,
  findContentControl,
  findContentControls,
  isInlineContent,
  type ContentControlInfo,
  type ContentControlAddress,
  type ContentControlStep,
} from './contentControls';

/** A create request failed: the target couldn't be resolved, or the wrap is invalid. */
export class ContentControlCreateError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ContentControlCreateError';
  }
}

/**
 * Where to place a new content control.
 *
 * - `text` — wrap an exact text span **inside a paragraph** (an inline control).
 *   Locate the paragraph by Word `w14:paraId` (`paraId`) or by a structural
 *   {@link ContentControlAddress} (`paragraph`, as returned by
 *   {@link findContentControls}); then wrap the chosen `occurrence` of `text`
 *   (matched against {@link getParagraphText getParagraphText(para)}).
 * - `blocks` — wrap a contiguous run of block-level content (paragraphs/tables)
 *   from `from` through `to` (inclusive; `to` defaults to `from`) in a
 *   block-level control. Both addresses must share the same container and must
 *   not resolve inside a table cell (the model can't hold a block SDT there —
 *   use `kind: 'text'`).
 *
 * Targets resolve within the **body** story; header/footer create is not
 * supported (throws).
 */
export type CreateContentControlTarget =
  | {
      kind: 'text';
      /** Locate the paragraph by Word `w14:paraId`. */
      paraId?: string;
      /** Locate the paragraph by a structural address (alternative to `paraId`). */
      paragraph?: ContentControlAddress;
      /** Exact substring to wrap, matched against `getParagraphText(para)`. */
      text: string;
      /** Which occurrence of `text` to wrap when it repeats (1-based; default 1). */
      occurrence?: number;
    }
  | {
      kind: 'blocks';
      /** First block to wrap (a block-level structural address). */
      from: ContentControlAddress;
      /** Last block to wrap, inclusive (defaults to `from` — wrap a single block). */
      to?: ContentControlAddress;
    };

/** Modeled properties for a control created by {@link createContentControl}. */
export interface NewContentControlProps {
  /** Control type (default `richText`). */
  sdtType?: SdtType;
  /** Developer identifier (`w:tag`). */
  tag?: string;
  /** Friendly name (`w:alias`). */
  alias?: string;
  /** Numeric id (`w:id`). Default: auto-assigned, unique across the document. */
  id?: number;
  /** Lock setting (`w:lock`). */
  lock?: SdtProperties['lock'];
  /** Dropdown/combobox list items. */
  listItems?: { displayText: string; value: string }[];
  /** Date display format (`w:date/w:dateFormat`), for `date` controls. */
  dateFormat?: string;
  /** Initial checkbox state, for `checkbox` controls. */
  checked?: boolean;
  /** Whether the control starts in placeholder state (`w:showingPlcHdr`). */
  showingPlaceholder?: boolean;
}

/** Length an inline item contributes to {@link getParagraphText} (runs + hyperlinks only). */
function itemTextLength(item: Paragraph['content'][number]): number {
  if (item.type === 'run') return getRunText(item).length;
  if (item.type === 'hyperlink') return getHyperlinkText(item).length;
  return 0;
}

/** A `w:t` text node sliced to `[a, b)`, preserving boundary whitespace. */
function sliceText(node: Extract<RunContent, { type: 'text' }>, a: number, b?: number): RunContent {
  const text = node.text.slice(a, b);
  const next: Extract<RunContent, { type: 'text' }> = { type: 'text', text };
  if (node.preserveSpace || /^\s|\s$/.test(text)) next.preserveSpace = true;
  return next;
}

/**
 * Split a run's content at text-character offset `n` into `[left, right]` runs,
 * preserving the run's formatting and tracked property changes. Non-text run
 * content (tabs, breaks, drawings) lands on the side its position falls on.
 */
function splitRunAtChar(run: Run, n: number): [Run, Run] {
  const left: RunContent[] = [];
  const right: RunContent[] = [];
  let count = 0;
  for (const c of run.content) {
    if (c.type === 'text') {
      if (count >= n) {
        right.push(c);
      } else if (count + c.text.length <= n) {
        left.push(c);
        count += c.text.length;
      } else {
        const k = n - count;
        left.push(sliceText(c, 0, k));
        right.push(sliceText(c, k));
        count += c.text.length;
      }
    } else {
      (count >= n ? right : left).push(c);
    }
  }
  return [
    { ...run, content: left },
    { ...run, content: right },
  ];
}

/**
 * Wrap the character range `[start, end)` of a paragraph's inline `content` in
 * `sdt`, splitting runs at the boundaries. Items fully inside the range become
 * the control's content (preserving interior fields/tabs/breaks wholesale);
 * only an item straddling a boundary must be a run (to split).
 *
 * Throws {@link ContentControlCreateError} if a boundary falls inside a non-run
 * item (hyperlink/field), if the range overlaps an existing inline control, or
 * if the range contains content that cannot live inside an inline control
 * (bookmarks, comment ranges, tracked-change wrappers).
 */
function wrapInlineSpan(
  content: Paragraph['content'],
  start: number,
  end: number,
  sdt: InlineSdt
): Paragraph['content'] {
  const before: Paragraph['content'] = [];
  const matched: InlineSdt['content'] = [];
  const after: Paragraph['content'] = [];
  let pos = 0;
  for (const item of content) {
    const len = itemTextLength(item);
    const iStart = pos;
    const iEnd = pos + len;
    pos = iEnd;
    if (iEnd <= start) {
      before.push(item);
      continue;
    }
    if (iStart >= end) {
      after.push(item);
      continue;
    }
    if (iStart >= start && iEnd <= end) {
      // Fully inside the span.
      if (item.type === 'inlineSdt') {
        throw new ContentControlCreateError(
          'The text span overlaps an existing inline content control; cannot wrap it.'
        );
      }
      if (!isInlineContent(item)) {
        throw new ContentControlCreateError(
          `The text span contains a '${item.type}' that cannot be placed inside an inline content control.`
        );
      }
      matched.push(item);
      continue;
    }
    // Straddles a boundary: must be a run to split cleanly.
    if (item.type !== 'run') {
      throw new ContentControlCreateError(
        item.type === 'inlineSdt'
          ? 'The text span overlaps an existing inline content control; cannot wrap it.'
          : `A content-control boundary falls inside a '${item.type}'; wrap on a run boundary instead.`
      );
    }
    const a = Math.max(start, iStart) - iStart;
    const b = Math.min(end, iEnd) - iStart;
    const [leftRun, rest] = splitRunAtChar(item, a);
    const [midRun, rightRun] = splitRunAtChar(rest, b - a);
    if (leftRun.content.length > 0) before.push(leftRun);
    matched.push(midRun);
    if (rightRun.content.length > 0) after.push(rightRun);
  }
  return [...before, { ...sdt, content: matched }, ...after];
}

/** Highest `w:id` across every control in the document (body + header/footer). */
function maxControlId(doc: Document): number {
  return findContentControls(doc, {}, { scope: 'all' }).reduce(
    (max, c) => (c.id != null && c.id > max ? c.id : max),
    0
  );
}

/** Build the {@link SdtProperties} for a created control, with a synthesized raw `w:sdtPr`. */
function buildCreatedProps(props: NewContentControlProps, id: number): SdtProperties {
  const out: SdtProperties = { sdtType: props.sdtType ?? 'richText', id };
  if (props.tag != null) out.tag = props.tag;
  if (props.alias != null) out.alias = props.alias;
  if (props.lock != null) out.lock = props.lock;
  if (props.listItems != null) out.listItems = props.listItems;
  if (props.dateFormat != null) out.dateFormat = props.dateFormat;
  if (props.checked != null) out.checked = props.checked;
  if (props.showingPlaceholder) out.showingPlaceholder = true;
  out.rawPropertiesXml = synthesizeSdtPr(out);
  return out;
}

/**
 * Rebuild body `blocks`, applying `transform` to the FIRST paragraph matching
 * `predicate` (descending block SDTs and table cells), immutably. `state.done`
 * stops the walk after the first match. Returns the same reference if unchanged.
 */
function mapFirstParagraph(
  blocks: BlockContent[],
  predicate: (p: Paragraph) => boolean,
  transform: (p: Paragraph) => Paragraph,
  state: { done: boolean }
): BlockContent[] {
  let changed = false;
  const out: BlockContent[] = [];
  for (const block of blocks) {
    if (state.done) {
      out.push(block);
      continue;
    }
    if (block.type === 'paragraph') {
      if (predicate(block)) {
        out.push(transform(block));
        state.done = true;
        changed = true;
      } else {
        out.push(block);
      }
    } else if (block.type === 'blockSdt') {
      const next = mapFirstParagraph(block.content, predicate, transform, state);
      if (next !== block.content) {
        out.push({ ...block, content: next });
        changed = true;
      } else {
        out.push(block);
      }
    } else if (block.type === 'table') {
      let tableChanged = false;
      const rows = block.rows.map((row) => {
        if (state.done) return row;
        let rowChanged = false;
        const cells = row.cells.map((cell) => {
          if (state.done) return cell;
          const cellContent = cell.content as BlockContent[];
          const next = mapFirstParagraph(cellContent, predicate, transform, state);
          if (next !== cellContent) {
            rowChanged = true;
            return { ...cell, content: next as typeof cell.content };
          }
          return cell;
        });
        if (rowChanged) {
          tableChanged = true;
          return { ...row, cells };
        }
        return row;
      });
      if (tableChanged) {
        out.push({ ...block, rows });
        changed = true;
      } else {
        out.push(block);
      }
    } else {
      out.push(block);
    }
  }
  return changed ? out : blocks;
}

/** Navigate a body structural address to the paragraph it resolves to (or throw). */
function paragraphAtAddress(doc: Document, address: ContentControlAddress): Paragraph {
  if (address.location.part !== 'body') {
    throw new ContentControlCreateError(
      'createContentControl only supports targets in the document body.'
    );
  }
  let blocks: BlockContent[] = doc.package.document.content;
  let node: BlockContent | undefined;
  for (let i = 0; i < address.steps.length; i++) {
    const step = address.steps[i];
    if (step.kind === 'block') {
      node = blocks[step.index];
      if (i < address.steps.length - 1) {
        if (node?.type === 'blockSdt') blocks = node.content;
        else throw new ContentControlCreateError('Address does not resolve to a paragraph.');
      }
    } else if (step.kind === 'cell') {
      if (node?.type !== 'table') {
        throw new ContentControlCreateError('Address does not resolve to a paragraph.');
      }
      const cell = node.rows[step.row]?.cells[step.col];
      if (!cell) throw new ContentControlCreateError('Address cell coordinate out of range.');
      blocks = cell.content as BlockContent[];
    } else {
      throw new ContentControlCreateError('Address does not resolve to a paragraph.');
    }
  }
  if (node?.type !== 'paragraph') {
    throw new ContentControlCreateError('Address does not resolve to a paragraph.');
  }
  return node;
}

/** Resolve, locate, split, and wrap a text span in a paragraph (inline create). */
function createInlineControl(
  doc: Document,
  target: Extract<CreateContentControlTarget, { kind: 'text' }>,
  props: SdtProperties
): Document {
  if (!target.text) {
    throw new ContentControlCreateError('`text` must be a non-empty string.');
  }
  const occurrence = target.occurrence ?? 1;
  if (occurrence < 1) {
    throw new ContentControlCreateError('`occurrence` must be >= 1.');
  }

  // Build the transform that locates the span and wraps it.
  const transform = (para: Paragraph): Paragraph => {
    const full = getParagraphText(para);
    // Find the requested occurrence of the exact substring.
    let from = -1;
    let count = 0;
    for (
      let i = full.indexOf(target.text);
      i !== -1;
      i = full.indexOf(target.text, i + target.text.length)
    ) {
      if (++count === occurrence) {
        from = i;
        break;
      }
    }
    if (from === -1) {
      throw new ContentControlCreateError(
        count === 0
          ? `Text ${JSON.stringify(target.text)} was not found in the paragraph.`
          : `Occurrence ${occurrence} of ${JSON.stringify(target.text)} not found (only ${count} present).`
      );
    }
    const sdt: InlineSdt = { type: 'inlineSdt', properties: props, content: [] };
    return { ...para, content: wrapInlineSpan(para.content, from, from + target.text.length, sdt) };
  };

  // Resolve the paragraph by address or paraId, rebuilding the body immutably.
  if (target.paragraph) {
    const para = paragraphAtAddress(doc, target.paragraph);
    const state = { done: false };
    const content = mapFirstParagraph(
      doc.package.document.content,
      (p) => p === para,
      transform,
      state
    );
    if (!state.done)
      throw new ContentControlCreateError('Address does not resolve to a paragraph.');
    return rebuild(doc, content);
  }
  if (target.paraId != null) {
    const state = { done: false };
    const content = mapFirstParagraph(
      doc.package.document.content,
      (p) => p.paraId === target.paraId,
      transform,
      state
    );
    if (!state.done) {
      throw new ContentControlCreateError(
        `No paragraph found with paraId ${JSON.stringify(target.paraId)}.`
      );
    }
    return rebuild(doc, content);
  }
  throw new ContentControlCreateError('A `text` target needs either `paraId` or `paragraph`.');
}

/** Field-wise equality of two structural step paths (order-independent of key serialization). */
function stepsEqual(a: ContentControlStep[], b: ContentControlStep[]): boolean {
  if (a.length !== b.length) return false;
  return a.every((s, i) => {
    const t = b[i];
    if (s.kind !== t.kind) return false;
    if (s.kind === 'block') return s.index === (t as typeof s).index;
    if (s.kind === 'cell') return s.row === (t as typeof s).row && s.col === (t as typeof s).col;
    return s.index === (t as typeof s).index; // 'inline'
  });
}

/** Replace `blocks[from..to]` with a single {@link BlockSdt}, descending `containerPath` immutably. */
function wrapBlockRange(
  blocks: BlockContent[],
  containerPath: ContentControlStep[],
  from: number,
  to: number,
  properties: SdtProperties
): BlockContent[] {
  if (containerPath.length === 0) {
    if (from < 0 || to >= blocks.length) {
      throw new ContentControlCreateError('Block range is out of bounds.');
    }
    const wrapper: BlockSdt = { type: 'blockSdt', properties, content: blocks.slice(from, to + 1) };
    return [...blocks.slice(0, from), wrapper, ...blocks.slice(to + 1)];
  }
  const [step, ...rest] = containerPath;
  if (step.kind !== 'block') {
    throw new ContentControlCreateError('Block create cannot target content inside a table cell.');
  }
  const block = blocks[step.index];
  if (block?.type !== 'blockSdt') {
    throw new ContentControlCreateError('Block range does not resolve to a block container.');
  }
  const nextContent = wrapBlockRange(block.content, rest, from, to, properties);
  return blocks.map((b, i) => (i === step.index ? { ...block, content: nextContent } : b));
}

/** Wrap a contiguous block range in a block-level control (block create). */
function createBlockControl(
  doc: Document,
  target: Extract<CreateContentControlTarget, { kind: 'blocks' }>,
  properties: SdtProperties
): Document {
  const from = target.from;
  const to = target.to ?? target.from;
  if (from.location.part !== 'body' || to.location.part !== 'body') {
    throw new ContentControlCreateError(
      'createContentControl only supports targets in the document body.'
    );
  }
  if (from.steps.some((s) => s.kind === 'cell') || to.steps.some((s) => s.kind === 'cell')) {
    throw new ContentControlCreateError(
      "Cannot wrap blocks inside a table cell in a block-level control. Use { kind: 'text' } for inline controls."
    );
  }
  const fromLast = from.steps[from.steps.length - 1];
  const toLast = to.steps[to.steps.length - 1];
  if (fromLast?.kind !== 'block' || toLast?.kind !== 'block') {
    throw new ContentControlCreateError('Block range addresses must end at a block index.');
  }
  const fromContainer = from.steps.slice(0, -1);
  const toContainer = to.steps.slice(0, -1);
  if (!stepsEqual(fromContainer, toContainer)) {
    throw new ContentControlCreateError('Block range crosses container boundaries.');
  }
  if (fromLast.index > toLast.index) {
    throw new ContentControlCreateError('Block range `from` must not come after `to`.');
  }
  const content = wrapBlockRange(
    doc.package.document.content,
    fromContainer,
    fromLast.index,
    toLast.index,
    properties
  );
  return rebuild(doc, content);
}

/**
 * Wrap existing content in a **new** content control (`w:sdt`), returning a new
 * {@link Document} and the created {@link ContentControlInfo}. Pure — the input
 * is not mutated.
 *
 * Two target shapes (see {@link CreateContentControlTarget}):
 *
 * - `{ kind: 'text', … }` wraps an exact text span inside a paragraph in an
 *   **inline** control — the form needed inside table cells and mid-sentence,
 *   where block controls aren't allowed. Runs are split at the span boundaries
 *   (formatting preserved); interior fields/tabs/breaks are kept wholesale.
 * - `{ kind: 'blocks', … }` wraps a contiguous run of block-level content in a
 *   **block** control. Targets inside a table cell are rejected (the model can't
 *   hold a block SDT there) — use `kind: 'text'` instead.
 *
 * The control's `w:sdtPr` is synthesized from `props` (Word-correct element
 * order; a `date` control writes its format to `<w:dateFormat>`, not
 * `@w:fullDate`). The `w:id` is auto-assigned (max existing across the document
 * + 1) when `props.id` is omitted, and the control round-trips so
 * `findContentControl(doc, { tag })` resolves it after a save/reload.
 *
 * @throws {@link ContentControlCreateError} when the target can't be resolved,
 * the text/occurrence isn't found, the span overlaps an existing control or
 * crosses a non-run boundary, or a block target resolves inside a table cell.
 */
export function createContentControl(
  doc: Document,
  target: CreateContentControlTarget,
  props: NewContentControlProps = {},
  options: { force?: boolean } = {}
): { doc: Document; control: ContentControlInfo } {
  // Resolve a document-unique `w:id`. A caller-supplied id that collides is
  // rejected (a duplicate `w:id` is invalid OOXML); with `{ force: true }` we
  // auto-assign a fresh unique id instead of erroring. Keeping the id unique
  // also guarantees the find-back below resolves the control we just created.
  let id = props.id ?? maxControlId(doc) + 1;
  if (props.id != null && findContentControl(doc, { id: props.id }, { scope: 'all' })) {
    if (!options.force) {
      throw new ContentControlCreateError(
        `A content control with id ${props.id} already exists. Pass { force: true } to auto-assign a unique id instead.`
      );
    }
    id = maxControlId(doc) + 1;
  }
  const properties = buildCreatedProps(props, id);

  const nextDoc =
    target.kind === 'text'
      ? createInlineControl(doc, target, properties)
      : createBlockControl(doc, target, properties);

  const control = findContentControl(nextDoc, { id }, { scope: 'all' });
  if (!control) {
    // Should be unreachable: the control was just inserted with this unique id.
    throw new ContentControlCreateError('Created control could not be located after insertion.');
  }
  return { doc: nextDoc, control };
}
