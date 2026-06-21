/**
 * Document TextBox/anchored shape → PM textBox node + sibling-paragraph
 * extraction (Document → ProseMirror direction).
 *
 * Word stores text boxes as `<w:r><mc:AlternateContent><...><w:txbxContent>`
 * embedded in a host paragraph. This module pulls them out into sibling PM
 * `textBox` nodes (anchored before the host paragraph, in-flow ones after)
 * and converts the shape body into a paragraph-bearing PM node. The host
 * paragraph is dropped if extracting the text boxes leaves it empty.
 */

import type { Node as PMNode } from 'prosemirror-model';
import { schema } from '../../schema';
import type { Paragraph, ParagraphContent, Run, TextBox, Shape } from '../../../types/document';
import { emuToPixels } from '../../../docx/imageParser';
import type { StyleResolver } from '../../styles';
import { isAnchoredDocxTextBox, textBoxAnchorAttrsFromDocx } from '../textBoxAnchors';
import { convertParagraph } from './paragraph';

/**
 * Convert a paragraph block to PM nodes, extracting text boxes as sibling nodes.
 * Skips ghost empty paragraphs that only contained text box drawings.
 */
export function convertParagraphWithTextBoxes(
  block: Paragraph,
  styleResolver: StyleResolver | null
): PMNode[] {
  const textBoxes = extractTextBoxesFromParagraph(block);
  const pmParagraph = convertParagraph(block, styleResolver);
  const nodes: PMNode[] = [];
  const isEmptyAfterExtraction = textBoxes.length > 0 && pmParagraph.content.size === 0;
  const { anchored, inFlow } = partitionTextBoxesByAnchor(textBoxes);

  for (const tb of anchored) {
    nodes.push(convertTextBox(tb, styleResolver));
  }

  if (!isEmptyAfterExtraction) {
    nodes.push(pmParagraph);
  }

  for (const tb of inFlow) {
    nodes.push(convertTextBox(tb, styleResolver));
  }
  return nodes;
}

function partitionTextBoxesByAnchor(textBoxes: TextBox[]): {
  anchored: TextBox[];
  inFlow: TextBox[];
} {
  const anchored: TextBox[] = [];
  const inFlow: TextBox[] = [];

  for (const textBox of textBoxes) {
    if (isAnchoredDocxTextBox(textBox)) {
      anchored.push(textBox);
    } else {
      inFlow.push(textBox);
    }
  }

  return { anchored, inFlow };
}

/**
 * Collect a paragraph's {@link Run}s in document order, descending through the
 * same inline wrappers the parser does (`collectRunsThroughInlineWrappers` in
 * `blockContentParser.ts`). A text box can be anchored from a run nested inside
 * a content control, hyperlink, field, or tracked-change wrapper; if we only
 * looked at top-level runs (the prior behavior), the editor save path would
 * silently drop that text box. Mirror the wrapper set so the editor recovers
 * exactly what headless does.
 *
 * Wrappers expose their nested content under different keys: hyperlink uses
 * `children`, complexField uses `fieldResult`, everything else (inlineSdt,
 * simpleField, ins/del/moveFrom/moveTo) uses `content`. (`smartTag` and inline
 * `customXml` carry no model node — the parser flattens / skips them — so there
 * is nothing to recurse into for those.)
 */
function collectRunsThroughInlineWrappers(content: readonly ParagraphContent[]): Run[] {
  const runs: Run[] = [];
  for (const item of content) {
    if (item.type === 'run') {
      runs.push(item);
      continue;
    }
    const nested =
      (
        item as {
          content?: ParagraphContent[];
          children?: ParagraphContent[];
          fieldResult?: ParagraphContent[];
        }
      ).content ??
      (item as { children?: ParagraphContent[] }).children ??
      (item as { fieldResult?: ParagraphContent[] }).fieldResult;
    if (Array.isArray(nested)) {
      runs.push(...collectRunsThroughInlineWrappers(nested));
    }
  }
  return runs;
}

/**
 * Extract text boxes from paragraph runs.
 * Text boxes appear as ShapeContent where the shape has textBody,
 * or as DrawingContent that contains a text box instead of an image.
 *
 * Runs are collected through inline wrappers (content controls, hyperlinks,
 * fields, tracked-change wrappers) so a text box anchored from a nested run is
 * recovered, not just a top-level one.
 */
function extractTextBoxesFromParagraph(paragraph: Paragraph): TextBox[] {
  const textBoxes: TextBox[] = [];
  for (const run of collectRunsThroughInlineWrappers(paragraph.content)) {
    for (const rc of run.content) {
      if (rc.type === 'shape' && 'shape' in rc) {
        const shape = rc.shape as Shape;
        if (shape.textBody && shape.textBody.content.length > 0) {
          // Convert shape with text body to TextBox
          textBoxes.push({
            type: 'textBox',
            id: shape.id,
            size: shape.size,
            position: shape.position,
            wrap: shape.wrap,
            fill: shape.fill,
            outline: shape.outline,
            content: shape.textBody.content,
            margins: shape.textBody.margins,
          });
        }
      }
    }
  }
  return textBoxes;
}

/**
 * Convert a TextBox to a ProseMirror textBox node
 */
function convertTextBox(textBox: TextBox, styleResolver: StyleResolver | null): PMNode {
  const widthPx = textBox.size?.width ? emuToPixels(textBox.size.width) : 200;
  const heightPx = textBox.size?.height ? emuToPixels(textBox.size.height) : undefined;

  // Convert fill color
  let fillColor: string | undefined;
  if (textBox.fill?.color?.rgb) {
    fillColor = `#${textBox.fill.color.rgb}`;
  }

  // Convert outline
  let outlineWidth: number | undefined;
  let outlineColor: string | undefined;
  let outlineStyle: string | undefined;
  if (textBox.outline && textBox.outline.width) {
    outlineWidth = Math.round((textBox.outline.width / 914400) * 96 * 100) / 100;
    if (textBox.outline.color?.rgb) {
      outlineColor = `#${textBox.outline.color.rgb}`;
    }
    outlineStyle = textBox.outline.style || 'solid';
  }

  // Convert margins from EMU to pixels
  const marginTop = textBox.margins?.top != null ? emuToPixels(textBox.margins.top) : 4;
  const marginBottom = textBox.margins?.bottom != null ? emuToPixels(textBox.margins.bottom) : 4;
  const marginLeft = textBox.margins?.left != null ? emuToPixels(textBox.margins.left) : 7;
  const marginRight = textBox.margins?.right != null ? emuToPixels(textBox.margins.right) : 7;

  // Convert text box content (paragraphs) to PM nodes
  const contentNodes: PMNode[] = [];
  for (const para of textBox.content) {
    contentNodes.push(convertParagraph(para, styleResolver));
  }

  // Ensure at least one paragraph
  if (contentNodes.length === 0) {
    contentNodes.push(schema.node('paragraph', {}, []));
  }

  return schema.node(
    'textBox',
    {
      width: widthPx,
      height: heightPx,
      textBoxId: textBox.id,
      fillColor,
      outlineWidth,
      outlineColor,
      outlineStyle,
      marginTop,
      marginBottom,
      marginLeft,
      marginRight,
      ...textBoxAnchorAttrsFromDocx(textBox),
    },
    contentNodes
  );
}
