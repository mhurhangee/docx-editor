/**
 * Paragraph Formatting Commands — thin re-exports from extension system
 *
 * Alignment, line spacing, indentation, lists, paragraph styles.
 * All implementations live in extensions/; this file re-exports
 * for backward compatibility.
 * @packageDocumentation
 * @public
 */

import type { Command } from 'prosemirror-state';
import type {
  ParagraphAlignment,
  LineSpacingRule,
  TabStopAlignment,
  TabLeader,
} from '../../types/document';
import { singletonManager } from '../schema';

// Re-export types and query helpers from extensions
import type { ResolvedStyleAttrs, GenerateTOCOptions } from '../extensions/core/ParagraphExtension';
export type { ResolvedStyleAttrs, GenerateTOCOptions } from '../extensions/core/ParagraphExtension';
export {
  getParagraphAlignment,
  getStyleId,
  getParagraphTabs,
  getParagraphBidi,
} from '../extensions/core/ParagraphExtension';
export { isInList, getListInfo } from '../extensions/features/ListExtension';

// ============================================================================
// COMMANDS — delegated to singleton extension manager
// ============================================================================

const cmds = singletonManager.getCommands();

// Alignment
export function setAlignment(alignment: ParagraphAlignment): Command {
  return cmds.setAlignment(alignment);
}
export const alignLeft: Command = cmds.alignLeft();
export const alignCenter: Command = cmds.alignCenter();
export const alignRight: Command = cmds.alignRight();
export const alignJustify: Command = cmds.alignJustify();

// Line spacing
export function setLineSpacing(value: number, rule: LineSpacingRule = 'auto'): Command {
  return cmds.setLineSpacing(value, rule);
}
export const singleSpacing: Command = cmds.singleSpacing();
export const oneAndHalfSpacing: Command = cmds.oneAndHalfSpacing();
export const doubleSpacing: Command = cmds.doubleSpacing();

// Indentation
export function increaseIndent(amount: number = 720): Command {
  return cmds.increaseIndent(amount);
}
export function decreaseIndent(amount: number = 720): Command {
  return cmds.decreaseIndent(amount);
}
export function setIndentLeft(twips: number): Command {
  return cmds.setIndentLeft(twips);
}
export function setIndentRight(twips: number): Command {
  return cmds.setIndentRight(twips);
}
export function setIndentFirstLine(twips: number, hanging?: boolean): Command {
  return cmds.setIndentFirstLine(twips, hanging);
}

// Lists
export const toggleBulletList: Command = cmds.toggleBulletList();
export const toggleNumberedList: Command = cmds.toggleNumberedList();
export const increaseListLevel: Command = cmds.increaseListLevel();
export const decreaseListLevel: Command = cmds.decreaseListLevel();
export const removeList: Command = cmds.removeList();

// Spacing
export function setSpaceBefore(twips: number): Command {
  return cmds.setSpaceBefore(twips);
}
export function setSpaceAfter(twips: number): Command {
  return cmds.setSpaceAfter(twips);
}

// Paragraph styles
export function applyStyle(styleId: string, resolvedAttrs?: ResolvedStyleAttrs): Command {
  return cmds.applyStyle(styleId, resolvedAttrs);
}
export const clearStyle: Command = cmds.clearStyle();

// Section breaks
export function insertSectionBreak(
  breakType: 'nextPage' | 'continuous' | 'oddPage' | 'evenPage'
): Command {
  return cmds.insertSectionBreak(breakType);
}
export const removeSectionBreak: Command = cmds.removeSectionBreak();

// Tab stops
export function addTabStop(
  position: number,
  alignment: TabStopAlignment = 'left',
  leader: TabLeader = 'none'
): Command {
  return cmds.addTabStop(position, alignment, leader);
}
export function removeTabStop(position: number): Command {
  return cmds.removeTabStop(position);
}

// Text direction
export const setRtl: Command = cmds.setRtl();
export const setLtr: Command = cmds.setLtr();

// Table of Contents
export const generateTOC: Command = cmds.generateTOC();

/**
 * Options-aware Table of Contents command. Scope the heading levels, set a
 * custom title (or omit it with `null`/`""`), and toggle hyperlinked entries.
 * Omitting `options` is equivalent to the {@link generateTOC} default.
 *
 * @public
 */
export function generateTableOfContents(options?: GenerateTOCOptions): Command {
  return cmds.generateTOC(options);
}
