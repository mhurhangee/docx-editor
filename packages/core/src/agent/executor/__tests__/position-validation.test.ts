import { describe, test, expect } from 'bun:test';
import { validatePosition, validateRange } from '../helpers';
import { executeInsertText, executeReplaceText } from '../textCommands';
import { createEmptyDocument } from '../../../utils/createDocument';
import type { InsertTextCommand, ReplaceTextCommand } from '../../../types/agentApi';

describe('validatePosition', () => {
  test('accepts non-negative integers', () => {
    expect(() => validatePosition({ paragraphIndex: 0, offset: 0 })).not.toThrow();
    expect(() => validatePosition({ paragraphIndex: 3, offset: 12 })).not.toThrow();
  });

  test('rejects negative, fractional, or non-finite values', () => {
    expect(() => validatePosition({ paragraphIndex: -1, offset: 0 })).toThrow(/paragraphIndex/);
    expect(() => validatePosition({ paragraphIndex: 0, offset: -5 })).toThrow(/offset/);
    expect(() => validatePosition({ paragraphIndex: 1.5, offset: 0 })).toThrow(/paragraphIndex/);
    expect(() => validatePosition({ paragraphIndex: 0, offset: NaN })).toThrow(/offset/);
    expect(() => validatePosition({ paragraphIndex: 0, offset: Infinity })).toThrow(/offset/);
  });
});

describe('validateRange', () => {
  test('accepts an ordered range', () => {
    expect(() =>
      validateRange({
        start: { paragraphIndex: 0, offset: 0 },
        end: { paragraphIndex: 0, offset: 4 },
      })
    ).not.toThrow();
  });

  test('rejects an inverted range', () => {
    expect(() =>
      validateRange({
        start: { paragraphIndex: 0, offset: 6 },
        end: { paragraphIndex: 0, offset: 2 },
      })
    ).toThrow(/end precedes start/);
    expect(() =>
      validateRange({
        start: { paragraphIndex: 2, offset: 0 },
        end: { paragraphIndex: 1, offset: 0 },
      })
    ).toThrow(/end precedes start/);
  });
});

describe('executor rejects invalid input', () => {
  test('executeInsertText throws on a negative offset', () => {
    const doc = createEmptyDocument();
    const cmd: InsertTextCommand = {
      type: 'insertText',
      position: { paragraphIndex: 0, offset: -1 },
      text: 'x',
    };
    expect(() => executeInsertText(doc, cmd)).toThrow(/offset/);
  });

  test('executeReplaceText throws on an inverted range', () => {
    const doc = createEmptyDocument();
    const cmd: ReplaceTextCommand = {
      type: 'replaceText',
      range: { start: { paragraphIndex: 0, offset: 5 }, end: { paragraphIndex: 0, offset: 1 } },
      text: 'x',
    };
    expect(() => executeReplaceText(doc, cmd)).toThrow(/end precedes start/);
  });
});
