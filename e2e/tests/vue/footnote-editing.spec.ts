import { test, expect } from '@playwright/test';

// Vue mirror of e2e/tests/footnote-editing.spec.ts (React). Core-parity subset:
// clicking a painted footnote (`.layout-footnote-content[data-footnote-id]`)
// enters footnote-edit mode for THAT footnote, places the caret in its hidden
// view, and routes typing into it so the painted footnote updates live; a
// visible caret renders within the clicked footnote bounds.
test.describe('Vue: Footnote editing (click → focus → type)', () => {
  test('typing in a clicked footnote updates that footnote', async ({ page }) => {
    await page.goto('http://localhost:5174/?e2e=1');
    await page.locator('.docx-editor-vue').waitFor();
    await page.locator('.paged-editor__pages').waitFor();

    await page
      .locator('input[type="file"]')
      .first()
      .setInputFiles('e2e/fixtures/footnote-bottom-overflow.docx');

    // Footnotes paint into the footnote area at the bottom of the page.
    await page.waitForSelector('.layout-footnote-area', { timeout: 15000 });
    const footnote = page.locator('.layout-footnote-content[data-footnote-id]').first();
    await expect(footnote).toBeVisible();

    const fnId = await footnote.getAttribute('data-footnote-id');
    expect(fnId).not.toBeNull();
    const before = (await footnote.textContent()) ?? '';

    // Click a painted run in the footnote → enter footnote-edit mode + caret.
    await footnote.locator('span[data-pm-start]').first().click();

    // Type a sentinel; it must land in THIS footnote (identified by id).
    const sentinel = 'ZZQ';
    await page.keyboard.type(sentinel);

    const sameFootnote = page.locator(`.layout-footnote-content[data-footnote-id="${fnId}"]`);
    await expect(sameFootnote).toContainText(sentinel, { timeout: 5000 });
    const after = (await sameFootnote.textContent()) ?? '';
    expect(after).not.toBe(before);
  });

  test('a visible caret renders within the clicked footnote bounds', async ({ page }) => {
    await page.goto('http://localhost:5174/?e2e=1');
    await page.locator('.docx-editor-vue').waitFor();
    await page.locator('.paged-editor__pages').waitFor();

    await page
      .locator('input[type="file"]')
      .first()
      .setInputFiles('e2e/fixtures/footnote-bottom-overflow.docx');

    await page.waitForSelector('.layout-footnote-area', { timeout: 15000 });
    const footnote = page.locator('.layout-footnote-content[data-footnote-id]').first();
    await expect(footnote).toBeVisible();
    const fnId = await footnote.getAttribute('data-footnote-id');
    expect(fnId).not.toBeNull();

    // Click a painted run in the footnote → enter footnote-edit mode + caret.
    await footnote.locator('span[data-pm-start]').first().click();

    // The blinking caret is painted asynchronously (rAF / `painter:painted`
    // after the selection transaction), so poll until the thin blue div appears
    // and report whether its center sits inside the clicked footnote.
    const placement = await page
      .waitForFunction(
        (id: string) => {
          const caret = Array.from(
            document.querySelectorAll<HTMLElement>('div[aria-hidden="true"]')
          ).find((el) => {
            const w = parseFloat(el.style.width || '0');
            return w > 0 && w <= 3 && /caret-blink/i.test(el.style.animation || '');
          });
          if (!caret) return null;
          const r = caret.getBoundingClientRect();
          const cx = r.left + r.width / 2;
          const cy = r.top + r.height / 2;
          const fn = document.querySelector<HTMLElement>(
            `.layout-footnote-content[data-footnote-id="${id}"]`
          );
          if (!fn) return null;
          const b = fn.getBoundingClientRect();
          return {
            inFootnote:
              cy >= b.top - 2 && cy <= b.bottom + 2 && cx >= b.left - 2 && cx <= b.right + 2,
          };
        },
        fnId as string,
        { timeout: 5000 }
      )
      .then((handle) => handle.jsonValue());

    expect(placement.inFootnote).toBe(true);
  });
});
