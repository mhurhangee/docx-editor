import { test, expect } from '@playwright/test';
import { EditorPage } from '../helpers/editor-page';

// Step 3 of footnote-editing unification: a click inside a painted footnote
// (`.layout-footnote-content[data-footnote-id]`) enters footnote-edit mode for
// THAT footnote, places the caret at the clicked char in its hidden view, and
// focuses it — so typing routes to the correct footnote and the painted
// footnote text updates live. Clicking body text afterward exits footnote mode
// and restores body editing.
test.describe('Footnote editing (click → focus → type)', () => {
  test('typing in a clicked footnote updates that footnote, then body click exits', async ({
    page,
  }) => {
    const editor = new EditorPage(page);
    await editor.goto();
    await editor.waitForReady();

    await page
      .locator('input[type="file"][accept=".docx"]')
      .setInputFiles('e2e/fixtures/footnote-bottom-overflow.docx');

    // Footnotes paint into the footnote area at the bottom of the page.
    await page.waitForSelector('.layout-footnote-area', { timeout: 15000 });
    const footnote = page.locator('.layout-footnote-content[data-footnote-id]').first();
    await expect(footnote).toBeVisible();

    // Capture the clicked footnote's id + its text before editing.
    const fnId = await footnote.getAttribute('data-footnote-id');
    expect(fnId).not.toBeNull();
    const before = (await footnote.textContent()) ?? '';

    // Click a painted run in the footnote → enter footnote-edit mode + caret.
    await footnote.locator('span[data-pm-start]').first().click();

    // Type a sentinel; it must land in THIS footnote (identified by id).
    const sentinel = 'ZZQ';
    await page.keyboard.type(sentinel);

    // The painted footnote with this id now contains the typed text.
    const sameFootnote = page.locator(`.layout-footnote-content[data-footnote-id="${fnId}"]`);
    await expect(sameFootnote).toContainText(sentinel, { timeout: 5000 });
    const after = (await sameFootnote.textContent()) ?? '';
    expect(after).not.toBe(before);

    // Click body text → exit footnote mode; the body becomes editable again.
    const bodySpan = page.locator('.layout-page-content span[data-pm-start]').first();
    await bodySpan.click();

    // After exiting, typing into the body lands in body content, not the
    // footnote. Type a distinct sentinel and assert the footnote is unchanged
    // while body content grew.
    const bodyBefore = (await page.locator('.layout-page-content').first().textContent()) ?? '';
    const footnoteAfterExit = (await sameFootnote.textContent()) ?? '';
    await page.keyboard.type('YYW');
    await expect(page.locator('.layout-page-content').first()).toContainText('YYW', {
      timeout: 5000,
    });
    // Footnote text did not change after the body edit → footnote mode exited.
    expect((await sameFootnote.textContent()) ?? '').toBe(footnoteAfterExit);
    expect(bodyBefore).not.toContain('YYW');
  });

  // Step 4: clicking into a footnote must draw a VISIBLE caret over the painted
  // footnote (the hidden view holds the selection; the painter is the sole
  // visible renderer). The caret is a thin (width:2) blinking div portalled near
  // the painter container — black like the body caret (not the HF blue). Its
  // bounding box must lie inside the clicked footnote's `.layout-footnote-content`.
  test('a visible caret renders within the clicked footnote bounds', async ({ page }) => {
    const editor = new EditorPage(page);
    await editor.goto();
    await editor.waitForReady();

    await page
      .locator('input[type="file"][accept=".docx"]')
      .setInputFiles('e2e/fixtures/footnote-bottom-overflow.docx');

    await page.waitForSelector('.layout-footnote-area', { timeout: 15000 });
    const footnote = page.locator('.layout-footnote-content[data-footnote-id]').first();
    await expect(footnote).toBeVisible();
    const fnId = await footnote.getAttribute('data-footnote-id');
    expect(fnId).not.toBeNull();

    // Click a painted run in the footnote → enter footnote-edit mode + caret.
    await footnote.locator('span[data-pm-start]').first().click();

    // The blinking caret is painted asynchronously (rAF / `painter:painted`
    // after the selection transaction), so poll until the thin caret div appears
    // and report whether its center sits inside the clicked footnote. Identify it
    // color-agnostically by its thin width + blink animation (caret is black now).
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

  // Click-drag across footnote text must select a range: the entry arms drag
  // (dragAnchorRef + isDraggingRef), and the move handler extends the selection
  // through the footnote surface. A non-empty range paints selection-rect divs.
  test('click-drag selects a range within the footnote', async ({ page }) => {
    const editor = new EditorPage(page);
    await editor.goto();
    await editor.waitForReady();

    await page
      .locator('input[type="file"][accept=".docx"]')
      .setInputFiles('e2e/fixtures/footnote-bottom-overflow.docx');

    await page.waitForSelector('.layout-footnote-area', { timeout: 15000 });
    const fn = page.locator('.layout-footnote-content[data-footnote-id]').first();
    await expect(fn).toBeVisible();
    const span = fn.locator('span[data-pm-start]').first();
    await span.click(); // enter footnote-edit mode (caret), so the view is mounted
    await page.waitForTimeout(150);

    // Drag across the footnote's first line (wide x-range for a multi-char range).
    const fb = await fn.boundingBox();
    const sb = await span.boundingBox();
    if (!fb || !sb) throw new Error('no footnote box');
    const y = sb.y + sb.height / 2;
    await page.mouse.move(fb.x + 4, y);
    await page.mouse.down();
    await page.mouse.move(fb.x + fb.width * 0.6, y, { steps: 10 });
    await page.mouse.up();
    await page.waitForTimeout(100);

    // A selection range paints translucent-blue selection-rect divs (width > 3
    // distinguishes them from the thin caret). Assert at least one appeared.
    const hasSelectionRect = await page.evaluate(() => {
      return Array.from(document.querySelectorAll<HTMLElement>('div[aria-hidden="true"]')).some(
        (el) => {
          const w = parseFloat(el.style.width || '0');
          return /66, 133, 244/i.test(el.style.background || '') && w > 3;
        }
      );
    });
    expect(hasSelectionRect).toBe(true);
  });
});
