import { expect, type Locator, type Page } from '@playwright/test';

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Lightning required marker next to the label (asterisk).
 */
export async function expectModalFieldMarkedRequired(
  modal: Locator,
  fieldLabel: string,
): Promise<void> {
  const trimmed = fieldLabel.trim();
  const nameRe = new RegExp(escapeRegExp(trimmed), 'i');
  const candidates = modal.locator('.slds-form-element');
  const n = await candidates.count();
  for (let i = 0; i < n; i++) {
    const row = candidates.nth(i);
    const labelLoc = row
      .locator('.slds-form-element__label, .slds-form-element__legend, legend, label')
      .first();
    if (!(await labelLoc.isVisible().catch(() => false))) continue;
    const txt = await labelLoc.innerText();
    if (!nameRe.test(txt)) continue;
    await expect(
      row.locator('abbr.slds-required, abbr[title="required"], .slds-required'),
      `Required marker (asterisk) missing for: ${trimmed}`,
    ).toBeVisible({ timeout: 10_000 });
    return;
  }
  throw new Error(
    `Could not find Lead modal field row matching workbook label: "${trimmed}"`,
  );
}

/**
 * Opens a picklist/combobox and selects the first real option (skips "--None--").
 */
export async function selectFirstNonEmptyPicklistOption(
  page: Page,
  modal: Locator,
  fieldNamePattern: RegExp,
): Promise<void> {
  const combo = modal.getByRole('combobox', { name: fieldNamePattern });
  await expect(combo).toBeVisible({ timeout: 20_000 });
  await combo.click();

  const listbox = page
    .locator('[role="listbox"]')
    .filter({ has: page.locator('[role="option"]') })
    .last();
  await expect(listbox).toBeVisible({ timeout: 15_000 });

  const options = listbox.getByRole('option');
  const count = await options.count();
  for (let i = 0; i < count; i++) {
    const opt = options.nth(i);
    const t = (await opt.innerText()).replace(/\u00a0/g, ' ').trim();
    if (!t) continue;
    if (/^(--\s*)?none\s*(--)?$/i.test(t)) continue;
    await opt.click();
    return;
  }

  throw new Error(
    `No selectable option found for picklist matching ${fieldNamePattern}`,
  );
}
