import { test, expect } from '@playwright/test';
import { loadSalesforceCredentials } from '../lib/salesforceCredentials';
import { logQaMilestone } from '../lib/qaSalesforceStatus';
import {
  loadLeadFieldConfigRowsFromTibbiyahWorkbook,
  resolveTibbiyahLeadConfigPathAsync,
} from '../lib/tibbiyahLeadConfig';
import { reportLeadModalExcelUiComparison } from '../lib/leadModalExcelUiComparison';
import {
  reportBusinessUnitAndPortfolioPicklistValidation,
  reportProcurementClassificationPicklistValidation,
} from '../lib/leadModalPicklistComparison';

function tibbiyahSheetSelector(): number | string {
  const raw = process.env.TIBBIYAH_LEAD_CONFIG_SHEET?.trim();
  if (!raw) return 0;
  if (/^\d+$/.test(raw)) return Number.parseInt(raw, 10);
  return raw;
}

test.describe('QA Salesforce', () => {
  test('opens Chromium, logs in, and opens New Lead from Leads', async ({
    page,
  }) => {
    /** Steps 5–7 touch many picklists; allow time so the browser is not torn down mid–Step 7. */
    test.setTimeout(300_000);

    const { url, username, password } = await loadSalesforceCredentials();

    await page.context().grantPermissions(['geolocation'], {
      origin: new URL(url).origin,
    });

    await page.goto(url);

    await page.locator('#username').fill(username);
    await page.locator('#password').fill(password);
    await page.locator('#Login').click();

    await expect(page.locator('#username')).toBeHidden({ timeout: 60_000 });
    logQaMilestone({ index: 1, total: 7, name: 'Step 1: Login' });

    try {
      await page
        .getByRole('button', { name: 'Log In to Sandbox' })
        .click({ timeout: 10_000 });
    } catch {
      /* optional sandbox interstitial */
    }

    logQaMilestone({ index: 2, total: 7, name: 'Step 2: Home' });

    await page.getByRole('link', { name: 'Leads' }).click({ timeout: 90_000 });

    await expect(page).toHaveURL(/\/Lead\//i, { timeout: 60_000 });
    logQaMilestone({ index: 3, total: 7, name: 'Step 3: Lead list view' });

    await page.context().grantPermissions(['geolocation'], {
      origin: new URL(page.url()).origin,
    });

    const newLeadButton = page.getByRole('button', { name: 'New' }).first();
    await expect(newLeadButton).toBeVisible({ timeout: 60_000 });
    await newLeadButton.click();

    await page.waitForTimeout(6_500);

    const leadModal = page.getByRole('dialog', { name: /new lead/i });
    await expect(leadModal).toBeVisible({ timeout: 30_000 });
    logQaMilestone({ index: 4, total: 7, name: 'Step 4: Lead creation modal' });

    let workbookReady = false;
    let completedStep5 = false;
    let completedStep6 = false;
    let completedStep7 = false;
    try {
      const tibbiyahPath = await resolveTibbiyahLeadConfigPathAsync();
      const sheetSel = tibbiyahSheetSelector();
      const fieldRows = await loadLeadFieldConfigRowsFromTibbiyahWorkbook(
        tibbiyahPath,
        sheetSel,
      );
      workbookReady = true;

      await reportLeadModalExcelUiComparison(page, leadModal, fieldRows);
      completedStep5 = true;
      logQaMilestone({
        index: 5,
        total: 7,
        name: 'Step 5: Validate fields (Excel vs UI)',
      });

      await reportBusinessUnitAndPortfolioPicklistValidation(
        page,
        leadModal,
        tibbiyahPath,
        sheetSel,
      );
      completedStep6 = true;
      logQaMilestone({
        index: 6,
        total: 7,
        name: 'Step 6: Validate Business Unit & Portfolio picklists',
      });

      await reportProcurementClassificationPicklistValidation(
        page,
        leadModal,
        tibbiyahPath,
        sheetSel,
      );
      completedStep7 = true;
      logQaMilestone({
        index: 7,
        total: 7,
        name: 'Step 7: Validate Procurement Classification picklists',
      });
    } catch (err) {
      const msg = (err as Error).message;
      console.log(
        `[QA Salesforce] Steps 5–7 halted (each Playwright run starts fresh; milestones below reflect how far this run got): ${msg}`,
      );
      if (!workbookReady) {
        logQaMilestone({
          index: 5,
          total: 7,
          name: 'Step 5: Validate fields (Excel vs UI)',
          outcome: 'skipped',
        });
        logQaMilestone({
          index: 6,
          total: 7,
          name: 'Step 6: Validate Business Unit & Portfolio picklists',
          outcome: 'skipped',
        });
        logQaMilestone({
          index: 7,
          total: 7,
          name: 'Step 7: Validate Procurement Classification picklists',
          outcome: 'skipped',
        });
      } else if (!completedStep5) {
        logQaMilestone({
          index: 5,
          total: 7,
          name: 'Step 5: Validate fields (Excel vs UI)',
          outcome: 'failed',
        });
        logQaMilestone({
          index: 6,
          total: 7,
          name: 'Step 6: Validate Business Unit & Portfolio picklists',
          outcome: 'skipped',
        });
        logQaMilestone({
          index: 7,
          total: 7,
          name: 'Step 7: Validate Procurement Classification picklists',
          outcome: 'skipped',
        });
      } else if (!completedStep6) {
        logQaMilestone({
          index: 6,
          total: 7,
          name: 'Step 6: Validate Business Unit & Portfolio picklists',
          outcome: 'failed',
        });
        logQaMilestone({
          index: 7,
          total: 7,
          name: 'Step 7: Validate Procurement Classification picklists',
          outcome: 'skipped',
        });
      } else if (!completedStep7) {
        logQaMilestone({
          index: 7,
          total: 7,
          name: 'Step 7: Validate Procurement Classification picklists',
          outcome: 'failed',
        });
      }
      throw err;
    }
  });
});
