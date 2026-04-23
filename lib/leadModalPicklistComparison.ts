import type { Locator, Page } from '@playwright/test';
import {
  loadBuPortfolioDependencyRowsFromTibbiyahWorkbook,
  loadPicklistDefinitionsFromTibbiyahWorkbook,
  loadProcurementDependencyRowsFromTibbiyahWorkbook,
  splitTibbiyahPicklistCellTokens,
  type BuPortfolioExcelPair,
  type ProcurementTripleExcelRow,
  type TibbiyahPicklistDefinition,
} from './tibbiyahLeadConfig';
import {
  businessUnitValueForPortfolioValidation,
  comboboxAccessibleNamePattern,
  dismissOpenPicklistListboxIfVisible,
  getLeadModalPicklistComboboxInteractionState,
  normPicklistMatchKey,
  readPicklistComboboxDisplayedValueAny,
  picklistComboVisibilityTimeoutMs,
  picklistListboxVisibilityTimeoutMs,
  picklistValuesMatchLoose,
  picklistOptionDisplayText,
  resolveComboboxForPicklistField,
  resolvePicklistListboxAfterComboOpened,
  tryExpandLeadModalSectionByHeading,
  trySelectPicklistOptionByLabel,
  waitForLightningStableState,
} from './leadModalExcelUiComparison';
import {
  buildPicklistModalStructuredBugs,
  buildProcurementClassificationStructuredBugs,
  buildStep6CombinedStructuredBugs,
  buildStep6PicklistStructuredBugs,
  logStructuredQaBugs,
} from './qaBugReport';

export type PicklistComparisonTableRow = {
  'Picklist field': string;
  /** Portfolio picklist options depend on Business Unit. */
  'Depends on': string;
  'Excel values (#)': string;
  'On UI': string;
  'UI options (#)': string;
  'Missing on UI': string;
  Status: string;
};

function dependsOnForPicklistField(fieldName: string): string {
  if (/portfolio/i.test(fieldName.trim())) return 'Business Unit';
  return '—';
}

export type ReadPicklistFromModalOptions = {
  /** Override default listbox visibility wait (slow-loading picklists). */
  listboxTimeoutMs?: number;
  /**
   * Resolve the combobox via the same `.slds-form-element` as the visible label (helps Request Type
   * when `getByRole('combobox', { name })` does not match Lightning’s accessible name).
   */
  preferFormElementLabel?: boolean;
};

/** Opens picklist for `fieldLabel`, reads options, dismisses without Escape. */
export async function tryReadPicklistOptionsFromModal(
  page: Page,
  modal: Locator,
  fieldLabel: string,
  readOpts?: ReadPicklistFromModalOptions,
): Promise<{
  ok: boolean;
  fieldOnUi: boolean;
  options: string[];
  error?: string;
}> {
  const pattern = comboboxAccessibleNamePattern(fieldLabel);
  const listboxTimeout =
    readOpts?.listboxTimeoutMs ?? picklistListboxVisibilityTimeoutMs();
  try {
    if (page.isClosed()) {
      return {
        ok: false,
        fieldOnUi: false,
        options: [],
        error: 'Page, context, or browser has been closed',
      };
    }
    await dismissOpenPicklistListboxIfVisible(page, modal);
    const combo = await resolveComboboxForPicklistField(modal, fieldLabel, pattern, {
      preferFormRow: Boolean(readOpts?.preferFormElementLabel),
    });
    await combo.scrollIntoViewIfNeeded().catch(() => {});
    const fieldOnUi = await combo.isVisible().catch(() => false);
    if (!fieldOnUi) {
      return {
        ok: false,
        fieldOnUi: false,
        options: [],
        error: 'Picklist field not found on modal',
      };
    }
    await combo.waitFor({
      state: 'visible',
      timeout: picklistComboVisibilityTimeoutMs(),
    });
    await combo.click({ force: true });
    if (page.isClosed()) {
      return {
        ok: false,
        fieldOnUi: true,
        options: [],
        error: 'Page, context, or browser has been closed',
      };
    }
    await waitForLightningStableState(page, { modal });
    const listbox = await resolvePicklistListboxAfterComboOpened(page, combo);
    await listbox.waitFor({
      state: 'visible',
      timeout: listboxTimeout,
    });
    const optionLoc = listbox.getByRole('option');
    const count = await optionLoc.count();
    const options: string[] = [];
    for (let i = 0; i < count; i++) {
      const el = optionLoc.nth(i);
      const raw = await el
        .evaluate((node) => {
          const e = node as HTMLElement;
          return (
            e.innerText ||
            e.textContent ||
            e.getAttribute('aria-label') ||
            ''
          ).trim();
        })
        .catch(() => '');
      const t = picklistOptionDisplayText(raw);
      if (t) options.push(t);
    }
    await dismissOpenPicklistListboxIfVisible(page, modal);
    return { ok: true, fieldOnUi: true, options };
  } catch (e) {
    const err = e as Error;
    try {
      await dismissOpenPicklistListboxIfVisible(page, modal);
    } catch (cleanErr) {
      console.warn(
        `tryReadPicklistOptionsFromModal: listbox cleanup failed after "${fieldLabel}": ${(cleanErr as Error).message}`,
      );
    }
    return {
      ok: false,
      fieldOnUi: true,
      options: [],
      error: err.message,
    };
  }
}

function uiHasValue(uiOptions: string[], expected: string): boolean {
  if (!expected?.trim()) return true;
  return uiOptions.some((o) => picklistValuesMatchLoose(o, expected));
}

function summarizeMissing(missing: string[], maxLen = 120): string {
  if (missing.length === 0) return '—';
  const s = missing.join('; ');
  return s.length <= maxLen ? s : `${s.slice(0, maxLen - 3)}...`;
}

/** Semicolon-separated Request Type options for Step 7 console table (truncated). */
function summarizeRtPicklistOptionsForTable(opts: string[], maxLen = 450): string {
  if (!opts.length) return '—';
  const s = opts
    .map((o) => o.replace(/\s+/g, ' ').trim())
    .filter(Boolean)
    .join('; ');
  if (!s) return '—';
  return s.length <= maxLen ? s : `${s.slice(0, maxLen - 3)}...`;
}

function filterRealPicklistOptions(opts: string[]): string[] {
  return opts.filter((o) => {
    const t = o.replace(/\s+/g, ' ').trim();
    if (!t) return false;
    if (/^(--\s*)?none\s*(--)?$/i.test(t)) return false;
    return true;
  });
}

function procurementPairKey(sectorUi: string, channelUi: string): string {
  return `${sectorUi}\x1f${channelUi}`;
}

function findMatchingUiOption(
  excelValue: string,
  uiOptions: string[],
): string | null {
  const t = excelValue.replace(/\s+/g, ' ').trim();
  if (!t) return null;
  for (const o of uiOptions) {
    if (picklistValuesMatchLoose(o, t)) return o;
  }
  const ek = normPicklistMatchKey(t);
  for (const o of uiOptions) {
    if (normPicklistMatchKey(o) === ek) return o;
  }
  return null;
}

/** `0` = unlimited. */
function procurementMatrixMaxSectors(): number {
  const raw = process.env.TIBBIYAH_PROCUREMENT_MATRIX_MAX_SECTORS?.trim();
  if (raw === '0' || raw === '') return 0;
  if (raw && /^\d+$/.test(raw)) return Math.max(0, Number.parseInt(raw, 10));
  return 0;
}

/** `0` = unlimited. */
function procurementMatrixMaxChannels(): number {
  const raw = process.env.TIBBIYAH_PROCUREMENT_MATRIX_MAX_CHANNELS?.trim();
  if (raw === '0' || raw === '') return 0;
  if (raw && /^\d+$/.test(raw)) return Math.max(0, Number.parseInt(raw, 10));
  return 0;
}

function procurementMatrixStepDelayMs(): number {
  const raw = process.env.TIBBIYAH_PROCUREMENT_MATRIX_STEP_MS?.trim();
  if (raw && /^\d+$/.test(raw)) return Number.parseInt(raw, 10);
  return 350;
}

export type ProcurementPicklistCaptureMatrix = {
  sectorOptions: string[];
  channelOptionsBySector: Record<string, string[]>;
  requestTypeOptionsByPair: Record<string, string[]>;
  captureLog: string[];
};

/**
 * 1) Open **Procurement Sector** → capture all options.  
 * 2) For **each** sector value: select it → open **Procurement Channel** → capture all options.  
 * 3) For **each** channel value: select it → open **Request Type** → capture all options.  
 * Stores Channel lists per sector and Request Type lists per (sector, channel) pair.
 */
async function captureFullProcurementPicklistMatrixFromModal(
  page: Page,
  dlg: Locator,
  sectorL: string,
  channelL: string,
  rtL: string,
  options?: { restrictSectors?: Set<string> },
): Promise<ProcurementPicklistCaptureMatrix> {
  const captureLog: string[] = [];
  const channelOptionsBySector: Record<string, string[]> = {};
  const requestTypeOptionsByPair: Record<string, string[]> = {};
  const readOpts = {
    preferFormElementLabel: true,
    listboxTimeoutMs: picklistListboxVisibilityTimeoutMs(),
  };
  const stepMs = procurementMatrixStepDelayMs();

  const sectorRead = await tryReadPicklistOptionsFromModal(
    page,
    dlg,
    sectorL,
    readOpts,
  );
  if (!sectorRead.ok) {
    captureLog.push(`Procurement Sector read failed: ${sectorRead.error ?? 'unknown'}`);
    return {
      sectorOptions: [],
      channelOptionsBySector,
      requestTypeOptionsByPair,
      captureLog,
    };
  }

  let sectorOptions = filterRealPicklistOptions(sectorRead.options);
  if (options?.restrictSectors && options.restrictSectors.size > 0) {
    const sectorBag = options.restrictSectors;
    sectorOptions = sectorOptions.filter((s) =>
      [...sectorBag].some((ex) => picklistValuesMatchLoose(s, ex)),
    );
    captureLog.push(
      `Sectors restricted to Excel-only set (${sectorOptions.length} UI sector(s) after filter).`,
    );
  }

  const maxS = procurementMatrixMaxSectors();
  if (maxS > 0 && sectorOptions.length > maxS) {
    captureLog.push(
      `Note: sector walk capped at ${maxS} (set TIBBIYAH_PROCUREMENT_MATRIX_MAX_SECTORS=0 for unlimited).`,
    );
    sectorOptions = sectorOptions.slice(0, maxS);
  }

  captureLog.push(
    `Procurement Sector: captured ${sectorOptions.length} option(s) from UI.`,
  );

  for (const sectorUi of sectorOptions) {
    if (page.isClosed()) break;
    const selS = await trySelectPicklistOptionByLabel(
      page,
      dlg,
      comboboxAccessibleNamePattern(sectorL),
      sectorUi,
      { formElementLabel: sectorL },
    );
    if (!selS.ok) {
      captureLog.push(
        `Could not select Procurement Sector "${sectorUi}": ${selS.error ?? 'unknown'}`,
      );
      continue;
    }
    if (stepMs > 0) await page.waitForTimeout(stepMs);
    const postS = procurementPostSectorDelayMs();
    if (postS > 0) await page.waitForTimeout(postS);

    const chRead = await tryReadPicklistOptionsFromModal(
      page,
      dlg,
      channelL,
      readOpts,
    );
    if (!chRead.ok) {
      captureLog.push(
        `Procurement Channel read failed [Sector="${sectorUi}"]: ${chRead.error ?? 'unknown'}`,
      );
      channelOptionsBySector[sectorUi] = [];
      continue;
    }

    let channelOpts = filterRealPicklistOptions(chRead.options);
    const maxC = procurementMatrixMaxChannels();
    if (maxC > 0 && channelOpts.length > maxC) {
      captureLog.push(
        `Note: channel walk for "${sectorUi}" capped at ${maxC} (set TIBBIYAH_PROCUREMENT_MATRIX_MAX_CHANNELS=0 for unlimited).`,
      );
      channelOpts = channelOpts.slice(0, maxC);
    }
    channelOptionsBySector[sectorUi] = channelOpts;
    captureLog.push(
      `  [Sector="${sectorUi}"] Procurement Channel: ${channelOpts.length} option(s).`,
    );

    for (const channelUi of channelOpts) {
      if (page.isClosed()) break;
      const selC = await trySelectPicklistOptionByLabel(
        page,
        dlg,
        comboboxAccessibleNamePattern(channelL),
        channelUi,
        { formElementLabel: channelL },
      );
      if (!selC.ok) {
        captureLog.push(
          `Could not select Procurement Channel "${channelUi}" [Sector="${sectorUi}"]: ${selC.error ?? 'unknown'}`,
        );
        continue;
      }
      if (stepMs > 0) await page.waitForTimeout(stepMs);
      const postC = procurementPostChannelDelayMs();
      if (postC > 0) await page.waitForTimeout(postC);

      const rtRead = await tryReadPicklistOptionsFromModal(page, dlg, rtL, readOpts);
      const pk = procurementPairKey(sectorUi, channelUi);
      if (rtRead.ok) {
        requestTypeOptionsByPair[pk] = filterRealPicklistOptions(rtRead.options);
      } else {
        requestTypeOptionsByPair[pk] = [];
        captureLog.push(
          `  Request Type read [${sectorUi} / ${channelUi}]: ${rtRead.error ?? 'failed'} (empty list).`,
        );
      }
      captureLog.push(
        `    [${sectorUi} → ${channelUi}] Request Type: ${requestTypeOptionsByPair[pk].length} option(s).`,
      );
    }
  }

  return {
    sectorOptions,
    channelOptionsBySector,
    requestTypeOptionsByPair,
    captureLog,
  };
}

function buildExcelVsMatrixComparisonRows(
  mapRows: ProcurementTripleExcelRow[],
  matrix: ProcurementPicklistCaptureMatrix,
): ProcurementClassificationTableRow[] {
  const results: ProcurementClassificationTableRow[] = [];

  for (const pr of mapRows) {
    const baseRow: Omit<ProcurementClassificationTableRow, 'Status'> = {
      'Excel row': pr.sheetRow,
      'Excel Sector': pr.sector,
      'Excel Channel': pr.channel,
      'Excel Request Type': pr.requestTypeRaw || '(empty)',
      'UI Sector (captured)': '',
      'UI Channel (captured)': '',
      'UI Request Type (captured)': '',
      'Channel enabled': '',
      'Request Type enabled': '',
      'Expected RT enabled': excelImpliesRequestTypeDisabled(pr.requestTypeRaw)
        ? 'No'
        : 'Yes',
      'Missing RT on UI': '—',
      'RT picklist options (captured)': '—',
    };

    const sectorUi = findMatchingUiOption(pr.sector, matrix.sectorOptions);
    const chList = sectorUi
      ? matrix.channelOptionsBySector[sectorUi] ?? []
      : [];
    const channelUi = sectorUi
      ? findMatchingUiOption(pr.channel, chList)
      : null;
    const pk =
      sectorUi && channelUi ? procurementPairKey(sectorUi, channelUi) : '';
    const rtList = pk ? matrix.requestTypeOptionsByPair[pk] ?? [] : [];

    const rtCaptured = summarizeRtPicklistOptionsForTable(rtList, 500);

    let status = 'OK';
    if (matrix.sectorOptions.length === 0) {
      status = 'BUG: No Procurement Sector options captured from UI';
    } else if (!sectorUi) {
      status = `BUG: Excel Sector "${pr.sector}" not found in captured Sector list`;
    } else if (!channelUi) {
      status = `BUG: Excel Channel "${pr.channel}" not found in captured Channel list for Sector "${sectorUi}"`;
    } else if (!excelImpliesRequestTypeDisabled(pr.requestTypeRaw)) {
      const tokens = splitTibbiyahPicklistCellTokens(pr.requestTypeRaw);
      const missing = tokens.filter(
        (t) => t.trim() && !uiHasValue(rtList, t),
      );
      if (missing.length > 0) {
        status = `BUG: ${missing.length} Excel Request Type token(s) not in captured list for this Sector/Channel`;
      }
    } else if (rtList.length > 0) {
      status =
        'BUG: Excel expects Request Type disabled but captured picklist has option(s)';
    }

    const missingRt =
      !excelImpliesRequestTypeDisabled(pr.requestTypeRaw) && sectorUi && channelUi
        ? summarizeMissing(
            splitTibbiyahPicklistCellTokens(pr.requestTypeRaw).filter(
              (t) => t.trim() && !uiHasValue(rtList, t),
            ),
          )
        : '—';

    results.push({
      ...baseRow,
      'UI Sector (captured)': sectorUi ?? '—',
      'UI Channel (captured)': channelUi ?? '—',
      'UI Request Type (captured)': '—',
      'Channel enabled': sectorUi ? 'Yes' : '—',
      'Request Type enabled':
        !sectorUi || !channelUi ? '—' : rtList.length > 0 ? 'Yes' : 'No',
      'Missing RT on UI': missingRt,
      'RT picklist options (captured)': rtCaptured,
      Status: status,
    });
  }

  return results;
}

/** Excel Picklist Values vs Lead modal comboboxes. Does not fail the test. */
export async function reportPicklistExcelVsLeadModal(
  page: Page,
  modal: Locator,
  definitions: TibbiyahPicklistDefinition[],
): Promise<void> {
  console.log('\nPicklist: Excel vs modal\n');

  if (definitions.length === 0) {
    console.log('(No picklist definitions in workbook.)\n');
    return;
  }

  const table: PicklistComparisonTableRow[] = [];

  for (const def of definitions) {
    const read = await tryReadPicklistOptionsFromModal(page, modal, def.fieldName);
    const excelCount = def.expectedValues.length;
    const uiCount = read.options.length;

    if (!read.fieldOnUi) {
      table.push({
        'Picklist field': def.fieldName,
        'Depends on': dependsOnForPicklistField(def.fieldName),
        'Excel values (#)': String(excelCount),
        'On UI': 'No',
        'UI options (#)': '—',
        'Missing on UI': summarizeMissing(def.expectedValues),
        Status: 'BUG: Picklist field not on modal',
      });
      continue;
    }

    if (!read.ok) {
      table.push({
        'Picklist field': def.fieldName,
        'Depends on': dependsOnForPicklistField(def.fieldName),
        'Excel values (#)': String(excelCount),
        'On UI': 'Yes',
        'UI options (#)': String(uiCount),
        'Missing on UI': '—',
        Status: `BUG: Could not read options (${read.error ?? 'unknown'})`,
      });
      continue;
    }

    const missing = def.expectedValues.filter(
      (ev) => ev.trim() && !uiHasValue(read.options, ev),
    );

    let status = 'OK';
    if (excelCount === 0) {
      status = 'OK (no values listed in Excel for this field)';
    } else if (missing.length > 0) {
      status = `BUG: ${missing.length} Excel value(s) not found on UI`;
    }

    table.push({
      'Picklist field': def.fieldName,
      'Depends on': dependsOnForPicklistField(def.fieldName),
      'Excel values (#)': String(excelCount),
      'On UI': 'Yes',
      'UI options (#)': String(uiCount),
      'Missing on UI': summarizeMissing(missing),
      Status: status,
    });
  }

  console.table(table);

  const bugs = table.filter((r) => r.Status.startsWith('BUG'));
  if (bugs.length > 0) {
    console.log(`Picklist: ${bugs.length} BUG row(s) (see table).\n`);
    logStructuredQaBugs(
      'Picklist — structured bug report(s)',
      buildPicklistModalStructuredBugs(table),
    );
  } else {
    console.log('Picklist: no BUG rows.\n');
  }
}

function postBuDelayBeforePortfolioReadMs(): number {
  const raw = process.env.TIBBIYAH_BU_WALKTHROUGH_POST_BU_MS?.trim();
  if (raw && /^\d+$/.test(raw)) return Number.parseInt(raw, 10);
  return 1_200;
}

/** Wait after Business Unit change before Portfolio actions (Step 6 only; keeps runs fast). */
function step6PostBuDelayMs(): number {
  const raw = process.env.TIBBIYAH_STEP6_POST_BU_MS?.trim();
  if (raw && /^\d+$/.test(raw)) return Number.parseInt(raw, 10);
  return 350;
}

/** Pause between successive Portfolio selections on the same BU (0 = fastest). */
function step6InterPortfolioSelectDelayMs(): number {
  const raw = process.env.TIBBIYAH_STEP6_INTER_PORTFOLIO_MS?.trim();
  if (raw && /^\d+$/.test(raw)) return Number.parseInt(raw, 10);
  return 25;
}

function findBusinessUnitAndPortfolioDefinitions(
  definitions: TibbiyahPicklistDefinition[],
): { buDef?: TibbiyahPicklistDefinition; pfDef?: TibbiyahPicklistDefinition } {
  const strip = (s: string) => s.replace(/\*/g, '').trim();
  const buDef = definitions.find((d) => /business\s*unit/i.test(strip(d.fieldName)));
  const pfDef = definitions.find((d) => /^portfolio\b/i.test(strip(d.fieldName)));
  return { buDef, pfDef };
}

function picklistTableRow(
  def: TibbiyahPicklistDefinition,
  read: Awaited<ReturnType<typeof tryReadPicklistOptionsFromModal>>,
  dependsOnOverride?: string,
): PicklistComparisonTableRow {
  const excelCount = def.expectedValues.length;
  const uiCount = read.options.length;
  const depends =
    dependsOnOverride ?? dependsOnForPicklistField(def.fieldName);

  if (!read.fieldOnUi) {
    return {
      'Picklist field': def.fieldName,
      'Depends on': depends,
      'Excel values (#)': String(excelCount),
      'On UI': 'No',
      'UI options (#)': '—',
      'Missing on UI': summarizeMissing(def.expectedValues),
      Status: 'BUG: Picklist field not on modal',
    };
  }

  if (!read.ok) {
    return {
      'Picklist field': def.fieldName,
      'Depends on': depends,
      'Excel values (#)': String(excelCount),
      'On UI': 'Yes',
      'UI options (#)': String(uiCount),
      'Missing on UI': '—',
      Status: `BUG: Could not read options (${read.error ?? 'unknown'})`,
    };
  }

  const missing = def.expectedValues.filter(
    (ev) => ev.trim() && !uiHasValue(read.options, ev),
  );

  let status = 'OK';
  if (excelCount === 0) {
    status = 'OK (no values listed in Excel for this field)';
  } else if (missing.length > 0) {
    status = `BUG: ${missing.length} Excel value(s) not found on UI`;
  }

  return {
    'Picklist field': def.fieldName,
    'Depends on': depends,
    'Excel values (#)': String(excelCount),
    'On UI': 'Yes',
    'UI options (#)': String(uiCount),
    'Missing on UI': summarizeMissing(missing),
    Status: status,
  };
}

export type BuPortfolioDependencyTableRow = {
  'Business Unit': string;
  'Excel Portfolio(s)': string;
  'Missing on UI': string;
  Status: string;
};

const BU_PORTFOLIO_DEP_COLUMNS = [
  'Business Unit',
  'Excel Portfolio(s)',
  'Missing on UI',
  'Status',
] as const satisfies ReadonlyArray<keyof BuPortfolioDependencyTableRow>;

function logBuPortfolioDependencyTable(rows: BuPortfolioDependencyTableRow[]): void {
  if (rows.length === 0) {
    console.log('(no rows)\n');
    return;
  }
  const widths = BU_PORTFOLIO_DEP_COLUMNS.map((col) =>
    Math.max(
      col.length,
      ...rows.map((r) => String(r[col]).length),
    ),
  );
  const pad = (s: string, w: number) => s.padEnd(w, ' ');
  const line = (cells: string[]) =>
    cells.map((c, i) => pad(c, widths[i])).join('  ');
  console.log(line([...BU_PORTFOLIO_DEP_COLUMNS]));
  console.log(
    line(BU_PORTFOLIO_DEP_COLUMNS.map((_, i) => '-'.repeat(widths[i]))),
  );
  for (const r of rows) {
    console.log(
      line(BU_PORTFOLIO_DEP_COLUMNS.map((c) => String(r[c]))),
    );
  }
  console.log('');
}

function portfolioFieldLabelForModal(): string {
  return process.env.TIBBIYAH_PORTFOLIO_FIELD_LABEL?.trim() || 'Portfolio';
}

function businessUnitFieldLabelForModal(): string {
  return process.env.TIBBIYAH_BUSINESS_UNIT_FIELD_LABEL?.trim() || 'Business Unit';
}

function dedupeBusinessUnitsPreserveOrder(units: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const u of units) {
    const t = u.replace(/\s+/g, ' ').trim();
    if (!t) continue;
    const k = normPicklistMatchKey(t);
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(t);
  }
  return out;
}

/** Step 6 when no B+C BU/Portfolio map: Picklist Values for Business Unit + Portfolio. */
async function reportStep6BuPortfolioPicklistFallback(
  page: Page,
  modal: Locator,
  workbookPath: string,
  sheetIndexOrName: number | string,
): Promise<void> {
  const definitions = await loadPicklistDefinitionsFromTibbiyahWorkbook(
    workbookPath,
    sheetIndexOrName,
  );
  const { buDef, pfDef } = findBusinessUnitAndPortfolioDefinitions(definitions);
  const table: PicklistComparisonTableRow[] = [];

  if (!buDef || !pfDef) {
    console.log('\nStep 6: no B/C map and no Picklist Values for BU+Portfolio.\n');
    const missing = [!buDef && 'Business Unit', !pfDef && 'Portfolio']
      .filter(Boolean)
      .join('; ');
    table.push({
      'Picklist field': missing || 'Business Unit; Portfolio',
      'Depends on': '—',
      'Excel values (#)': '—',
      'On UI': '—',
      'UI options (#)': '—',
      'Missing on UI': '—',
      Status: 'BUG: Missing Excel BU/Portfolio source',
    });
    console.table(table);
    logStructuredQaBugs(
      'Step 6 — structured bug report(s)',
      buildStep6PicklistStructuredBugs(table),
    );
    return;
  }

  console.log('\nStep 6: BU & Portfolio (Picklist Values fallback).\n');

  const buRead = await tryReadPicklistOptionsFromModal(
    page,
    modal,
    buDef.fieldName,
  );
  table.push(picklistTableRow(buDef, buRead));

  const envBu = businessUnitValueForPortfolioValidation();
  const tryOrder = [envBu, ...buDef.expectedValues];
  let selectedBu = '';
  if (buRead.ok && buRead.options.length > 0) {
    for (const c of tryOrder) {
      const t = c?.trim();
      if (!t) continue;
      if (uiHasValue(buRead.options, t)) {
        selectedBu =
          buRead.options.find((o) => picklistValuesMatchLoose(o, t)) ?? t;
        break;
      }
    }
    if (!selectedBu) {
      const firstReal = buRead.options.find(
        (o) => !/^(--\s*)?none\s*(--)?$/i.test(o.trim()),
      );
      if (firstReal) selectedBu = firstReal.trim();
    }
  }

  let portfolioDepends = 'Business Unit (unchanged)';
  if (selectedBu) {
    const sel = await trySelectPicklistOptionByLabel(
      page,
      modal,
      /Business Unit/i,
      selectedBu,
    );
    if (sel.ok) {
      portfolioDepends = `Business Unit = "${selectedBu}"`;
      const waitMs = postBuDelayBeforePortfolioReadMs();
      if (waitMs > 0) await page.waitForTimeout(waitMs);
    } else {
      portfolioDepends = `Business Unit (select "${selectedBu}" failed: ${sel.error ?? 'unknown'})`;
    }
  }

  const pfRead = await tryReadPicklistOptionsFromModal(page, modal, pfDef.fieldName);
  table.push(picklistTableRow(pfDef, pfRead, portfolioDepends));

  console.table(table);

  const bugs = table.filter((r) => r.Status.startsWith('BUG'));
  if (bugs.length > 0) {
    console.log(`Step 6: ${bugs.length} BUG row(s) (see table).\n`);
    logStructuredQaBugs(
      'Step 6 — structured bug report(s)',
      buildStep6PicklistStructuredBugs(table),
    );
  } else {
    console.log('Step 6: no BUG rows.\n');
  }
}

/**
 * **Step 6:** Business Unit â†’ Portfolio (Excel B+C map, or Picklist Values fallback).
 * `TIBBIYAH_STEP6_PORTFOLIO_UI_SELECT=1` clicks each Portfolio value on the UI.
 */
export async function reportBusinessUnitAndPortfolioPicklistValidation(
  page: Page,
  modal: Locator,
  workbookPath: string,
  sheetIndexOrName: number | string,
): Promise<void> {
  console.log(
    '\nStep 6: Business Unit â†’ Portfolio only (does not validate Procurement Sector / Channel / Request Type).\n',
  );

  const mapRows = await loadBuPortfolioDependencyRowsFromTibbiyahWorkbook(
    workbookPath,
    sheetIndexOrName,
  );

  if (mapRows.length > 0) {
    await reportStep6UsingBuPortfolioColumns(page, modal, mapRows);
  } else {
    await reportStep6BuPortfolioPicklistFallback(
      page,
      modal,
      workbookPath,
      sheetIndexOrName,
    );
  }
}

async function reportStep6UsingBuPortfolioColumns(
  page: Page,
  modal: Locator,
  mapRows: BuPortfolioExcelPair[],
): Promise<void> {
  const buLabel = businessUnitFieldLabelForModal();
  const pfLabel = portfolioFieldLabelForModal();
  const portfolioUiSelect =
    process.env.TIBBIYAH_STEP6_PORTFOLIO_UI_SELECT?.trim() === '1';
  const showBuSummaryTable =
    process.env.TIBBIYAH_STEP6_BU_SUMMARY_TABLE?.trim() === '1';

  console.log(
    '\nStep 6: Business Unit & Portfolio — Excel B+C dependency map (BU and Portfolio columns only).\n',
  );

  const uniqueBus = dedupeBusinessUnitsPreserveOrder(
    mapRows.map((r) => r.businessUnit),
  );
  const buDef: TibbiyahPicklistDefinition = {
    fieldName: buLabel,
    expectedValues: uniqueBus,
  };

  const summaryTable: PicklistComparisonTableRow[] = [];
  if (showBuSummaryTable) {
    const buRead = await tryReadPicklistOptionsFromModal(page, modal, buDef.fieldName);
    summaryTable.push(picklistTableRow(buDef, buRead));
    console.table(summaryTable);
  }

  const depRows: BuPortfolioDependencyTableRow[] = [];
  let currentBu = '';
  let cachedPfOptions: string[] = [];
  let cachedPfReadOk = false;
  const skipPortfolioOptionRead =
    process.env.TIBBIYAH_STEP6_SKIP_PORTFOLIO_OPTION_READ?.trim() === '1';
  const skipPortfolioReadEffective =
    skipPortfolioOptionRead && portfolioUiSelect;

  for (const pr of mapRows) {
    const bu = pr.businessUnit.trim();
    if (bu !== currentBu) {
      const sel = await trySelectPicklistOptionByLabel(
        page,
        modal,
        comboboxAccessibleNamePattern(buLabel),
        bu,
      );
      if (!sel.ok) {
        depRows.push({
          'Business Unit': bu,
          'Excel Portfolio(s)': pr.portfolioRaw || '—',
          'Missing on UI': '—',
          Status: `BUG: Could not select Business Unit (${sel.error ?? 'unknown'})`,
        });
        currentBu = '';
        cachedPfOptions = [];
        cachedPfReadOk = false;
        continue;
      }
      currentBu = bu;
      const waitMs = step6PostBuDelayMs();
      if (waitMs > 0) await page.waitForTimeout(waitMs);
      if (skipPortfolioReadEffective) {
        cachedPfOptions = [];
        cachedPfReadOk = false;
      } else {
        const pfRead = await tryReadPicklistOptionsFromModal(page, modal, pfLabel);
        cachedPfReadOk = pfRead.ok;
        cachedPfOptions = pfRead.ok ? pfRead.options : [];
      }
    }

    const expectedTokens = splitTibbiyahPicklistCellTokens(pr.portfolioRaw);

    if (expectedTokens.length === 0) {
      depRows.push({
        'Business Unit': bu,
        'Excel Portfolio(s)': pr.portfolioRaw || '(empty)',
        'Missing on UI': '—',
        Status: 'OK',
      });
      continue;
    }

    if (portfolioUiSelect) {
      const interPfMs = step6InterPortfolioSelectDelayMs();
      const pfPattern = comboboxAccessibleNamePattern(pfLabel);
      const failedSelections: string[] = [];
      for (const tok of expectedTokens) {
        const t = tok.trim();
        if (!t) continue;
        const selPf = await trySelectPicklistOptionByLabel(
          page,
          modal,
          pfPattern,
          t,
        );
        if (!selPf.ok) failedSelections.push(t);
        if (interPfMs > 0) await page.waitForTimeout(interPfMs);
      }
      const status =
        failedSelections.length === 0
          ? 'OK'
          : `BUG: could not select ${failedSelections.length} Portfolio value(s) on UI`;
      depRows.push({
        'Business Unit': bu,
        'Excel Portfolio(s)': pr.portfolioRaw,
        'Missing on UI':
          failedSelections.length > 0 ? summarizeMissing(failedSelections) : '—',
        Status: status,
      });
      continue;
    }

    if (!cachedPfReadOk) {
      depRows.push({
        'Business Unit': bu,
        'Excel Portfolio(s)': pr.portfolioRaw || '—',
        'Missing on UI': summarizeMissing(expectedTokens),
        Status: `BUG: Could not read Portfolio picklist for Business Unit "${bu}"`,
      });
      continue;
    }

    const missing = expectedTokens.filter(
      (ev) => ev.trim() && !uiHasValue(cachedPfOptions, ev),
    );
    const status =
      missing.length === 0
        ? 'OK'
        : `BUG: ${missing.length} Excel Portfolio value(s) not on UI for this BU`;

    depRows.push({
      'Business Unit': bu,
      'Excel Portfolio(s)': pr.portfolioRaw,
      'Missing on UI': summarizeMissing(missing),
      Status: status,
    });
  }

  logBuPortfolioDependencyTable(depRows);

  const bugSummary =
    summaryTable.filter((r) => r.Status.startsWith('BUG')).length +
    depRows.filter((r) => r.Status.startsWith('BUG')).length;
  if (bugSummary > 0) {
    console.log(`Step 6: ${bugSummary} BUG row(s) (see table(s)).\n`);
    logStructuredQaBugs(
      'Step 6 — structured bug report(s)',
      buildStep6CombinedStructuredBugs(summaryTable, depRows),
    );
  } else {
    console.log('Step 6: no BUG rows.\n');
  }
}

function procurementSectorFieldLabelForModal(): string {
  return (
    process.env.TIBBIYAH_PROCUREMENT_SECTOR_FIELD_LABEL?.trim() ||
    'Procurement Sector'
  );
}

function procurementChannelFieldLabelForModal(): string {
  return (
    process.env.TIBBIYAH_PROCUREMENT_CHANNEL_FIELD_LABEL?.trim() ||
    'Procurement Channel'
  );
}

function requestTypeFieldLabelForModal(): string {
  return process.env.TIBBIYAH_REQUEST_TYPE_FIELD_LABEL?.trim() || 'Request Type';
}

function procurementPostSectorDelayMs(): number {
  const raw = process.env.TIBBIYAH_PROCUREMENT_POST_SECTOR_MS?.trim();
  if (raw && /^\d+$/.test(raw)) return Number.parseInt(raw, 10);
  return 400;
}

function procurementPostChannelDelayMs(): number {
  const raw = process.env.TIBBIYAH_PROCUREMENT_POST_CHANNEL_MS?.trim();
  if (raw && /^\d+$/.test(raw)) return Number.parseInt(raw, 10);
  return 900;
}

/** Workbook marks Request Type as intentionally disabled for this dependency row. */
function excelImpliesRequestTypeDisabled(requestTypeRaw: string): boolean {
  const t = requestTypeRaw.replace(/\s+/g, ' ').trim();
  if (!t) return true;
  if (/^\*?\s*field\s*disabled\s*\*?$/i.test(t)) return true;
  return false;
}

export type ProcurementClassificationTableRow = {
  'Excel row': number;
  'Excel Sector': string;
  'Excel Channel': string;
  'Excel Request Type': string;
  'UI Sector (captured)': string;
  'UI Channel (captured)': string;
  'UI Request Type (captured)': string;
  'Channel enabled': string;
  'Request Type enabled': string;
  'Expected RT enabled': string;
  'Missing RT on UI': string;
  /** Values read from the open Request Type picklist (click-to-open), `;`-separated. */
  'RT picklist options (captured)': string;
  Status: string;
};

/**
 * **Step 7:** Procurement Sector → Channel → Request Type (Excel **B78–D78** area by default).
 * 1) Logs the workbook map. 2) Captures **all** Sector options from the UI, then for **each**
 * Sector all Channel options, then for **each** Channel all Request Type options (full matrix).
 * 3) Compares every Excel dependency row against that matrix. Does **not** throw.
 */
export async function reportProcurementClassificationPicklistValidation(
  page: Page,
  _leadModal: Locator,
  workbookPath: string,
  sheetIndexOrName: number | string,
): Promise<void> {
  try {
    const sectorL = procurementSectorFieldLabelForModal();
    const channelL = procurementChannelFieldLabelForModal();
    const rtL = requestTypeFieldLabelForModal();

    console.log(
      '\nStep 7: Procurement Classification (Sector → Channel → Request Type).\n',
    );

    let mapRows: ProcurementTripleExcelRow[] = [];
    try {
      mapRows = await loadProcurementDependencyRowsFromTibbiyahWorkbook(
        workbookPath,
        sheetIndexOrName,
      );
    } catch (e) {
      console.log(
        `Step 7: could not load Procurement map from workbook: ${(e as Error).message}\n`,
      );
      return;
    }

    const excelPreview = mapRows.map((r) => ({
      'Sheet row': r.sheetRow,
      'Procurement Sector': r.sector,
      'Procurement Channel': r.channel,
      'Request Type (Excel cell)': r.requestTypeRaw || '(empty)',
    }));
    console.log(
      'Workbook reference (Procurement Sector / Channel / Request Type map; default header **B78–D78**):',
    );
    if (excelPreview.length === 0) {
      console.log('(no data rows after header — check workbook or env overrides.)\n');
    } else {
      console.table(excelPreview);
    }

    if (mapRows.length === 0) {
      console.log(
        'Step 7: no dependency rows to validate (skipping UI walk).\n',
      );
      return;
    }

    const dlg = page.getByRole('dialog', { name: /new lead/i }).first();
    await dlg.waitFor({ state: 'visible', timeout: 30_000 }).catch(() => {});
    await waitForLightningStableState(page, { modal: dlg });
    if (page.isClosed()) {
      console.log(
        'Step 7: browser or page is already closed before UI walk (skipping).\n',
      );
      return;
    }
    console.log(
      'Step 7: re-attached to the New Lead dialog for picklist queries (avoids stale locator handles).\n',
    );
    await tryExpandLeadModalSectionByHeading(dlg, /procurement classification/i);

    if (page.isClosed()) {
      console.log('Step 7: page closed before matrix capture (skipping).\n');
      return;
    }

    const restrictSectorsFromExcelOnly =
      process.env.TIBBIYAH_PROCUREMENT_MATRIX_SECTORS_FROM_EXCEL_ONLY?.trim() ===
      '1';
    const excelSectorBag = new Set(
      mapRows.map((r) => r.sector.replace(/\s+/g, ' ').trim()).filter(Boolean),
    );

    console.log(
      '\nStep 7: Full dependent picklist capture — every UI Sector value, then every Channel for that Sector, then every Request Type for that Channel.\n',
    );

    const matrix = await captureFullProcurementPicklistMatrixFromModal(
      page,
      dlg,
      sectorL,
      channelL,
      rtL,
      restrictSectorsFromExcelOnly
        ? { restrictSectors: excelSectorBag }
        : undefined,
    );

    console.log(
      '\n--- Variable 1 · All Procurement Sector values captured from UI ---\n',
    );
    console.log(
      matrix.sectorOptions.length
        ? matrix.sectorOptions.join('  |  ')
        : '(none — Sector picklist could not be read)',
    );

    console.log(
      '\n--- Variable 2 · All Procurement Channel values per Sector (captured from UI) ---\n',
    );
    for (const s of matrix.sectorOptions) {
      const ch = matrix.channelOptionsBySector[s] ?? [];
      console.log(
        `[Sector = "${s}"] → ${ch.length ? ch.join('  |  ') : '(none / not captured)'}`,
      );
    }

    console.log(
      '\n--- Variable 3 · All Request Type values per Sector + Channel (captured from UI) ---\n',
    );
    for (const s of matrix.sectorOptions) {
      for (const c of matrix.channelOptionsBySector[s] ?? []) {
        const opts =
          matrix.requestTypeOptionsByPair[procurementPairKey(s, c)] ?? [];
        console.log(
          `[Sector = "${s}"] [Channel = "${c}"] → ${opts.length ? opts.join('  |  ') : '(none — empty or disabled)'}`,
        );
      }
    }

    if (matrix.captureLog.length > 0) {
      console.log('\nStep 7 capture log:\n' + matrix.captureLog.join('\n') + '\n');
    }

    const matrixSummaryRows = matrix.sectorOptions.flatMap((s) => {
      const channels = matrix.channelOptionsBySector[s] ?? [];
      if (channels.length === 0) {
        return [
          {
            Sector: s,
            Channel: '—',
            'RT options (#)': 0,
            'RT options (preview)': '—',
          },
        ];
      }
      return channels.map((c) => {
        const rt =
          matrix.requestTypeOptionsByPair[procurementPairKey(s, c)] ?? [];
        return {
          Sector: s,
          Channel: c,
          'RT options (#)': rt.length,
          'RT options (preview)': summarizeRtPicklistOptionsForTable(rt, 120),
        };
      });
    });
    console.log('Step 7: matrix summary (one row per Sector + Channel UI pair):\n');
    console.table(matrixSummaryRows);

    const results = buildExcelVsMatrixComparisonRows(mapRows, matrix);

    console.log('Step 7: Excel workbook rows vs captured UI matrix:\n');
    console.table(results);

    const bugs = results.filter((r) => r.Status.startsWith('BUG'));
    if (bugs.length > 0) {
      console.log(`Step 7: ${bugs.length} BUG row(s) (see table).\n`);
      logStructuredQaBugs(
        'Step 7 — Procurement Classification structured bug report(s)',
        buildProcurementClassificationStructuredBugs(results),
      );
    } else {
      console.log('Step 7: no BUG rows.\n');
    }
  } catch (e) {
    console.log(
      `Step 7: validation halted without failing the test: ${(e as Error).message}\n`,
    );
  }
}
