import { expect } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';
import type { TibbiyahConfigRow } from './tibbiyahLeadConfig';
import { isPortfolioLabel } from './tibbiyahLeadConfig';
import { buildStep5StructuredBugs, logStructuredQaBugs } from './qaBugReport';

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

async function comboboxIsUsable(c: Locator): Promise<boolean> {
  if (!(await c.isVisible().catch(() => false))) return false;
  if ((await c.getAttribute('aria-disabled')) === 'true') return false;
  const inp = c.locator('input').first();
  if ((await inp.count()) > 0) {
    if ((await inp.getAttribute('disabled')) !== null) return false;
    if ((await inp.getAttribute('aria-disabled')) === 'true') return false;
  }
  return true;
}

async function comboboxIsExpanded(c: Locator): Promise<boolean> {
  if ((await c.getAttribute('aria-expanded')) === 'true') return true;
  const inp = c.locator('input').first();
  return (
    (await inp.count()) > 0 && (await inp.getAttribute('aria-expanded')) === 'true'
  );
}

/**
 * Resolves the combobox instance to drive for this accessible name:
 * 1) Prefer **expanded + usable** (dropdown already open for this control).
 * 2) Else **last usable visible** (duplicate combobox names / dependent picklists are often last).
 * 3) Else **first usable visible**.
 * Avoids returning a visible but **disabled** clone (common cause: nothing runs after “Opening … dropdown”).
 */
export async function pickVisibleComboboxInModal(
  modal: Locator,
  namePattern: RegExp,
): Promise<Locator> {
  try {
    const all = modal.getByRole('combobox', { name: namePattern });
    const n = await all.count();
    let expandedUsable: Locator | null = null;
    let firstUsable: Locator | null = null;
    let lastUsable: Locator | null = null;
    for (let i = 0; i < n; i++) {
      const c = all.nth(i);
      if (!(await comboboxIsUsable(c))) continue;
      firstUsable ??= c;
      lastUsable = c;
      if (await comboboxIsExpanded(c)) expandedUsable = c;
    }
    if (expandedUsable) return expandedUsable;
    if (lastUsable) return lastUsable;
    if (firstUsable) return firstUsable;
    for (let i = 0; i < n; i++) {
      const c = all.nth(i);
      if (await c.isVisible().catch(() => false)) return c;
    }
    return all.first();
  } catch {
    /** Session or dialog gone — return a non-matching combobox so callers treat field as absent. */
    return modal.getByRole('combobox', { name: '__tibbiyah_no_combobox__' });
  }
}

/** Readable label for a Lightning listbox row (innerText is often empty in LWC). */
async function lightningListboxOptionLabel(opt: Locator): Promise<string> {
  let t = (await opt.innerText().catch(() => ''))?.replace(/\s+/g, ' ').trim() ?? '';
  if (t) return t;
  t = (await opt.textContent().catch(() => ''))?.replace(/\s+/g, ' ').trim() ?? '';
  if (t) return t;
  const aria = (await opt.getAttribute('aria-label').catch(() => null))?.trim();
  if (aria) return aria;
  const title = (await opt.getAttribute('title').catch(() => null))?.trim();
  if (title) return title;
  const dataVal = (await opt.getAttribute('data-value').catch(() => null))?.trim();
  if (dataVal) return dataVal;
  const span = opt.locator('.slds-media__body, span.slds-truncate').first();
  if ((await span.count()) > 0) {
    const st = (await span.innerText().catch(() => ''))?.replace(/\s+/g, ' ').trim() ?? '';
    if (st) return st;
  }
  return '';
}

export type LeadFieldComparisonRow = {
  Type: string;
  'Excel label': string;
  'Required (Excel)': string;
  'Present on UI': string;
  'Required (UI)': string;
  Status: string;
};

const STEP5_FIELD_TABLE_COLUMNS = [
  'Type',
  'Excel label',
  'Required (Excel)',
  'Present on UI',
  'Required (UI)',
  'Status',
] as const satisfies ReadonlyArray<keyof LeadFieldComparisonRow>;

function logStep5LeadFieldComparisonTable(rows: LeadFieldComparisonRow[]): void {
  if (rows.length === 0) {
    console.log('(no rows)\n');
    return;
  }
  const widths = STEP5_FIELD_TABLE_COLUMNS.map((col) =>
    Math.max(col.length, ...rows.map((r) => String(r[col]).length)),
  );
  const pad = (s: string, w: number) => s.padEnd(w, ' ');
  const line = (cells: string[]) =>
    cells.map((c, i) => pad(c, widths[i])).join('  ');
  console.log(line([...STEP5_FIELD_TABLE_COLUMNS]));
  console.log(
    line(STEP5_FIELD_TABLE_COLUMNS.map((_, i) => '-'.repeat(widths[i]))),
  );
  for (const r of rows) {
    console.log(line(STEP5_FIELD_TABLE_COLUMNS.map((c) => String(r[c]))));
  }
  console.log('');
}

function fieldLabelMatchPattern(fieldLabel: string): RegExp {
  const trimmed = fieldLabel.replace(/\u00a0/g, ' ').replace(/\*/g, '').trim();
  return new RegExp(escapeRegExp(trimmed), 'i');
}

/**
 * Excel often appends hints in parentheses (e.g. Address (Street, City, ...)) while the
 * Lead modal shows only "Address". Try the full label first, then repeatedly strip a
 * trailing ` ( ... )` segment for UI matching.
 */
function fieldLabelCandidatesForUiLookup(fieldLabel: string): string[] {
  const raw = fieldLabel.replace(/\u00a0/g, ' ').trim();
  const out: string[] = [];
  const seen = new Set<string>();
  const push = (s: string) => {
    const t = s.replace(/\s+/g, ' ').trim();
    if (t && !seen.has(t)) {
      seen.add(t);
      out.push(t);
    }
  };
  push(raw);
  let cur = raw;
  for (;;) {
    const next = cur.replace(/\s*\([^)]*\)\s*$/u, '').trim();
    if (!next || next === cur) break;
    cur = next;
    push(cur);
  }
  return out;
}

async function tryLeadModalFormElementRowForLabel(
  modal: Locator,
  labelForMatch: string,
): Promise<{ presentOnUi: boolean; requiredOnUi: boolean } | null> {
  const trimmed = labelForMatch.replace(/\u00a0/g, ' ').trim();
  if (!trimmed) return null;
  const nameRe = new RegExp(escapeRegExp(trimmed), 'i');
  const formRows = modal.locator('.slds-form-element');
  const n = await formRows.count();
  for (let i = 0; i < n; i++) {
    const row = formRows.nth(i);
    const labelLoc = row
      .locator(
        '.slds-form-element__label, .slds-form-element__legend, legend, label',
      )
      .first();
    if (!(await labelLoc.isVisible().catch(() => false))) continue;
    const txt = await labelLoc.innerText();
    if (!nameRe.test(txt)) continue;
    const requiredOnUi = await row
      .locator('abbr.slds-required, abbr[title="required"], .slds-required')
      .isVisible()
      .catch(() => false);
    return { presentOnUi: true, requiredOnUi };
  }
  return null;
}

/**
 * Lightning booleans often expose `role="checkbox"` or `role="switch"` with the label
 * as the accessible name, while the visible `.slds-form-element__label` walk misses them
 * (shadow DOM, sr-only label, or non-standard layout).
 */
async function getLeadModalCheckboxLikeControlUiState(
  modal: Locator,
  fieldLabel: string,
): Promise<{ presentOnUi: boolean; requiredOnUi: boolean } | null> {
  const trimmed = fieldLabel.replace(/\u00a0/g, ' ').trim();
  if (!trimmed) return null;
  const namePattern = fieldLabelMatchPattern(fieldLabel);

  for (const role of ['checkbox', 'switch'] as const) {
    const control = modal.getByRole(role, { name: namePattern }).first();
    if (!(await control.isVisible().catch(() => false))) continue;

    let requiredOnUi =
      ((await control.getAttribute('aria-required')) ?? '').toLowerCase() ===
      'true';

    if (!requiredOnUi) {
      const hostRow = control.locator(
        'xpath=ancestor::*[contains(concat(" ", normalize-space(@class), " "), " slds-form-element ")][1]',
      );
      if ((await hostRow.count()) > 0) {
        requiredOnUi = await hostRow
          .locator(
            'abbr.slds-required, abbr[title="required"], .slds-required',
          )
          .first()
          .isVisible()
          .catch(() => false);
      }
    }

    return { presentOnUi: true, requiredOnUi };
  }

  return null;
}

/**
 * Returns whether the modal shows a matching field row and a Lightning required marker.
 */
export async function getLeadModalFieldUiState(
  modal: Locator,
  fieldLabel: string,
): Promise<{ presentOnUi: boolean; requiredOnUi: boolean }> {
  for (const label of fieldLabelCandidatesForUiLookup(fieldLabel)) {
    const hit = await tryLeadModalFormElementRowForLabel(modal, label);
    if (hit) return hit;
  }

  for (const label of fieldLabelCandidatesForUiLookup(fieldLabel)) {
    const checkboxLike = await getLeadModalCheckboxLikeControlUiState(
      modal,
      label,
    );
    if (checkboxLike) return checkboxLike;
  }

  return { presentOnUi: false, requiredOnUi: false };
}

/**
 * Whether a Lightning section heading / card title for this modal is visible.
 */
export async function isLeadModalSectionVisible(
  modal: Locator,
  sectionTitle: string,
): Promise<boolean> {
  const exact = new RegExp(
    '^\\s*' + escapeRegExp(sectionTitle.trim()) + '\\s*$',
    'i',
  );
  const loose = new RegExp(escapeRegExp(sectionTitle.trim()), 'i');

  const byHeading = modal.getByRole('heading', { name: exact });
  if (await byHeading.first().isVisible().catch(() => false)) return true;

  for (const level of [2, 3, 4] as const) {
    const h = modal.getByRole('heading', { name: loose, level });
    if (await h.first().isVisible().catch(() => false)) return true;
  }

  const blockTitles = modal.locator(
    [
      '.slds-section__title',
      '.slds-section__title-action',
      'records-form-section-header',
      'records-record-layout-section .section-header-label',
      'lightning-card .slds-card__header-title',
      'slot[name="title"]',
    ].join(', '),
  );
  const n = await blockTitles.filter({ hasText: loose }).count();
  if (n > 0) {
    const first = blockTitles.filter({ hasText: loose }).first();
    if (await first.isVisible().catch(() => false)) return true;
  }

  const legend = modal.locator('legend').filter({ hasText: loose });
  if (await legend.first().isVisible().catch(() => false)) return true;

  return false;
}

/**
 * SLDS combobox dropdown listboxes on the page (often portaled under `body`), excluding
 * dual-list `source-list-*` / `selected-list-*` panels.
 */
function lightningComboboxListboxes(page: Page): Locator {
  return page
    .locator(
      [
        '[role="listbox"]',
        ':not([id^="source-list-"])',
        ':not([id^="selected-list-"])',
        ':not([id^="target-list-"])',
      ].join(''),
    )
    .filter({ has: page.locator('[role="option"]') });
}

/**
 * When a dropdown is open, the **last visible** listbox with options is usually the active combobox
 * panel (most recently opened). Used when `aria-controls` resolves to a host that is not yet visible
 * or does not contain `[role=option]` nodes (common with Lightning portals).
 */
async function firstBackwardsVisibleListboxWithOptions(
  page: Page,
): Promise<Locator | null> {
  const roots = lightningComboboxListboxes(page);
  const n = Math.min(await roots.count(), 40);
  for (let i = n - 1; i >= 0; i--) {
    const lb = roots.nth(i);
    if (!(await lb.isVisible().catch(() => false))) continue;
    if ((await lb.getByRole('option').count()) > 0) return lb;
  }
  return null;
}

/** Legacy “any open combobox listbox” tail (prefer {@link firstBackwardsVisibleListboxWithOptions} when panel is open). */
function activePicklistListbox(page: Page): Locator {
  return lightningComboboxListboxes(page).last();
}

/**
 * Primary visible line for a picklist option (Lightning often adds a second line or symbols).
 */
export function picklistOptionDisplayText(raw: string): string {
  const t = raw.replace(/\u00a0/g, ' ');
  const lines = t
    .split(/\r?\n/)
    .map((l) => l.replace(/\s+/g, ' ').trim())
    .filter((l) => l.length > 0);
  const primary = lines[0] ?? t.replace(/\s+/g, ' ').trim();
  return primary
    .replace(/^[\s\u2713\u2714\u2611\u2610✓✔·•\-–—]+/u, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Lowercase key for comparing Excel picklist values to UI option text. */
export function normPicklistMatchKey(raw: string): string {
  return picklistOptionDisplayText(raw)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/[\u2013\u2014\u2212]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

export function picklistValuesMatchLoose(uiOptionRaw: string, expectedRaw: string): boolean {
  const e = normPicklistMatchKey(expectedRaw);
  if (!e) return true;
  const keys = new Set<string>();
  for (const line of uiOptionRaw.split(/\r?\n/)) {
    const k = normPicklistMatchKey(line);
    if (k) keys.add(k);
  }
  const primary = normPicklistMatchKey(uiOptionRaw);
  keys.add(primary);
  for (const oN of keys) {
    if (oN === e) return true;
    if (e.length >= 6 && oN.length >= 6 && (oN.includes(e) || e.includes(oN))) {
      return true;
    }
  }
  return false;
}

/**
 * Playwright `name` pattern for the combobox tied to a field label (tight match for BU / Portfolio).
 */
export function comboboxAccessibleNamePattern(fieldLabel: string): RegExp {
  const stripped = fieldLabel.replace(/\*/g, '').replace(/\s+/g, ' ').trim();
  const t = stripped.toLowerCase();
  if (t === 'business unit' || /^business\s+unit$/i.test(stripped)) {
    return /^\s*\*?\s*Business\s+Unit\b/i;
  }
  if (t === 'portfolio' || /^portfolio$/i.test(stripped)) {
    return /^\s*\*?\s*Portfolio\b/i;
  }
  if (/^procurement\s+sector$/i.test(stripped)) {
    return /^\s*\*?\s*Procurement\s+Sector\b/i;
  }
  if (/^procurement\s+channel$/i.test(stripped)) {
    return /^\s*\*?\s*Procurement\s+Channel\b/i;
  }
  if (t === 'request type' || /^request\s+type$/i.test(stripped)) {
    /** Lightning often puts extra text in the accessible name; match the words anywhere. */
    return /\bRequest\s+Type\b/i;
  }
  return new RegExp(escapeRegExp(stripped), 'i');
}

/**
 * Finds the combobox inside the same `.slds-form-element` as the visible label/legend.
 * More reliable than `getByRole('combobox', { name })` when the accessible name does not match exactly.
 */
export async function pickComboboxViaFormElementLabel(
  modal: Locator,
  fieldLabel: string,
): Promise<Locator | null> {
  const want = fieldLabel
    .replace(/\u00a0/g, ' ')
    .replace(/\*/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!want) return null;
  const wantKey = normPicklistMatchKey(want);
  const blocks = modal.locator('.slds-form-element');
  const n = await blocks.count();
  for (let i = 0; i < n; i++) {
    const block = blocks.nth(i);
    const lbl = block
      .locator('.slds-form-element__label, label, legend')
      .first();
    const raw = await lbl.innerText().catch(() => '');
    const cleaned = raw
      .replace(/\u00a0/g, ' ')
      .replace(/\*/g, '')
      .replace(/\s+/g, ' ')
      .trim();
    if (!cleaned) continue;
    const key = normPicklistMatchKey(cleaned);
    if (key !== wantKey && !cleaned.toLowerCase().includes(want.toLowerCase())) {
      continue;
    }
    const combo = block.getByRole('combobox').first();
    if ((await combo.count()) === 0) continue;
    if (await combo.isVisible().catch(() => false)) return combo;
  }
  return null;
}

export async function resolveComboboxForPicklistField(
  modal: Locator,
  fieldLabel: string,
  namePattern: RegExp,
  opts?: { preferFormRow?: boolean },
): Promise<Locator> {
  if (opts?.preferFormRow && fieldLabel.trim()) {
    const via = await pickComboboxViaFormElementLabel(modal, fieldLabel.trim());
    if (via && (await comboboxIsUsable(via))) return via;
  }
  return pickVisibleComboboxInModal(modal, namePattern);
}

/** Expands a collapsed Lightning accordion / section header on the Lead modal (best-effort). */
export async function tryExpandLeadModalSectionByHeading(
  modal: Locator,
  headingPattern: RegExp,
): Promise<void> {
  const btn = modal.getByRole('button', { name: headingPattern }).first();
  if (await btn.isVisible().catch(() => false)) {
    const exp = ((await btn.getAttribute('aria-expanded')) ?? '').toLowerCase();
    if (exp === 'false') await btn.click({ timeout: 3_000 }).catch(() => {});
    return;
  }
  const hit = modal.locator('button, h2, h3').filter({ hasText: headingPattern }).first();
  if (await hit.isVisible().catch(() => false)) {
    await hit.click({ timeout: 3_000 }).catch(() => {});
  }
}

/**
 * For a field label, whether any matching combobox is visible and at least one is **usable**
 * (not `aria-disabled` / `disabled`). Duplicate accessible names (dependent picklists) are scanned.
 */
export async function getLeadModalPicklistComboboxInteractionState(
  modal: Locator,
  fieldLabel: string,
): Promise<{ onUi: boolean; enabled: boolean }> {
  try {
    const pattern = comboboxAccessibleNamePattern(fieldLabel);
    const all = modal.getByRole('combobox', { name: pattern });
    const n = await all.count();
    let onUi = false;
    let enabled = false;
    for (let i = 0; i < n; i++) {
      const c = all.nth(i);
      if (!(await c.isVisible().catch(() => false))) continue;
      onUi = true;
      if (await comboboxIsUsable(c)) enabled = true;
    }
    return { onUi, enabled };
  } catch {
    return { onUi: false, enabled: false };
  }
}

/**
 * Reads the primary displayed value from the **last visible** combobox for this label
 * (works when the control is disabled; {@link pickVisibleComboboxInModal} skips disabled nodes).
 */
export async function readPicklistComboboxDisplayedValueAny(
  modal: Locator,
  fieldLabel: string,
): Promise<string> {
  try {
    const pattern = comboboxAccessibleNamePattern(fieldLabel);
    const all = modal.getByRole('combobox', { name: pattern });
    const n = await all.count();
    for (let i = n - 1; i >= 0; i--) {
      const c = all.nth(i);
      if (!(await c.isVisible().catch(() => false))) continue;
      const innerText = await c.innerText().catch(() => '');
      const inputVal = await c
        .locator('input')
        .first()
        .inputValue()
        .catch(() => '');
      const merged = [innerText, inputVal].filter(Boolean).join('\n');
      const t =
        picklistOptionDisplayText(merged) || picklistOptionDisplayText(inputVal);
      if (t) return t.replace(/\s+/g, ' ').trim();
    }
    return '';
  } catch {
    return '';
  }
}

/**
 * Uses `aria-controls` from the Lightning combobox when present so we read the **correct**
 * listbox (not another `[role="listbox"]` on the page from `.last()`).
 */
export async function resolvePicklistListboxAfterComboOpened(
  page: Page,
  combo: Locator,
): Promise<Locator> {
  const ariaWait = listboxAriaResolveWaitMs();
  if (ariaWait > 0) await page.waitForTimeout(ariaWait);
  let controls = (await combo.getAttribute('aria-controls'))?.trim() ?? '';
  if (!controls) {
    const input = combo.locator('input[aria-controls]').first();
    if ((await input.count()) > 0) {
      controls = (await input.getAttribute('aria-controls'))?.trim() ?? '';
    }
  }
  if (controls) {
    const escaped = controls.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    const host = page.locator(`[id="${escaped}"]`).first();
    if ((await host.count()) > 0) {
      const innerList = host.locator('[role="listbox"]').first();
      /** Lightning often paints the portal panel a tick after `aria-controls` is set; without this wait we fall through to `activePicklistListbox().last()` and hang on the wrong listbox (common after changing picklist values). */
      if ((await innerList.count()) > 0) {
        try {
          await innerList.waitFor({ state: 'visible', timeout: 5_000 });
          return innerList;
        } catch {
          /* inner listbox not visible yet */
        }
      }
      const role = await host.getAttribute('role');
      if (role === 'listbox') {
        try {
          await host.waitFor({ state: 'visible', timeout: 3_000 });
          return host;
        } catch {
          /* */
        }
      }
    }
  }
  const openPanel = await firstBackwardsVisibleListboxWithOptions(page);
  if (openPanel) return openPanel;
  return activePicklistListbox(page);
}

function picklistActionDelayMs(): number {
  const raw = process.env.TIBBIYAH_PICKLIST_ACTION_DELAY_MS?.trim();
  if (raw && /^\d+$/.test(raw)) return Number.parseInt(raw, 10);
  return 0;
}

/** Max wait for combobox visible (smaller = faster failures; success usually resolves early). */
export function picklistComboVisibilityTimeoutMs(): number {
  const raw = process.env.TIBBIYAH_PICKLIST_COMBO_TIMEOUT_MS?.trim();
  if (raw && /^\d+$/.test(raw)) {
    return Math.max(1_500, Number.parseInt(raw, 10));
  }
  return 8_000;
}

/** Max wait for picklist panel after opening combobox. */
export function picklistListboxVisibilityTimeoutMs(): number {
  const raw = process.env.TIBBIYAH_PICKLIST_LISTBOX_TIMEOUT_MS?.trim();
  if (raw && /^\d+$/.test(raw)) {
    return Math.max(1_500, Number.parseInt(raw, 10));
  }
  return 8_000;
}

function listboxAriaResolveWaitMs(): number {
  const raw = process.env.TIBBIYAH_LISTBOX_ARIA_WAIT_MS?.trim();
  if (raw && /^\d+$/.test(raw)) {
    return Math.max(0, Number.parseInt(raw, 10));
  }
  return 40;
}

/** Throws immediately if the browser/context was closed (avoids hanging waits). */
export function assertPageNotClosed(page: Page): void {
  if (page.isClosed()) {
    console.error('Page closed detected');
    throw new Error('Page is closed');
  }
}

function lightningStableSettleMs(): number {
  const raw = process.env.TIBBIYAH_LIGHTNING_STABLE_SETTLE_MS?.trim();
  if (raw && /^\d+$/.test(raw)) return Math.max(0, Number.parseInt(raw, 10));
  return 0;
}

function lightningStableAfterComboClickMs(): number {
  const raw = process.env.TIBBIYAH_LIGHTNING_AFTER_COMBO_CLICK_MS?.trim();
  if (raw && /^\d+$/.test(raw)) return Math.max(0, Number.parseInt(raw, 10));
  return 120;
}

export type WaitLightningStableStateOptions = {
  /** Limit spinner waits to this subtree (recommended: Lead modal). */
  modal?: Locator;
  /** Optional extra idle after spinners (default **0**). */
  settleMs?: number;
  /** Best-effort `networkidle` (never fails the step if it times out). */
  waitNetworkIdle?: boolean;
};

/**
 * Wait for Salesforce Lightning UI to settle: spinners hidden, optional `networkidle`, optional settle.
 * Does **not** fail if spinners are absent. Always guards on `page.isClosed()`.
 */
export async function waitForLightningStableState(
  page: Page,
  options?: WaitLightningStableStateOptions,
): Promise<void> {
  if (page.isClosed()) return;
  const root = options?.modal ?? page.locator('body');
  const spinnerLoc = root.locator(
    [
      '.slds-spinner:not(.slds-spinner_inline)',
      '.slds-spinner_container .slds-spinner',
      'lightning-spinner',
      '[class*="Spinner"]',
    ].join(', '),
  );
  const n = Math.min(await spinnerLoc.count(), 20);
  for (let i = 0; i < n; i++) {
    if (page.isClosed()) return;
    const s = spinnerLoc.nth(i);
    if (await s.isVisible().catch(() => false)) {
      await s
        .waitFor({ state: 'hidden', timeout: 20_000 })
        .catch(() => undefined);
    }
  }
  const globalSpin = page.locator('.slds-spinner').first();
  if (await globalSpin.isVisible().catch(() => false)) {
    await globalSpin
      .waitFor({ state: 'hidden', timeout: 8_000 })
      .catch(() => undefined);
  }
  if (options?.waitNetworkIdle) {
    await page.waitForLoadState('networkidle', { timeout: 6_000 }).catch(() => undefined);
  }
  const settle = options?.settleMs ?? lightningStableSettleMs();
  if (settle > 0) {
    if (page.isClosed()) return;
    await page.waitForTimeout(settle);
  }
}

/**
 * True if the combobox shows `expected` as the **selected** value (primary line / input only).
 * Avoids `norm(blob).includes(norm(expected))`, which could treat unrelated substrings as a match
 * during multi-value or dependent-picklist checks.
 */
export async function verifyPicklistComboboxShowsValue(
  modal: Locator,
  fieldLabel: string,
  expected: string,
): Promise<boolean> {
  const pattern = comboboxAccessibleNamePattern(fieldLabel);
  const combo = await pickVisibleComboboxInModal(modal, pattern);
  if (!(await combo.isVisible().catch(() => false))) return false;
  const innerText = await combo.innerText().catch(() => '');
  const inputVal = await combo
    .locator('input')
    .first()
    .inputValue()
    .catch(() => '');
  const primary = picklistOptionDisplayText(
    [innerText, inputVal].filter(Boolean).join('\n'),
  );
  const fromInput = picklistOptionDisplayText(inputVal);
  const expKey = normPicklistMatchKey(expected);
  if (!expKey) return true;
  for (const cand of [primary, fromInput, inputVal, innerText]) {
    const t = cand.replace(/\s+/g, ' ').trim();
    if (!t) continue;
    const line = picklistOptionDisplayText(t);
    if (!line) continue;
    if (normPicklistMatchKey(line) === expKey) return true;
    if (picklistValuesMatchLoose(line, expected)) return true;
  }
  return false;
}

/**
 * Closes an open combobox listbox **without** using Escape, which often dismisses the
 * whole New Lead modal in SLDS. Clicks a **safe** in-modal area (body padding, not the
 * header close control) so the Lead form stays open. No-op when no listbox is open.
 */
export async function dismissOpenPicklistListboxIfVisible(
  page: Page,
  modal: Locator,
): Promise<void> {
  if (page.isClosed()) return;
  const listbox = activePicklistListbox(page);
  if (!(await listbox.isVisible().catch(() => false))) return;

  const body = modal
    .locator(
      [
        '.slds-modal__content',
        'records-modal',
        'lightning-modal-body',
        'slot[name="modalBody"]',
      ].join(', '),
    )
    .first();
  if (await body.isVisible().catch(() => false)) {
    await body.click({ position: { x: 12, y: 12 }, timeout: 8_000 });
  } else {
    const title = modal.getByRole('heading', { name: /new lead/i }).first();
    if (await title.isVisible().catch(() => false)) {
      await title.click({ position: { x: 4, y: 4 }, timeout: 8_000 });
    } else {
      await modal.click({ position: { x: 40, y: 80 }, timeout: 8_000 });
    }
  }

  if (page.isClosed()) return;
  try {
    await listbox.waitFor({ state: 'hidden', timeout: 8_000 });
  } catch {
    if (await listbox.isVisible().catch(() => false)) {
      console.warn(
        'dismissOpenPicklistListboxIfVisible: combobox listbox still visible after dismiss click — continuing (Step 5 validation will proceed).',
      );
    }
  }
}

/** @alias {@link dismissOpenPicklistListboxIfVisible} */
export async function dismissOpenPicklistListboxInLeadModal(
  page: Page,
  modal: Locator,
): Promise<void> {
  return dismissOpenPicklistListboxIfVisible(page, modal);
}

/**
 * Best-effort Business Unit selection; does not throw (used for comparison runs).
 */
export async function trySelectFirstNonEmptyPicklistOption(
  page: Page,
  modal: Locator,
  fieldNamePattern: RegExp,
): Promise<{ ok: boolean; error?: string }> {
  try {
    await dismissOpenPicklistListboxInLeadModal(page, modal);
    const combo = await pickVisibleComboboxInModal(modal, fieldNamePattern);
    await combo.scrollIntoViewIfNeeded().catch(() => {});
    await combo.waitFor({
      state: 'visible',
      timeout: picklistComboVisibilityTimeoutMs(),
    });
    await combo.click();
    const listbox = await resolvePicklistListboxAfterComboOpened(page, combo);
    await listbox.waitFor({
      state: 'visible',
      timeout: picklistListboxVisibilityTimeoutMs(),
    });
    const slow = picklistActionDelayMs();
    if (slow > 0) await page.waitForTimeout(slow);
    const options = listbox.getByRole('option');
    const count = await options.count();
    for (let i = 0; i < count; i++) {
      const opt = options.nth(i);
      const raw = await opt.innerText();
      const t = picklistOptionDisplayText(raw);
      if (!t) continue;
      if (/^(--\s*)?none\s*(--)?$/i.test(t)) continue;
      await opt.click();
      return { ok: true };
    }
    return { ok: false, error: 'No selectable picklist option' };
  } catch (e) {
    return { ok: false, error: (e as Error).message };
  }
}

/** Default Business Unit value before validating dependent Portfolio (override with `TIBBIYAH_BUSINESS_UNIT_VALUE`). */
export const DEFAULT_BUSINESS_UNIT_FOR_PORTFOLIO = 'Al Hammad Hybrid';

export function businessUnitValueForPortfolioValidation(): string {
  const raw = process.env.TIBBIYAH_BUSINESS_UNIT_VALUE?.trim();
  return raw || DEFAULT_BUSINESS_UNIT_FOR_PORTFOLIO;
}

/**
 * Lightning picklists often expose options as `[role=option]`, `title=`, plain text,
 * or nested `span`s (recorded Salesforce flows). Tries listbox-local then dialog-wide
 * locators (options may render in a portal tied to the dialog).
 */
async function tryClickOpenedPicklistOptionWithLightningFallbacks(
  page: Page,
  modal: Locator,
  listbox: Locator,
  target: string,
): Promise<boolean> {
  const t = target.trim();
  if (!t) return false;

  const tryClick = async (loc: Locator): Promise<boolean> => {
    try {
      if (!(await loc.isVisible().catch(() => false))) return false;
      await loc.scrollIntoViewIfNeeded({ timeout: 2_000 }).catch(() => {});
      await loc.click({ timeout: 6_000, force: true });
      return true;
    } catch {
      return false;
    }
  };

  const nameExact = new RegExp(`^\\s*${escapeRegExp(t)}\\s*$`, 'i');
  const nameLoose = new RegExp(escapeRegExp(t), 'i');

  if (await tryClick(listbox.getByRole('option', { name: nameLoose }).first())) {
    return true;
  }
  if (await tryClick(listbox.getByRole('option', { name: t, exact: true }).first())) {
    return true;
  }
  if (await tryClick(listbox.getByRole('option', { name: nameExact }).first())) {
    return true;
  }
  if (await tryClick(modal.getByRole('option', { name: t, exact: true }).first())) {
    return true;
  }
  if (await tryClick(modal.getByRole('option', { name: nameExact }).first())) {
    return true;
  }
  if (await tryClick(page.getByRole('option', { name: t, exact: true }).first())) {
    return true;
  }
  if (await tryClick(page.getByRole('option', { name: nameExact }).first())) {
    return true;
  }

  if (await tryClick(modal.getByTitle(t, { exact: true }).first())) return true;
  if (await tryClick(page.getByTitle(t, { exact: true }).first())) return true;

  if (await tryClick(modal.getByText(t, { exact: true }).first())) return true;

  const spanExact = listbox.locator('span').filter({ hasText: nameExact }).first();
  if (await tryClick(spanExact)) return true;
  if (await tryClick(listbox.locator('span').filter({ hasText: t }).first())) return true;
  if (await tryClick(modal.locator('span').filter({ hasText: nameExact }).first())) {
    return true;
  }
  if (await tryClick(modal.locator('span').filter({ hasText: t }).first())) return true;
  if (await tryClick(page.locator('span').filter({ hasText: nameExact }).first())) {
    return true;
  }
  if (await tryClick(page.locator('span').filter({ hasText: t }).first())) return true;

  return false;
}

/**
 * Opens a picklist and selects the option whose label matches (case-insensitive, trimmed).
 */
export async function trySelectPicklistOptionByLabel(
  page: Page,
  modal: Locator,
  fieldNamePattern: RegExp,
  optionLabel: string,
  pickOpts?: { formElementLabel?: string },
): Promise<{ ok: boolean; error?: string }> {
  const target = optionLabel.trim();
  if (!target) return { ok: false, error: 'Empty option label' };
  try {
    await dismissOpenPicklistListboxInLeadModal(page, modal);
    const combo = await resolveComboboxForPicklistField(
      modal,
      pickOpts?.formElementLabel?.trim() ?? '',
      fieldNamePattern,
      { preferFormRow: Boolean(pickOpts?.formElementLabel?.trim()) },
    );
    await combo.scrollIntoViewIfNeeded({ block: 'center' }).catch(() => {});
    await combo.waitFor({
      state: 'visible',
      timeout: picklistComboVisibilityTimeoutMs(),
    });
    await combo.click({ force: true });
    const listbox = await resolvePicklistListboxAfterComboOpened(page, combo);
    await listbox.waitFor({
      state: 'visible',
      timeout: picklistListboxVisibilityTimeoutMs(),
    });
    const slow = picklistActionDelayMs();
    if (slow > 0) await page.waitForTimeout(slow);
    const optPat = new RegExp(`^\\s*${escapeRegExp(target)}\\s*$`, 'i');
    const namedFirst = listbox.getByRole('option', { name: optPat }).first();
    if (await namedFirst.isVisible().catch(() => false)) {
      await namedFirst.scrollIntoViewIfNeeded().catch(() => {});
      await namedFirst.click({ force: true });
      return { ok: true };
    }
    const options = listbox.getByRole('option');
    const count = await options.count();
    const targetKey = normPicklistMatchKey(target);
    for (let i = 0; i < count; i++) {
      const opt = options.nth(i);
      const raw = await lightningListboxOptionLabel(opt);
      if (!raw || /^(--\s*)?none\s*(--)?$/i.test(picklistOptionDisplayText(raw))) continue;
      if (picklistValuesMatchLoose(raw, target)) {
        await opt.scrollIntoViewIfNeeded().catch(() => {});
        await opt.click({ force: true });
        return { ok: true };
      }
      const tKey = normPicklistMatchKey(raw);
      if (tKey === targetKey) {
        await opt.scrollIntoViewIfNeeded().catch(() => {});
        await opt.click({ force: true });
        return { ok: true };
      }
    }
    const byName = listbox.getByRole('option', {
      name: new RegExp(`^\\s*${escapeRegExp(target)}\\s*$`, 'i'),
    });
    const named = byName.first();
    if (await named.isVisible().catch(() => false)) {
      await named.scrollIntoViewIfNeeded().catch(() => {});
      await named.click({ force: true });
      return { ok: true };
    }
    if (await tryClickOpenedPicklistOptionWithLightningFallbacks(page, modal, listbox, target)) {
      return { ok: true };
    }
    return { ok: false, error: `Option "${target}" not found in picklist` };
  } catch (e) {
    return { ok: false, error: (e as Error).message };
  }
}

function procurementRtClickGapMs(): number {
  const raw = process.env.TIBBIYAH_PROCUREMENT_RT_CLICK_GAP_MS?.trim();
  if (raw && /^\d+$/.test(raw)) return Number.parseInt(raw, 10);
  return 120;
}

/**
 * Opens the picklist for `fieldLabel` and selects each token in order (closes between picks).
 * Used when reading options fails but values may still be present and clickable (Step 7 recovery).
 */
export async function trySelectPicklistTokensSequentially(
  page: Page,
  modal: Locator,
  fieldLabel: string,
  tokens: string[],
): Promise<{ clicked: string[]; failed: string[] }> {
  if (page.isClosed()) {
    return {
      clicked: [],
      failed: tokens.map((t) => t.trim()).filter(Boolean),
    };
  }
  const pattern = comboboxAccessibleNamePattern(fieldLabel);
  const clicked: string[] = [];
  const failed: string[] = [];
  const gap = procurementRtClickGapMs();
  for (let i = 0; i < tokens.length; i++) {
    if (page.isClosed()) {
      for (let j = i; j < tokens.length; j++) {
        const u = tokens[j].replace(/\s+/g, ' ').trim();
        if (u) failed.push(u);
      }
      break;
    }
    const tok = tokens[i];
    const t = tok.replace(/\s+/g, ' ').trim();
    if (!t) continue;
    const sel = await trySelectPicklistOptionByLabel(page, modal, pattern, t, {
      formElementLabel: fieldLabel,
    });
    if (sel.ok) clicked.push(t);
    else failed.push(t);
    if (gap > 0) await page.waitForTimeout(gap);
  }
  return { clicked, failed };
}

function yn(v: boolean): string {
  return v ? 'Yes' : 'No';
}

function naUnlessPresent(present: boolean, required: boolean): string {
  if (!present) return '—';
  return yn(required);
}

function classifyFieldBug(
  excelReq: boolean,
  present: boolean,
  uiReq: boolean,
): string {
  if (!present) {
    return excelReq
      ? 'BUG: Not on UI (Excel expects field)'
      : 'BUG: Not on UI (listed in Excel)';
  }
  if (excelReq !== uiReq) {
    return excelReq
      ? 'BUG: Excel required; UI not required'
      : 'BUG: Excel optional; UI required';
  }
  return 'OK';
}

function classifySectionBug(present: boolean): string {
  if (!present) return 'BUG: Section not on Lead modal';
  return 'OK';
}

function businessUnitFieldLabelForModal(): string {
  return process.env.TIBBIYAH_BUSINESS_UNIT_FIELD_LABEL?.trim() || 'Business Unit';
}

/** Wait after selecting BU before re-reading Portfolio required marker (Lightning re-render). */
function step5PortfolioAfterBuSettleMs(): number {
  const raw = process.env.TIBBIYAH_STEP5_PORTFOLIO_AFTER_BU_MS?.trim();
  if (raw && /^\d+$/.test(raw)) return Math.max(0, Number.parseInt(raw, 10));
  return 400;
}

/**
 * Portfolio is often required on the UI only after Business Unit has a value.
 * When Excel marks Portfolio required but the asterisk is missing with an empty BU,
 * select BU and re-read the Portfolio row.
 */
async function portfolioFieldUiStateAfterBusinessUnitDependency(
  page: Page,
  modal: Locator,
  portfolioLabel: string,
  baseline: { presentOnUi: boolean; requiredOnUi: boolean },
): Promise<{ presentOnUi: boolean; requiredOnUi: boolean }> {
  let { presentOnUi, requiredOnUi } = baseline;
  if (!presentOnUi || requiredOnUi) return { presentOnUi, requiredOnUi };

  const buLabel = businessUnitFieldLabelForModal();
  const sel = await trySelectPicklistOptionByLabel(
    page,
    modal,
    comboboxAccessibleNamePattern(buLabel),
    businessUnitValueForPortfolioValidation(),
  );
  await dismissOpenPicklistListboxInLeadModal(page, modal);
  if (!sel.ok) return { presentOnUi, requiredOnUi };

  const settle = step5PortfolioAfterBuSettleMs();
  if (settle > 0) await page.waitForTimeout(settle);
  return getLeadModalFieldUiState(modal, portfolioLabel);
}

function unqualifiedReasonFieldLabelForModal(): string {
  return (
    process.env.TIBBIYAH_UNQUALIFIED_REASON_FIELD_LABEL?.trim() ||
    'Unqualified Reason'
  );
}

function unqualifiedReasonOtherPicklistValue(): string {
  return (
    process.env.TIBBIYAH_UNQUALIFIED_REASON_OTHER_VALUE?.trim() || 'Other'
  );
}

/** After choosing "Other" on Unqualified Reason, dependent fields need a tick to render. */
function step5AfterUnqualifiedReasonOtherMs(): number {
  const raw = process.env.TIBBIYAH_STEP5_UNQUALIFIED_OTHER_SETTLE_MS?.trim();
  if (raw && /^\d+$/.test(raw)) return Math.max(0, Number.parseInt(raw, 10));
  return 400;
}

/** Excel label for the free-text field that only appears when Unqualified Reason = Other. */
function isOtherUnqualifiedReasonFieldLabel(label: string): boolean {
  const t = label.replace(/\s+/g, ' ').trim();
  return /other\s+unqualified\s+(reason|readon)\b/i.test(t);
}

async function fieldUiStateAfterSelectingUnqualifiedReasonOther(
  page: Page,
  modal: Locator,
  dependentFieldLabel: string,
): Promise<{ presentOnUi: boolean; requiredOnUi: boolean }> {
  const urLabel = unqualifiedReasonFieldLabelForModal();
  await trySelectPicklistOptionByLabel(
    page,
    modal,
    comboboxAccessibleNamePattern(urLabel),
    unqualifiedReasonOtherPicklistValue(),
  );
  await dismissOpenPicklistListboxInLeadModal(page, modal);
  const ms = step5AfterUnqualifiedReasonOtherMs();
  if (ms > 0) await page.waitForTimeout(ms);
  return getLeadModalFieldUiState(modal, dependentFieldLabel);
}

function facilityDepartmentFieldLabelForModal(): string {
  return (
    process.env.TIBBIYAH_FACILITY_DEPARTMENT_FIELD_LABEL?.trim() ||
    'Facility Department'
  );
}

function facilityDepartmentOtherPicklistValue(): string {
  return (
    process.env.TIBBIYAH_FACILITY_DEPARTMENT_OTHER_VALUE?.trim() || 'Other'
  );
}

function step5AfterFacilityDepartmentOtherMs(): number {
  const raw = process.env.TIBBIYAH_STEP5_FACILITY_DEPT_OTHER_SETTLE_MS?.trim();
  if (raw && /^\d+$/.test(raw)) return Math.max(0, Number.parseInt(raw, 10));
  return 400;
}

function isOtherFacilityDepartmentFieldLabel(label: string): boolean {
  const t = label.replace(/\s+/g, ' ').trim();
  return /other\s+facility\s+department\b/i.test(t);
}

async function fieldUiStateAfterSelectingFacilityDepartmentOther(
  page: Page,
  modal: Locator,
  dependentFieldLabel: string,
): Promise<{ presentOnUi: boolean; requiredOnUi: boolean }> {
  const fdLabel = facilityDepartmentFieldLabelForModal();
  await trySelectPicklistOptionByLabel(
    page,
    modal,
    comboboxAccessibleNamePattern(fdLabel),
    facilityDepartmentOtherPicklistValue(),
  );
  await dismissOpenPicklistListboxInLeadModal(page, modal);
  const ms = step5AfterFacilityDepartmentOtherMs();
  if (ms > 0) await page.waitForTimeout(ms);
  return getLeadModalFieldUiState(modal, dependentFieldLabel);
}

/** Prints Excel vs Lead modal field/section table. Does not throw on mismatches. */
export async function reportLeadModalExcelUiComparison(
  page: Page,
  modal: Locator,
  rows: TibbiyahConfigRow[],
): Promise<void> {
  console.log('\nStep 5: Excel vs UI (fields & sections)\n');

  if (rows.length === 0) {
    console.log('(No field rows loaded from workbook.)\n');
    return;
  }

  const table: LeadFieldComparisonRow[] = [];

  for (const row of rows) {
    if (row.kind === 'section') {
      const present = await isLeadModalSectionVisible(modal, row.label);
      const status = classifySectionBug(present);
      table.push({
        Type: 'Section',
        'Excel label': row.label,
        'Required (Excel)': yn(row.requiredInExcel),
        'Present on UI': yn(present),
        'Required (UI)': '—',
        Status: status,
      });
      continue;
    }

    if (row.kind === 'field' && isPortfolioLabel(row.label)) {
      let { presentOnUi, requiredOnUi } = await getLeadModalFieldUiState(
        modal,
        row.label,
      );
      if (row.requiredInExcel && presentOnUi && !requiredOnUi) {
        ({ presentOnUi, requiredOnUi } =
          await portfolioFieldUiStateAfterBusinessUnitDependency(
            page,
            modal,
            row.label,
            { presentOnUi, requiredOnUi },
          ));
      }
      const status = classifyFieldBug(
        row.requiredInExcel,
        presentOnUi,
        requiredOnUi,
      );
      table.push({
        Type: 'Field',
        'Excel label': row.label,
        'Required (Excel)': yn(row.requiredInExcel),
        'Present on UI': yn(presentOnUi),
        'Required (UI)': naUnlessPresent(presentOnUi, requiredOnUi),
        Status: status,
      });
      continue;
    }

    if (row.kind === 'field' && isOtherUnqualifiedReasonFieldLabel(row.label)) {
      const { presentOnUi, requiredOnUi } =
        await fieldUiStateAfterSelectingUnqualifiedReasonOther(
          page,
          modal,
          row.label,
        );
      const status = classifyFieldBug(
        row.requiredInExcel,
        presentOnUi,
        requiredOnUi,
      );
      table.push({
        Type: 'Field',
        'Excel label': row.label,
        'Required (Excel)': yn(row.requiredInExcel),
        'Present on UI': yn(presentOnUi),
        'Required (UI)': naUnlessPresent(presentOnUi, requiredOnUi),
        Status: status,
      });
      continue;
    }

    if (row.kind === 'field' && isOtherFacilityDepartmentFieldLabel(row.label)) {
      const { presentOnUi, requiredOnUi } =
        await fieldUiStateAfterSelectingFacilityDepartmentOther(
          page,
          modal,
          row.label,
        );
      const status = classifyFieldBug(
        row.requiredInExcel,
        presentOnUi,
        requiredOnUi,
      );
      table.push({
        Type: 'Field',
        'Excel label': row.label,
        'Required (Excel)': yn(row.requiredInExcel),
        'Present on UI': yn(presentOnUi),
        'Required (UI)': naUnlessPresent(presentOnUi, requiredOnUi),
        Status: status,
      });
      continue;
    }

    const { presentOnUi, requiredOnUi } = await getLeadModalFieldUiState(
      modal,
      row.label,
    );
    const status = classifyFieldBug(
      row.requiredInExcel,
      presentOnUi,
      requiredOnUi,
    );
    table.push({
      Type: 'Field',
      'Excel label': row.label,
      'Required (Excel)': yn(row.requiredInExcel),
      'Present on UI': yn(presentOnUi),
      'Required (UI)': naUnlessPresent(presentOnUi, requiredOnUi),
      Status: status,
    });
  }

  logStep5LeadFieldComparisonTable(table);

  const bugs = table.filter((r) => r.Status.startsWith('BUG'));
  if (bugs.length > 0) {
    console.log(`Step 5: ${bugs.length} row(s) with BUG status (see table).\n`);
    logStructuredQaBugs(
      'Step 5 — structured bug report(s)',
      buildStep5StructuredBugs(bugs),
    );
  } else {
    console.log('Step 5: no BUG rows.\n');
  }
}

