import fs from 'fs';
import path from 'path';
import { tmpdir } from 'os';
import ExcelJS from 'exceljs';

const CREDENTIALS_DIR = path.join(process.cwd(), 'credentials');

/**
 * Default Tibbiyah Lead Config workbook (SharePoint / Excel on the web sharing link).
 * Override with `TIBBIYAH_LEAD_CONFIG_URL` or bypass remote fetch with
 * `TIBBIYAH_LEAD_CONFIG_PATH` / `TIBBIYAH_LEAD_CONFIG_LOCAL_ONLY=1`.
 */
export const DEFAULT_TIBBIYAH_LEAD_WORKBOOK_SHARE_URL =
  'https://horizontal-my.sharepoint.com/:x:/p/dpanchal/IQBVdP9QB0IrS7yLDaaU-iQaARnzZzduJdRgQtQVAjjpXRE?e=QG4Kri';

/** Prefer these names under `credentials/` (same folder as login workbook). */
const TIBBIYAH_WORKBOOK_CANDIDATES = [
  'Tibbiyah Lead Config Workbook.xlsx',
  'TibbiyahLeadConfigWorkbook.xlsx',
  'TibbiyahLeadConfig.xlsx',
];

function findTibbiyahWorkbookInCredentialsDir(): string | null {
  for (const name of TIBBIYAH_WORKBOOK_CANDIDATES) {
    const full = path.join(CREDENTIALS_DIR, name);
    if (fs.existsSync(full)) return full;
  }
  if (!fs.existsSync(CREDENTIALS_DIR)) return null;
  const files = fs.readdirSync(CREDENTIALS_DIR);
  const xlsx = files.filter(
    (f) =>
      f.toLowerCase().endsWith('.xlsx') &&
      !/^salesforcelogin/i.test(f) &&
      /tibbiyah/i.test(f),
  );
  if (xlsx.length === 1) return path.join(CREDENTIALS_DIR, xlsx[0]);
  if (xlsx.length > 1) {
    const exact = xlsx.find((f) => /lead.*config/i.test(f));
    if (exact) return path.join(CREDENTIALS_DIR, exact);
    return path.join(CREDENTIALS_DIR, xlsx[0]);
  }
  return null;
}

function resolveExplicitEnvWorkbookPathOrThrow(): string | null {
  const env = process.env.TIBBIYAH_LEAD_CONFIG_PATH?.trim();
  if (!env) return null;
  if (!fs.existsSync(env)) {
    throw new Error(
      `TIBBIYAH_LEAD_CONFIG_PATH not found: ${env}\n` +
        `Place Tibbiyah Lead Config Workbook under credentials/ or fix the path.`,
    );
  }
  return env;
}

function isZipLikeXlsxBuffer(buf: Buffer): boolean {
  return (
    buf.length >= 4 &&
    buf[0] === 0x50 &&
    buf[1] === 0x4b &&
    buf[2] === 0x03 &&
    buf[3] === 0x04
  );
}

function encodeSharingUrlForMicrosoftGraph(shareUrl: string): string {
  const b64 = Buffer.from(shareUrl, 'utf8')
    .toString('base64')
    .replace(/=+$/u, '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_');
  return `u!${b64}`;
}

async function fetchTibbiyahWorkbookBufferViaMicrosoftGraph(
  sharingUrl: string,
  accessToken: string,
): Promise<Buffer> {
  const shareId = encodeSharingUrlForMicrosoftGraph(sharingUrl);
  const metaUrl = `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem?$select=name,@microsoft.graph.downloadUrl`;
  const res = await fetch(metaUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json',
    },
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Graph shares/driveItem ${res.status}: ${text.slice(0, 400)}`);
  }
  const body = (await res.json()) as {
    '@microsoft.graph.downloadUrl'?: string;
  };
  const downloadUrl = body['@microsoft.graph.downloadUrl'];
  if (!downloadUrl || typeof downloadUrl !== 'string') {
    throw new Error('Graph response missing @microsoft.graph.downloadUrl');
  }
  const bin = await fetch(downloadUrl);
  if (!bin.ok) {
    throw new Error(`Graph pre-authenticated download failed: ${bin.status}`);
  }
  const buf = Buffer.from(await bin.arrayBuffer());
  if (!isZipLikeXlsxBuffer(buf)) {
    throw new Error('Graph download is not a valid .xlsx (ZIP) payload');
  }
  return buf;
}

async function fetchTibbiyahWorkbookBufferDirect(url: string): Promise<Buffer | null> {
  const urls = url.includes('?')
    ? [url, `${url}&download=1`]
    : [url, `${url}?download=1`];
  for (const u of urls) {
    const res = await fetch(u, {
      redirect: 'follow',
      headers: {
        Accept:
          'application/octet-stream,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*',
        'User-Agent':
          'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
      },
    });
    if (!res.ok) continue;
    const buf = Buffer.from(await res.arrayBuffer());
    if (isZipLikeXlsxBuffer(buf)) return buf;
  }
  return null;
}

function writeWorkbookBufferToTempXlsx(buf: Buffer): string {
  const dir = fs.mkdtempSync(path.join(tmpdir(), 'tibbiyah-lead-'));
  const out = path.join(dir, 'TibbiyahLeadConfigWorkbook.xlsx');
  fs.writeFileSync(out, buf);
  return out;
}

/**
 * Resolves the Tibbiyah Lead Config `.xlsx` path used by Playwright tests.
 *
 * Order:
 * 1. `TIBBIYAH_LEAD_CONFIG_PATH` when set and the file exists.
 * 2. Unless `TIBBIYAH_LEAD_CONFIG_LOCAL_ONLY=1`: download from SharePoint —
 *    prefers **Microsoft Graph** when `TIBBIYAH_LEAD_CONFIG_GRAPH_TOKEN` or
 *    `MICROSOFT_GRAPH_ACCESS_TOKEN` is set (recommended for org links), then
 *    tries a direct HTTP GET of `TIBBIYAH_LEAD_CONFIG_URL` or the default
 *    SharePoint sharing link (works when the link allows anonymous download).
 * 3. First matching workbook under `credentials/` (same as before).
 *
 * @see {@link DEFAULT_TIBBIYAH_LEAD_WORKBOOK_SHARE_URL}
 */
export async function resolveTibbiyahLeadConfigPathAsync(): Promise<string> {
  const fromEnv = resolveExplicitEnvWorkbookPathOrThrow();
  if (fromEnv) return fromEnv;

  const localOnly = process.env.TIBBIYAH_LEAD_CONFIG_LOCAL_ONLY === '1';
  if (!localOnly) {
    const url =
      process.env.TIBBIYAH_LEAD_CONFIG_URL?.trim() ||
      DEFAULT_TIBBIYAH_LEAD_WORKBOOK_SHARE_URL;
    const graphToken =
      process.env.TIBBIYAH_LEAD_CONFIG_GRAPH_TOKEN?.trim() ||
      process.env.MICROSOFT_GRAPH_ACCESS_TOKEN?.trim();

    if (graphToken) {
      try {
        const buf = await fetchTibbiyahWorkbookBufferViaMicrosoftGraph(
          url,
          graphToken,
        );
        return writeWorkbookBufferToTempXlsx(buf);
      } catch (err) {
        console.warn(
          '[Tibbiyah] Microsoft Graph workbook download failed:',
          (err as Error).message,
        );
      }
    }

    try {
      const buf = await fetchTibbiyahWorkbookBufferDirect(url);
      if (buf) return writeWorkbookBufferToTempXlsx(buf);
    } catch (err) {
      console.warn(
        '[Tibbiyah] Direct SharePoint URL fetch failed:',
        (err as Error).message,
      );
    }
  }

  const found = findTibbiyahWorkbookInCredentialsDir();
  if (found) return found;

  throw new Error(
    `Could not obtain Tibbiyah Lead Config workbook.\n` +
      `- Set TIBBIYAH_LEAD_CONFIG_GRAPH_TOKEN (or MICROSOFT_GRAPH_ACCESS_TOKEN) for Microsoft Graph access to the SharePoint link, or\n` +
      `- Place a copy under ${CREDENTIALS_DIR} (${TIBBIYAH_WORKBOOK_CANDIDATES.join(', ')}), or\n` +
      `- Set TIBBIYAH_LEAD_CONFIG_PATH to a local .xlsx file.`,
  );
}

/**
 * Tibbiyah Lead Config from `credentials/` only (no SharePoint download).
 * Prefer {@link resolveTibbiyahLeadConfigPathAsync} in tests so the SharePoint workbook is used.
 */
export function resolveTibbiyahLeadConfigPath(): string {
  const fromEnv = resolveExplicitEnvWorkbookPathOrThrow();
  if (fromEnv) return fromEnv;
  const found = findTibbiyahWorkbookInCredentialsDir();
  if (found) return found;
  throw new Error(
    `Missing Tibbiyah Lead Config workbook in ${CREDENTIALS_DIR}\n` +
      `Add one of: ${TIBBIYAH_WORKBOOK_CANDIDATES.join(', ')}\n` +
      `Or any .xlsx whose name contains "Tibbiyah", or set TIBBIYAH_LEAD_CONFIG_PATH.\n` +
      `For the SharePoint source, use resolveTibbiyahLeadConfigPathAsync() instead.`,
  );
}

/** Row containing this header starts the Picklist value section (ignored for field names). */
function isPicklistValueSectionHeader(cellText: string): boolean {
  const t = cellText.trim().toLowerCase();
  if (!t) return false;
  return (
    /^picklist\s*values?$/i.test(t) ||
    /^picklist\s*value\s*section$/i.test(t) ||
    /^picklist\s*value\b/i.test(t)
  );
}

function argbToRgb(argb: string): { r: number; g: number; b: number } | null {
  const hex = argb.replace(/^#/, '').toUpperCase();
  const body = hex.length === 8 ? hex.slice(2) : hex.length === 6 ? hex : null;
  if (!body || body.length !== 6) return null;
  const r = Number.parseInt(body.slice(0, 2), 16);
  const g = Number.parseInt(body.slice(2, 4), 16);
  const b = Number.parseInt(body.slice(4, 6), 16);
  if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) return null;
  return { r, g, b };
}

function isRedFontColor(color: Partial<ExcelJS.Color> | undefined): boolean {
  if (!color?.argb || typeof color.argb !== 'string') return false;
  const rgb = argbToRgb(color.argb);
  if (!rgb) return false;
  const { r, g, b } = rgb;
  return r >= 150 && r > g + 30 && r > b + 30;
}

function cellPlainText(cell: ExcelJS.Cell): string {
  const v = cell.value;
  if (v == null) return '';
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') {
    return String(v).trim();
  }
  if (typeof v === 'object' && v !== null && 'richText' in v) {
    const rt = (v as ExcelJS.CellRichTextValue).richText;
    return rt.map((p) => p.text).join('').trim();
  }
  if (typeof v === 'object' && v !== null && 'text' in v) {
    const t = (v as { text?: string }).text;
    return typeof t === 'string' ? t.trim() : '';
  }
  const t = cell.text;
  return typeof t === 'string' ? t.trim() : String(t ?? '').trim();
}

function cellUsesRedFont(cell: ExcelJS.Cell): boolean {
  if (isRedFontColor(cell.font?.color as Partial<ExcelJS.Color> | undefined)) {
    return true;
  }
  const v = cell.value;
  if (typeof v === 'object' && v !== null && 'richText' in v) {
    for (const part of (v as ExcelJS.CellRichTextValue).richText) {
      if (isRedFontColor(part.font?.color as Partial<ExcelJS.Color> | undefined)) {
        return true;
      }
    }
  }
  return false;
}

function cellUsesBoldFont(cell: ExcelJS.Cell): boolean {
  if (cell.font?.bold === true) return true;
  const v = cell.value;
  if (typeof v === 'object' && v !== null && 'richText' in v) {
    for (const part of (v as ExcelJS.CellRichTextValue).richText) {
      if (part.font?.bold === true && part.text.trim().length > 0) return true;
    }
  }
  return false;
}

function tokenizePicklistValueLine(line: string): string[] {
  const parts = line.split(/[,;|]/g).map((p) => p.replace(/\s+/g, ' ').trim());
  return parts.filter((p) => p.length > 0 && !/^(--\s*)?none\s*(--)?$/i.test(p));
}

/** Same token rules as Picklist Values cells (comma, semicolon, pipe); strips `--None--`. */
export function splitTibbiyahPicklistCellTokens(line: string): string[] {
  return tokenizePicklistValueLine(line);
}

function dedupePicklistValuesPreserveOrder(values: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const v of values) {
    const t = v.replace(/\s+/g, ' ').trim();
    if (!t) continue;
    const key = t.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(t);
  }
  return out;
}

function isLikelyFieldLabel(text: string): boolean {
  const t = text.trim();
  if (t.length < 2 || t.length > 200) return false;
  if (/^picklist/i.test(t)) return false;
  return true;
}

function normalizeConfigKey(text: string): string {
  return text.replace(/\s+/g, ' ').trim().toLowerCase();
}

/**
 * Lead creation modal section titles (workbook rows with these names are validated
 * as sections on the UI, not as data fields).
 */
export const LEAD_MODAL_SECTION_TITLES = [
  'Lead Information',
  'Procurement Classification',
  'Qualification & Commercial Context',
  'Company Information',
  'Communication Preference',
  'Address Information',
] as const;

const SECTION_CANONICAL_BY_NORMAL = new Map<string, string>();
for (const s of LEAD_MODAL_SECTION_TITLES) {
  SECTION_CANONICAL_BY_NORMAL.set(normalizeConfigKey(s), s);
}

/** If text matches a known section title, returns canonical section title; else null. */
export function matchLeadModalSectionTitle(text: string): string | null {
  return SECTION_CANONICAL_BY_NORMAL.get(normalizeConfigKey(text)) ?? null;
}

/** Workbook helper / header text that must not be validated as a field or section. */
function shouldIgnoreTibbiyahWorkbookCellLabel(raw: string): boolean {
  const n = normalizeConfigKey(raw);
  if (n === 'column 1' || n === 'column 2') return true;
  if (/^layout\s*:\s*lead\s*layout$/.test(n)) return true;
  return false;
}

export type TibbiyahConfigRow = {
  kind: 'field' | 'section';
  /** Field label or canonical section title. */
  label: string;
  /** True when the workbook cell uses red font (required in Excel for fields; for sections, “must appear” if you use red). */
  requiredInExcel: boolean;
  /** 1-based Excel column index this label was first read from (2 = B, 3 = C with defaults). */
  excelColumn: number;
};

/** @deprecated Use TibbiyahConfigRow; kept for readability in older call sites. */
export type TibbiyahFieldRow = TibbiyahConfigRow;

async function openTibbiyahWorksheet(
  workbookPath: string,
  sheetIndexOrName: number | string,
): Promise<{ ws: ExcelJS.Worksheet; picklistSectionStartRow: number }> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(workbookPath);
  const ws =
    typeof sheetIndexOrName === 'number'
      ? wb.worksheets[sheetIndexOrName]
      : wb.getWorksheet(sheetIndexOrName);
  if (!ws) {
    throw new Error(
      `Worksheet not found in Tibbiyah workbook: ${String(sheetIndexOrName)}`,
    );
  }

  let picklistSectionStartRow = ws.rowCount + 1;
  ws.eachRow((row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      const t = cellPlainText(cell);
      if (t && isPicklistValueSectionHeader(t)) {
        picklistSectionStartRow = Math.min(picklistSectionStartRow, rowNumber);
      }
    });
  });

  return { ws, picklistSectionStartRow };
}

/**
 * 1-based Excel columns to read (default: **2 and 3** = sheet columns **B** and **C**,
 * i.e. “Column 1” → B, “Column 2” → C).
 * Override with `TIBBIYAH_LEAD_FIELD_COLUMNS` (e.g. `2,3` or `1,2`). Legacy: `TIBBIYAH_LEAD_FIELD_COLUMN` for a single column only.
 */
function tibbiyahFieldColumnNumbers(): number[] {
  const multi = process.env.TIBBIYAH_LEAD_FIELD_COLUMNS?.trim();
  if (multi) {
    const nums = multi
      .split(/[,;\s]+/)
      .map((p) => Number.parseInt(p.trim(), 10))
      .filter((n) => Number.isFinite(n) && n >= 1);
    if (nums.length) return [...new Set(nums)];
  }
  const legacy = process.env.TIBBIYAH_LEAD_FIELD_COLUMN?.trim();
  if (legacy) {
    const n = Number.parseInt(legacy, 10);
    if (Number.isFinite(n) && n >= 1) return [n];
  }
  return [2, 3];
}

/**
 * Field and section labels from workbook columns (default **B** then **C** per row)
 * above the Picklist value section **and** above the Business Unit / Portfolio map
 * (see {@link getTibbiyahBuPortfolioMapStartRow}, default row **34**) and the Procurement
 * triple map (see {@link getTibbiyahProcurementMapStartRow}), so those cells are not
 * validated as Lead modal field labels (Steps 6–7 picklist maps handle them).
 *
 * Set `TIBBIYAH_LEAD_FIELDS_INCLUDE_BU_MAP_ROWS=1` to scan through the map rows again
 * (legacy behavior if the map starts later or you intentionally list fields there).
 * Set `TIBBIYAH_LEAD_FIELDS_INCLUDE_PROCUREMENT_MAP_ROWS=1` to include procurement map rows.
 *
 * Rows matching {@link LEAD_MODAL_SECTION_TITLES} are `kind: 'section'`; all other
 * non-empty labels are `kind: 'field'`. Required in Excel = red font on that cell.
 * Ignores `Layout : Lead Layout`, `Column 1`, `Column 2`.
 */
export async function loadLeadFieldConfigRowsFromTibbiyahWorkbook(
  workbookPath: string,
  sheetIndexOrName: number | string = 0,
): Promise<TibbiyahConfigRow[]> {
  const { ws, picklistSectionStartRow } = await openTibbiyahWorksheet(
    workbookPath,
    sheetIndexOrName,
  );

  const includeBuMapAsFields =
    process.env.TIBBIYAH_LEAD_FIELDS_INCLUDE_BU_MAP_ROWS?.trim() === '1';
  const includeProcurementMapAsFields =
    process.env.TIBBIYAH_LEAD_FIELDS_INCLUDE_PROCUREMENT_MAP_ROWS?.trim() === '1';
  const buMapFirstRow = getTibbiyahBuPortfolioMapStartRow();
  const procurementMapFirstRow = getTibbiyahProcurementMapStartRow();
  const fieldScanEndExclusive = includeBuMapAsFields
    ? picklistSectionStartRow
    : includeProcurementMapAsFields
      ? picklistSectionStartRow
      : Math.min(
          picklistSectionStartRow,
          buMapFirstRow,
          procurementMapFirstRow,
        );

  const columns = tibbiyahFieldColumnNumbers();
  const order: TibbiyahConfigRow[] = [];
  const indexByKey = new Map<string, number>();

  for (let r = 1; r < fieldScanEndExclusive; r++) {
    const excelRow = ws.getRow(r);
    for (const col of columns) {
      const cell = excelRow.getCell(col);
      const text = cellPlainText(cell);
      if (!isLikelyFieldLabel(text)) continue;
      const raw = text.replace(/\s+/g, ' ').trim();
      if (shouldIgnoreTibbiyahWorkbookCellLabel(raw)) continue;
      const sectionCanon = matchLeadModalSectionTitle(raw);
      const kind: 'field' | 'section' = sectionCanon ? 'section' : 'field';
      const label = sectionCanon ?? raw;
      const red = cellUsesRedFont(cell);
      const mergeKey = `${kind}\t${normalizeConfigKey(label)}`;
      const existing = indexByKey.get(mergeKey);
      if (existing === undefined) {
        indexByKey.set(mergeKey, order.length);
        order.push({ kind, label, requiredInExcel: red, excelColumn: col });
      } else if (red) {
        order[existing].requiredInExcel = true;
      }
    }
  }

  return order;
}

/**
 * Required field labels only (red in Excel). Does not throw if none found.
 */
export async function loadRequiredFieldLabelsFromTibbiyahWorkbook(
  workbookPath: string,
  sheetIndexOrName: number | string = 0,
): Promise<string[]> {
  const rows = await loadLeadFieldConfigRowsFromTibbiyahWorkbook(
    workbookPath,
    sheetIndexOrName,
  );
  return rows
    .filter((r) => r.kind === 'field' && r.requiredInExcel)
    .map((r) => r.label);
}

export type TibbiyahPicklistDefinition = {
  /** Bold picklist field label from the Picklist Values section. */
  fieldName: string;
  /** Non-bold values listed under that field in the workbook. */
  expectedValues: string[];
};

/**
 * Reads the **Picklist Values** region: each **bold** cell (or run) starts a new picklist
 * field name; following **non-bold** text on the same row and subsequent rows (until the
 * next bold field) are split into expected picklist values (commas, semicolons, pipes).
 */
export async function loadPicklistDefinitionsFromTibbiyahWorkbook(
  workbookPath: string,
  sheetIndexOrName: number | string = 0,
): Promise<TibbiyahPicklistDefinition[]> {
  const { ws, picklistSectionStartRow } = await openTibbiyahWorksheet(
    workbookPath,
    sheetIndexOrName,
  );

  if (picklistSectionStartRow > ws.rowCount) {
    return [];
  }

  const maxRow = ws.rowCount;
  const maxCol = Math.max(ws.columnCount || 20, 20);
  const definitions: TibbiyahPicklistDefinition[] = [];
  let current: { fieldName: string; values: string[] } | null = null;

  const flush = () => {
    if (!current) return;
    const name = current.fieldName.trim();
    if (!name) {
      current = null;
      return;
    }
    definitions.push({
      fieldName: name,
      expectedValues: dedupePicklistValuesPreserveOrder(current.values),
    });
    current = null;
  };

  for (let r = picklistSectionStartRow + 1; r <= maxRow; r++) {
    const excelRow = ws.getRow(r);
    const boldTexts: string[] = [];
    const normalTexts: string[] = [];

    for (let c = 1; c <= maxCol; c++) {
      const cell = excelRow.getCell(c);
      const t = cellPlainText(cell);
      if (!t || !t.trim()) continue;
      const normalized = t.replace(/\s+/g, ' ').trim();
      if (isPicklistValueSectionHeader(normalized)) continue;
      if (cellUsesBoldFont(cell)) boldTexts.push(normalized);
      else normalTexts.push(normalized);
    }

    if (boldTexts.length > 0) {
      flush();
      const fieldName = boldTexts.join(' ').trim();
      current = { fieldName, values: [] };
      for (const line of normalTexts) {
        for (const tok of tokenizePicklistValueLine(line)) {
          current.values.push(tok);
        }
      }
    } else if (current && normalTexts.length > 0) {
      for (const line of normalTexts) {
        for (const tok of tokenizePicklistValueLine(line)) {
          current.values.push(tok);
        }
      }
    }
  }

  flush();
  return definitions;
}

/** One row of the Business Unit ↔ Portfolio map (Excel columns B and C from the start row). */
export type BuPortfolioExcelPair = {
  /** 1-based worksheet row */
  sheetRow: number;
  businessUnit: string;
  /** Raw text from column C; may list multiple values separated by comma / semicolon / pipe. */
  portfolioRaw: string;
};

/** First worksheet row of the Business Unit (col B) / Portfolio (col C) map; same as Step 6 loader. */
export function getTibbiyahBuPortfolioMapStartRow(): number {
  const raw = process.env.TIBBIYAH_BU_PORTFOLIO_MAP_START_ROW?.trim();
  if (raw && /^\d+$/.test(raw)) {
    const n = Number.parseInt(raw, 10);
    if (n >= 1) return n;
  }
  return 34;
}

function excelColumnOneBasedFromEnv(
  envName: string,
  defaultOneBased: number,
): number {
  const raw = process.env[envName]?.trim();
  if (!raw) return defaultOneBased;
  if (/^\d+$/.test(raw)) {
    const n = Number.parseInt(raw, 10);
    return Number.isFinite(n) && n >= 1 ? n : defaultOneBased;
  }
  const letters = raw.toUpperCase().replace(/[^A-Z]/g, '');
  if (!letters) return defaultOneBased;
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n >= 1 ? n : defaultOneBased;
}

function looksLikeBuPortfolioHeaderRow(bText: string, cText: string): boolean {
  const b = bText.trim().toLowerCase();
  const c = cText.trim().toLowerCase();
  if (!b || !c) return false;
  return /business/.test(b) && /unit/.test(b) && /^portfolio/.test(c);
}

/** 1-based columns for the Procurement Sector / Channel / Request Type map (defaults B, C, D). */
export function procurementMapColumnNumbers(): {
  sector: number;
  channel: number;
  requestType: number;
} {
  return {
    sector: excelColumnOneBasedFromEnv('TIBBIYAH_PROCUREMENT_SECTOR_COL', 2),
    channel: excelColumnOneBasedFromEnv('TIBBIYAH_PROCUREMENT_CHANNEL_COL', 3),
    requestType: excelColumnOneBasedFromEnv('TIBBIYAH_PROCUREMENT_REQUEST_TYPE_COL', 4),
  };
}

function looksLikeProcurementSectorChannelHeaderRowOnly(
  bText: string,
  cText: string,
): boolean {
  const b = bText.trim().toLowerCase();
  const c = cText.trim().toLowerCase();
  if (!b || !c) return false;
  return (
    /procurement/.test(b) &&
    /sector/.test(b) &&
    /procurement/.test(c) &&
    /channel/.test(c)
  );
}

/**
 * True when BU-map columns **B/C** hold Procurement field headings, not a real Business Unit row.
 * Keeps Step 6 (BU → Portfolio) separate from the Procurement Sector / Channel / Request Type map.
 */
function isProcurementSchemaLeakingIntoBuColumns(bText: string, cText: string): boolean {
  if (looksLikeProcurementSectorChannelHeaderRowOnly(bText, cText)) return true;
  const b = bText.replace(/\s+/g, ' ').trim();
  const c = cText.replace(/\s+/g, ' ').trim();
  if (!b) return false;
  const nb = b.toLowerCase();
  const nc = c.toLowerCase();
  if (nb === 'procurement sector' && nc === 'procurement channel') return true;
  if (/^procurement\s+sector$/i.test(b) && /^procurement\s+channel$/i.test(c)) return true;
  if (/^procurement\s+sector$/i.test(b)) return true;
  if (/^procurement\s+classification$/i.test(b)) return true;
  if (/^request\s+type$/i.test(b) || /^request\s+type$/i.test(c)) return true;
  return false;
}

/**
 * True when the **Business Unit map columns** on this row look like a Procurement header row
 * (so this sheet area is procurement-only, not BU/Portfolio). Uses only `buColB` / `buColC`
 * so a combined layout (BU+Portfolio in B/C and Procurement in D–F) does not disable Step 6.
 */
function procurementHeaderBlocksBuMap(
  ws: ExcelJS.Worksheet,
  row: number,
  buColB: number,
  buColC: number,
): boolean {
  const s = cellPlainText(ws.getRow(row).getCell(buColB));
  const c = cellPlainText(ws.getRow(row).getCell(buColC));
  if (looksLikeProcurementSectorChannelHeaderRowOnly(s, c)) return true;
  const next = cellPlainText(ws.getRow(row).getCell(buColC + 1));
  if (looksLikeProcurementHeaderRow(s, c, next)) return true;
  return false;
}

/**
 * Reads the **Business Unit → Portfolio** dependency table: column **B** = Business Unit,
 * column **C** = Portfolio value(s) for that row, starting at **row 34** by default
 * (`TIBBIYAH_BU_PORTFOLIO_MAP_START_ROW`). If row 34 looks like headers (“Business Unit” /
 * “Portfolio”), data is read from the next row.
 *
 * Override columns with `TIBBIYAH_BU_PORTFOLIO_B_COL` / `TIBBIYAH_BU_PORTFOLIO_C_COL` (e.g. `B`, `C` or `2`, `3`).
 * Reading stops after **5** consecutive rows where both B and C are empty, or when the row
 * reaches the **Procurement** map start row (default **78**) if Sector/Channel use the same
 * columns as this BU map (so Public/Private/NUPCO rows are never read as Business Units).
 *
 * If the BU columns on the start row look like a **Procurement** header (not “Business Unit” /
 * “Portfolio”), this loader returns **no rows**. Rows whose B/C look like **Procurement Sector /
 * Channel** headings (misplaced above the procurement map) are **skipped** so they are not
 * treated as Business Units. Rows whose “Business Unit” cell is only `Public` / `Private` are
 * dropped (they belong in the Procurement map; default **row 78** / columns **B–D**; set
 * `TIBBIYAH_BU_ALLOW_PUBLIC_PRIVATE_AS_UNIT_NAMES=1` to keep them). Use
 * `TIBBIYAH_BU_PORTFOLIO_MAP_DISABLED=1` to skip the BU map entirely.
 */
export async function loadBuPortfolioDependencyRowsFromTibbiyahWorkbook(
  workbookPath: string,
  sheetIndexOrName: number | string = 0,
): Promise<BuPortfolioExcelPair[]> {
  const { ws } = await openTibbiyahWorksheet(workbookPath, sheetIndexOrName);
  const startRow = getTibbiyahBuPortfolioMapStartRow();
  const colB = excelColumnOneBasedFromEnv('TIBBIYAH_BU_PORTFOLIO_B_COL', 2);
  const colC = excelColumnOneBasedFromEnv('TIBBIYAH_BU_PORTFOLIO_C_COL', 3);

  if (process.env.TIBBIYAH_BU_PORTFOLIO_MAP_DISABLED?.trim() === '1') {
    return [];
  }

  if (procurementHeaderBlocksBuMap(ws, startRow, colB, colC)) {
    return [];
  }

  let r = startRow;
  const b0 = cellPlainText(ws.getRow(r).getCell(colB));
  const c0 = cellPlainText(ws.getRow(r).getCell(colC));
  if (looksLikeBuPortfolioHeaderRow(b0, c0)) r += 1;

  const out: BuPortfolioExcelPair[] = [];
  let consecutiveBlank = 0;
  const maxRow = ws.rowCount;
  const procurementMapStartRow = getTibbiyahProcurementMapStartRow();
  const { sector: procSectorCol, channel: procChannelCol } =
    procurementMapColumnNumbers();
  const procurementSharesBuColumns =
    colB === procSectorCol && colC === procChannelCol;

  for (; r <= maxRow; r++) {
    if (procurementSharesBuColumns && r >= procurementMapStartRow) {
      break;
    }
    const b = cellPlainText(ws.getRow(r).getCell(colB))
      .replace(/\s+/g, ' ')
      .trim();
    const c = cellPlainText(ws.getRow(r).getCell(colC))
      .replace(/\s+/g, ' ')
      .trim();

    if (!b && !c) {
      consecutiveBlank += 1;
      if (consecutiveBlank >= 5) break;
      continue;
    }
    consecutiveBlank = 0;
    if (!b) continue;
    if (isProcurementSchemaLeakingIntoBuColumns(b, c)) continue;

    out.push({ sheetRow: r, businessUnit: b, portfolioRaw: c });
  }

  const allowPublicPrivateBu =
    process.env.TIBBIYAH_BU_ALLOW_PUBLIC_PRIVATE_AS_UNIT_NAMES?.trim() === '1';
  if (!allowPublicPrivateBu) {
    return out.filter((row) => !/^(public|private)$/i.test(row.businessUnit.trim()));
  }

  return out;
}

/**
 * First worksheet row of the Procurement Sector / Channel / Request Type map (**header row**).
 * Tibbiyah Lead Config uses **B78** / **C78** / **D78** for the three headers by default (override
 * with `TIBBIYAH_PROCUREMENT_MAP_START_ROW` if your sheet differs).
 */
export function getTibbiyahProcurementMapStartRow(): number {
  const raw = process.env.TIBBIYAH_PROCUREMENT_MAP_START_ROW?.trim();
  if (raw && /^\d+$/.test(raw)) {
    const n = Number.parseInt(raw, 10);
    if (n >= 1) return n;
  }
  return 78;
}

function looksLikeProcurementHeaderRow(
  dText: string,
  eText: string,
  fText: string,
): boolean {
  const d = dText.trim().toLowerCase();
  const e = eText.trim().toLowerCase();
  const f = fText.trim().toLowerCase();
  if (!d || !e || !f) return false;
  return (
    /procurement/.test(d) &&
    /sector/.test(d) &&
    /procurement/.test(e) &&
    /channel/.test(e) &&
    /request/.test(f) &&
    /type/.test(f)
  );
}

/**
 * One row of the Procurement Sector → Channel → Request Type map (Excel **B, C, D** by default).
 * Blank Sector or Channel inherits the previous non-blank value on subsequent rows.
 */
export type ProcurementTripleExcelRow = {
  sheetRow: number;
  sector: string;
  channel: string;
  /** Raw text from Request Type column; may list multiple values (comma / semicolon / pipe). */
  requestTypeRaw: string;
};

/**
 * Reads **Procurement Sector → Procurement Channel → Request Type** from the workbook
 * (default header row **78** in columns **B, C, D**; override with `TIBBIYAH_PROCUREMENT_MAP_START_ROW`
 * and `TIBBIYAH_PROCUREMENT_SECTOR_COL` / `_CHANNEL_` / `_REQUEST_TYPE_`).
 *
 * On the New Lead modal, **Procurement Sector** is the controlling field; **Procurement Channel** depends
 * on Sector; **Request Type** depends on Channel. **Exception:** Sector **Private** + Channel **Direct Purchase**
 * keeps **Request Type** disabled (workbook should use `*Field Disabled*` or an empty Request Type cell).
 *
 * Stops after **5** consecutive rows where all three mapped cells are empty.
 */
export async function loadProcurementDependencyRowsFromTibbiyahWorkbook(
  workbookPath: string,
  sheetIndexOrName: number | string = 0,
): Promise<ProcurementTripleExcelRow[]> {
  const { ws } = await openTibbiyahWorksheet(workbookPath, sheetIndexOrName);
  const startRow = getTibbiyahProcurementMapStartRow();
  const { sector: colS, channel: colC, requestType: colR } =
    procurementMapColumnNumbers();

  let r = startRow;
  const d0 = cellPlainText(ws.getRow(r).getCell(colS));
  const e0 = cellPlainText(ws.getRow(r).getCell(colC));
  const f0 = cellPlainText(ws.getRow(r).getCell(colR));
  if (
    looksLikeProcurementHeaderRow(d0, e0, f0) ||
    looksLikeProcurementSectorChannelHeaderRowOnly(d0, e0)
  ) {
    r += 1;
  }

  const out: ProcurementTripleExcelRow[] = [];
  let consecutiveBlank = 0;
  const maxRow = ws.rowCount;
  let carrySector = '';
  let carryChannel = '';

  for (; r <= maxRow; r++) {
    const ds = cellPlainText(ws.getRow(r).getCell(colS))
      .replace(/\s+/g, ' ')
      .trim();
    const ch = cellPlainText(ws.getRow(r).getCell(colC))
      .replace(/\s+/g, ' ')
      .trim();
    const rt = cellPlainText(ws.getRow(r).getCell(colR))
      .replace(/\s+/g, ' ')
      .trim();

    if (!ds && !ch && !rt) {
      consecutiveBlank += 1;
      if (consecutiveBlank >= 5) break;
      continue;
    }
    consecutiveBlank = 0;

    if (ds) carrySector = ds;
    if (ch) carryChannel = ch;
    if (!carrySector || !carryChannel) continue;

    out.push({
      sheetRow: r,
      sector: carrySector,
      channel: carryChannel,
      requestTypeRaw: rt,
    });
  }

  return out;
}

export function isPortfolioLabel(label: string): boolean {
  return /^portfolio\b/i.test(label.trim());
}

export function isBusinessUnitLabel(label: string): boolean {
  return /business\s*unit/i.test(label.trim());
}
