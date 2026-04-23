import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';

export type SalesforceCredentials = {
  url: string;
  username: string;
  password: string;
};

const DEFAULT_WORKBOOK = path.join(
  process.cwd(),
  'credentials',
  'SalesforceLogin.xlsx',
);

function normalizeHeader(value: unknown): string {
  return String(value ?? '')
    .trim()
    .toLowerCase();
}

function cellText(cell: ExcelJS.Cell | undefined): string {
  if (!cell) return '';
  const raw = cell.text;
  return typeof raw === 'string' ? raw.trim() : String(raw ?? '').trim();
}

/**
 * Loads Salesforce login URL, username, and password from the first worksheet.
 * Row 1: headers (URL, Username, Password).
 * Row 2: values used for every test run (re-read from disk each call).
 */
export async function loadSalesforceCredentials(
  workbookPath: string = process.env.SALESFORCE_CREDENTIALS_PATH ?? DEFAULT_WORKBOOK,
): Promise<SalesforceCredentials> {
  if (!fs.existsSync(workbookPath)) {
    throw new Error(
      `Missing credential workbook: ${workbookPath}\n` +
        `Run "npm install" (which creates credentials/SalesforceLogin.xlsx from the template) or copy credentials/SalesforceLogin.template.xlsx to credentials/SalesforceLogin.xlsx.`,
    );
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(workbookPath);
  const sheet = workbook.worksheets[0];
  if (!sheet) {
    throw new Error(`No worksheets found in ${workbookPath}`);
  }

  const headerRow = sheet.getRow(1);
  const columnByHeader = new Map<string, number>();
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const key = normalizeHeader(cellText(cell));
    if (key) columnByHeader.set(key, colNumber);
  });

  const urlCol =
    columnByHeader.get('url') ??
    columnByHeader.get('login url') ??
    columnByHeader.get('salesforce url');
  const userCol =
    columnByHeader.get('username') ?? columnByHeader.get('user name');
  const passCol = columnByHeader.get('password');

  if (!urlCol || !userCol || !passCol) {
    throw new Error(
      `Row 1 in ${workbookPath} must include columns named URL, Username, and Password.`,
    );
  }

  const dataRow = sheet.getRow(2);
  const url = cellText(dataRow.getCell(urlCol));
  const username = cellText(dataRow.getCell(userCol));
  const password = cellText(dataRow.getCell(passCol));

  if (!url || !username || !password) {
    throw new Error(
      `Row 2 in ${workbookPath} must contain URL, Username, and Password (no blank cells).`,
    );
  }

  return { url, username, password };
}
