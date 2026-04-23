/**
 * Opens Playwright Codegen at the Salesforce URL from credentials/SalesforceLogin.xlsx (row 2).
 * Only the URL cell is required; username/password may be empty while you are setting up recording.
 */
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { spawnSync } from 'child_process';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.join(__dirname, '..');
const workbookPath =
  process.env.SALESFORCE_CREDENTIALS_PATH ??
  path.join(projectRoot, 'credentials', 'SalesforceLogin.xlsx');

function normalizeHeader(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase();
}

function cellText(cell) {
  if (!cell) return '';
  const raw = cell.text;
  return typeof raw === 'string' ? raw.trim() : String(raw ?? '').trim();
}

async function loadLoginUrl() {
  if (!fs.existsSync(workbookPath)) {
    throw new Error(
      `Missing ${workbookPath}. Run "npm run credentials:ensure" or npm install.`,
    );
  }
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(workbookPath);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('Excel file has no worksheets.');

  const headerRow = sheet.getRow(1);
  const columnByHeader = new Map();
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const key = normalizeHeader(cellText(cell));
    if (key) columnByHeader.set(key, colNumber);
  });

  const urlCol =
    columnByHeader.get('url') ??
    columnByHeader.get('login url') ??
    columnByHeader.get('salesforce url');

  if (!urlCol) {
    throw new Error(
      'Row 1 must include a URL column (header: URL, Login URL, or Salesforce URL).',
    );
  }

  const url = cellText(sheet.getRow(2).getCell(urlCol));
  if (!url) {
    throw new Error(
      `Put your QA Salesforce login URL in row 2 under the URL column in ${workbookPath}.`,
    );
  }

  return url;
}

async function main() {
  const url = await loadLoginUrl();
  console.log('Starting Playwright Codegen at:', url);
  console.log('Save your recording when done. Run "npm run play:auth" to replay the login test.\n');

  const result = spawnSync('npx', ['playwright', 'codegen', url], {
    cwd: projectRoot,
    stdio: 'inherit',
    shell: true,
    env: process.env,
  });

  if (result.error) throw result.error;
  process.exit(result.status ?? 1);
}

main().catch((err) => {
  console.error(err.message || err);
  process.exit(1);
});
