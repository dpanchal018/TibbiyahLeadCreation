/**
 * Ensures credentials/SalesforceLogin.template.xlsx exists (generated once).
 * If credentials/SalesforceLogin.xlsx is missing, copies the template so you can fill row 2.
 */
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const credDir = path.join(__dirname, '..', 'credentials');
const templatePath = path.join(credDir, 'SalesforceLogin.template.xlsx');
const livePath = path.join(credDir, 'SalesforceLogin.xlsx');

async function writeTemplate() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Salesforce', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  ws.getCell('A1').value = 'URL';
  ws.getCell('B1').value = 'Username';
  ws.getCell('C1').value = 'Password';
  for (const addr of ['A1', 'B1', 'C1']) {
    ws.getCell(addr).font = { bold: true };
  }
  ws.getColumn(1).width = 58;
  ws.getColumn(2).width = 38;
  ws.getColumn(3).width = 30;
  await wb.xlsx.writeFile(templatePath);
}

async function main() {
  fs.mkdirSync(credDir, { recursive: true });
  if (!fs.existsSync(templatePath)) {
    await writeTemplate();
  }
  if (!fs.existsSync(livePath)) {
    fs.copyFileSync(templatePath, livePath);
    console.log(
      'Created credentials/SalesforceLogin.xlsx — open it and enter your Salesforce login URL, username, and password on row 2 (under the headers).',
    );
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
