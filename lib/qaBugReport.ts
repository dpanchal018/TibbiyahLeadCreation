export type QaStructuredBug = {
  bugId: string;
  bugSummary: string;
  bugDescription: string;
  stepsToReproduce: string;
  actualResult: string;
  expectedResult: string;
};

const STEP5_STEPS = [
  '1. Sign in to Salesforce.',
  '2. Open the Leads tab and the Leads list view.',
  '3. Click New to open the New Lead modal.',
  '4. Run the automated Step 5 comparison (workbook field list vs modal).',
].join('\n');

const STEP6_STEPS = [
  '1. Sign in to Salesforce.',
  '2. Open the Leads tab and the Leads list view.',
  '3. Click New to open the New Lead modal.',
  '4. Run the automated Step 6 comparison (Business Unit / Portfolio vs workbook).',
].join('\n');

const STEP7_PROCUREMENT_STEPS = [
  '1. Sign in to Salesforce.',
  '2. Open the Leads tab and the Leads list view.',
  '3. Click New to open the New Lead modal.',
  '4. Run the automated Step 7 comparison (Procurement Sector / Channel / Request Type vs workbook B78–D78).',
].join('\n');

const PICKLIST_COMPARE_STEPS = [
  '1. Sign in to Salesforce.',
  '2. Open the Leads tab and the Leads list view.',
  '3. Click New to open the New Lead modal.',
  '4. Run the automated picklist comparison (workbook Picklist Values vs modal).',
].join('\n');

function padBugSeq(n: number): string {
  return String(n).padStart(3, '0');
}

function formatStructuredBug(b: QaStructuredBug): string {
  return [
    `Bug ID: ${b.bugId}`,
    `Bug Summary: ${b.bugSummary}`,
    `Bug Description: ${b.bugDescription}`,
    'Steps to Reproduce:',
    b.stepsToReproduce,
    'Actual Result:',
    b.actualResult,
    'Expected Result:',
    b.expectedResult,
  ].join('\n');
}

/** Logs each bug as a separate block, separated by `---`, using the requested QA format. */
export function logStructuredQaBugs(sectionHeading: string, bugs: QaStructuredBug[]): void {
  if (bugs.length === 0) return;
  console.log(`\n${sectionHeading}\n`);
  for (let i = 0; i < bugs.length; i++) {
    console.log(formatStructuredBug(bugs[i]));
    if (i < bugs.length - 1) console.log('\n---\n');
  }
  console.log('');
}

type Step5BugRow = {
  Type: string;
  'Excel label': string;
  'Required (Excel)': string;
  'Present on UI': string;
  'Required (UI)': string;
  Status: string;
};

function step5ExpectedForStatus(status: string, row: Step5BugRow): string {
  if (status.includes('Not on UI (Excel expects field)')) {
    return `The field "${row['Excel label']}" is required in the Tibbiyah Excel workbook and should appear on the New Lead modal with a recognizable label.`;
  }
  if (status.includes('Not on UI (listed in Excel)')) {
    return `The field "${row['Excel label']}" is listed in the workbook and should appear on the New Lead modal.`;
  }
  if (status.includes('Excel required; UI not required')) {
    return `With the field present on the UI, the required indicator on the modal should match Excel (Required (Excel): ${row['Required (Excel)']}).`;
  }
  if (status.includes('Excel optional; UI required')) {
    return `The required indicator on the modal should match Excel (Required (Excel): ${row['Required (Excel)']}).`;
  }
  if (status.includes('Section not on Lead modal')) {
    return `The section "${row['Excel label']}" from the workbook should appear as a section or heading on the New Lead modal.`;
  }
  return 'The New Lead modal should match the workbook definition for this row.';
}

export function buildStep5StructuredBugs(rows: Step5BugRow[]): QaStructuredBug[] {
  const bugs = rows.filter((r) => r.Status.startsWith('BUG'));
  return bugs.map((row, i) => {
    const bugId = `BUG-S5-${padBugSeq(i + 1)}`;
    const summary =
      row.Status.replace(/^BUG:\s*/i, '').trim() || 'Workbook vs New Lead modal mismatch';
    const desc = [
      `Step 5 (Excel vs UI) reported a mismatch for ${row.Type.toLowerCase()} "${row['Excel label']}".`,
      `Workbook Required (Excel): ${row['Required (Excel)']}.`,
      `Automation status: ${row.Status}.`,
    ].join(' ');
    const actual = [
      `Type: ${row.Type}`,
      `Excel label: ${row['Excel label']}`,
      `Required (Excel): ${row['Required (Excel)']}`,
      `Present on UI: ${row['Present on UI']}`,
      `Required (UI): ${row['Required (UI)']}`,
      `Status: ${row.Status}`,
    ].join('\n');
    return {
      bugId,
      bugSummary: summary,
      bugDescription: desc,
      stepsToReproduce: STEP5_STEPS,
      actualResult: actual,
      expectedResult: step5ExpectedForStatus(row.Status, row),
    };
  });
}

type Step6PicklistBugRow = {
  'Picklist field': string;
  'Depends on': string;
  'Excel values (#)': string;
  'On UI': string;
  'UI options (#)': string;
  'Missing on UI': string;
  Status: string;
};

function step6PicklistExpected(status: string, row: Step6PicklistBugRow): string {
  if (status.includes('Missing Excel BU/Portfolio source')) {
    return 'The workbook should contain either the B/C Business Unit & Portfolio map or Picklist Values for Business Unit and Portfolio so Step 6 can run.';
  }
  if (status.includes('Picklist field not on modal')) {
    return `Picklist "${row['Picklist field']}" should be visible on the New Lead modal.`;
  }
  if (status.includes('Could not read options')) {
    return `Picklist "${row['Picklist field']}" should open and list options that can be read for comparison.`;
  }
  if (status.includes('Excel value(s) not found')) {
    return `Every value listed in Excel for "${row['Picklist field']}" should appear in the modal picklist options (when ${row['Depends on']}).`;
  }
  return 'Business Unit and Portfolio picklists on the modal should match the workbook.';
}

export function buildStep6PicklistStructuredBugs(
  rows: Step6PicklistBugRow[],
  options?: {
    idPrefix?: string;
    stepsToReproduce?: string;
    descriptionPhase?: string;
    /** 1-based first sequence number for Bug ID (used when chaining Step 6 tables). */
    idStart?: number;
  },
): QaStructuredBug[] {
  const idPrefix = options?.idPrefix ?? 'S6';
  const idStart = Math.max(1, options?.idStart ?? 1);
  const steps = options?.stepsToReproduce ?? STEP6_STEPS;
  const phase = options?.descriptionPhase ?? 'Step 6 picklist';
  const bugs = rows.filter((r) => r.Status.startsWith('BUG'));
  return bugs.map((row, i) => {
    const bugId = `BUG-${idPrefix}-${padBugSeq(idStart + i)}`;
    const summary =
      row.Status.replace(/^BUG:\s*/i, '').trim() || 'Picklist validation mismatch';
    const desc = [
      `${phase} check failed for "${row['Picklist field']}".`,
      row['Depends on'] !== '—' ? `Dependency context: ${row['Depends on']}.` : '',
      `Status: ${row.Status}.`,
    ]
      .filter(Boolean)
      .join(' ');
    const actual = [
      `Picklist field: ${row['Picklist field']}`,
      `Depends on: ${row['Depends on']}`,
      `Excel values (#): ${row['Excel values (#)']}`,
      `On UI: ${row['On UI']}`,
      `UI options (#): ${row['UI options (#)']}`,
      `Missing on UI: ${row['Missing on UI']}`,
      `Status: ${row.Status}`,
    ].join('\n');
    return {
      bugId,
      bugSummary: summary,
      bugDescription: desc,
      stepsToReproduce: steps,
      actualResult: actual,
      expectedResult: step6PicklistExpected(row.Status, row),
    };
  });
}

type Step6BuPortfolioBugRow = {
  'Business Unit': string;
  'Excel Portfolio(s)': string;
  'Missing on UI': string;
  Status: string;
};

function step6BuPortfolioExpected(status: string, row: Step6BuPortfolioBugRow): string {
  if (status.includes('Could not select Business Unit')) {
    return `Business Unit "${row['Business Unit']}" from the workbook should be selectable on the New Lead modal.`;
  }
  if (status.includes('Could not read Portfolio picklist')) {
    return `After selecting Business Unit "${row['Business Unit']}", the Portfolio picklist should load and expose options for comparison.`;
  }
  if (status.includes('could not select') && status.includes('Portfolio')) {
    return `Each portfolio value listed in Excel for this BU should be selectable on the Portfolio picklist.`;
  }
  if (status.includes('Excel Portfolio value(s) not on UI')) {
    return `For Business Unit "${row['Business Unit']}", every portfolio token from Excel should appear in the Portfolio picklist on the UI.`;
  }
  return 'Excel Business Unit / Portfolio mapping should match what the New Lead modal offers for that BU.';
}

export function buildStep6BuPortfolioStructuredBugs(
  rows: Step6BuPortfolioBugRow[],
  options?: { idPrefix?: string; idStart?: number },
): QaStructuredBug[] {
  const idPrefix = options?.idPrefix ?? 'S6';
  const idStart = Math.max(1, options?.idStart ?? 1);
  const bugs = rows.filter((r) => r.Status.startsWith('BUG'));
  return bugs.map((row, i) => {
    const bugId = `BUG-${idPrefix}-${padBugSeq(idStart + i)}`;
    const summary =
      row.Status.replace(/^BUG:\s*/i, '').trim() || 'BU/Portfolio dependency mismatch';
    const desc = [
      `Step 6 (Excel B+C map) mismatch for Business Unit "${row['Business Unit']}".`,
      `Excel Portfolio(s): ${row['Excel Portfolio(s)']}.`,
      `Status: ${row.Status}.`,
    ].join(' ');
    const actual = [
      `Business Unit: ${row['Business Unit']}`,
      `Excel Portfolio(s): ${row['Excel Portfolio(s)']}`,
      `Missing on UI: ${row['Missing on UI']}`,
      `Status: ${row.Status}`,
    ].join('\n');
    return {
      bugId,
      bugSummary: summary,
      bugDescription: desc,
      stepsToReproduce: STEP6_STEPS,
      actualResult: actual,
      expectedResult: step6BuPortfolioExpected(row.Status, row),
    };
  });
}

/** Number Step 6 bugs sequentially across summary picklist rows and B+C dependency rows. */
export function buildStep6CombinedStructuredBugs(
  summaryPicklistRows: Step6PicklistBugRow[],
  buPortfolioRows: Step6BuPortfolioBugRow[],
): QaStructuredBug[] {
  const pickBugs = summaryPicklistRows.filter((r) => r.Status.startsWith('BUG'));
  const pick = buildStep6PicklistStructuredBugs(summaryPicklistRows, {
    idPrefix: 'S6',
    idStart: 1,
  });
  const dep = buildStep6BuPortfolioStructuredBugs(buPortfolioRows, {
    idPrefix: 'S6',
    idStart: pickBugs.length + 1,
  });
  return [...pick, ...dep];
}

/** Standalone picklist Excel vs modal report (not numbered as Step 6). */
export function buildPicklistModalStructuredBugs(rows: Step6PicklistBugRow[]): QaStructuredBug[] {
  return buildStep6PicklistStructuredBugs(rows, {
    idPrefix: 'PL',
    stepsToReproduce: PICKLIST_COMPARE_STEPS,
    descriptionPhase: 'Picklist (Excel vs modal)',
  });
}

export type ProcurementClassificationBugRow = {
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
  'RT picklist options (captured)': string;
  Status: string;
};

function step7ProcurementExpected(
  status: string,
  row: ProcurementClassificationBugRow,
): string {
  if (status.includes('Could not select Procurement Sector')) {
    return `Procurement Sector "${row['Excel Sector']}" from the workbook should be selectable on the New Lead modal.`;
  }
  if (status.includes('Procurement Channel not')) {
    return `After selecting Procurement Sector "${row['Excel Sector']}", Procurement Channel should be visible and enabled so the Channel from Excel can be chosen.`;
  }
  if (status.includes('Could not select Procurement Channel')) {
    return `Procurement Channel "${row['Excel Channel']}" should be selectable after the Sector from Excel is set.`;
  }
  if (status.includes('Request Type should be disabled')) {
    return `For this Sector/Channel row, Excel marks Request Type as disabled (empty or *Field Disabled*); the modal Request Type control should not be usable.`;
  }
  if (status.includes('Request Type not enabled')) {
    return `When Excel lists Request Type value(s) for this Sector/Channel, the Request Type picklist should be enabled and list those values.`;
  }
  if (status.includes('Could not read Request Type options')) {
    return `Request Type should open and list options that can be read for comparison after Channel is set.`;
  }
  if (status.includes('Excel Request Type token')) {
    return `Every Request Type token listed in Excel for this Sector/Channel should appear in the modal picklist.`;
  }
  if (status.includes('Playwright lost the UI target')) {
    return `Keep the test browser open until Step 7 completes, and allow enough time (test timeout) so Playwright does not dispose the page while the Request Type picklist is being read.`;
  }
  if (status.includes('Browser or page was closed')) {
    return `The browser tab or page was closed while Step 7 was running; rerun with the window left open until the procurement checks finish.`;
  }
  return 'Procurement Classification dependent picklists should match the workbook (B78–D78 map).';
}

export function buildProcurementClassificationStructuredBugs(
  rows: ProcurementClassificationBugRow[],
): QaStructuredBug[] {
  const bugs = rows.filter((r) => r.Status.startsWith('BUG'));
  return bugs.map((row, i) => {
    const bugId = `BUG-S7-${padBugSeq(i + 1)}`;
    const summary =
      row.Status.replace(/^BUG:\s*/i, '').trim() ||
      'Procurement Classification mismatch';
    const desc = [
      `Step 7 (Procurement Sector → Channel → Request Type) mismatch for workbook row ${row['Excel row']}.`,
      `Excel: Sector "${row['Excel Sector']}", Channel "${row['Excel Channel']}", Request Type cell "${row['Excel Request Type']}".`,
      `Status: ${row.Status}.`,
    ].join(' ');
    const actual = [
      `Excel row: ${row['Excel row']}`,
      `Excel Sector: ${row['Excel Sector']}`,
      `Excel Channel: ${row['Excel Channel']}`,
      `Excel Request Type: ${row['Excel Request Type']}`,
      `UI Sector (captured): ${row['UI Sector (captured)']}`,
      `UI Channel (captured): ${row['UI Channel (captured)']}`,
      `UI Request Type (captured): ${row['UI Request Type (captured)']}`,
      `Channel enabled: ${row['Channel enabled']}`,
      `Request Type enabled: ${row['Request Type enabled']}`,
      `Expected RT enabled: ${row['Expected RT enabled']}`,
      `Missing RT on UI: ${row['Missing RT on UI']}`,
      `RT picklist options (captured): ${row['RT picklist options (captured)']}`,
      `Status: ${row.Status}`,
    ].join('\n');
    return {
      bugId,
      bugSummary: summary,
      bugDescription: desc,
      stepsToReproduce: STEP7_PROCUREMENT_STEPS,
      actualResult: actual,
      expectedResult: step7ProcurementExpected(row.Status, row),
    };
  });
}
