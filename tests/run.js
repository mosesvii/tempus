'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { MockSpreadsheet, createAppsScriptContext } = require('./mock-apps-script');

const root = path.resolve(__dirname, '..');
let passed = 0;
function test(name, callback) {
  try { callback(); passed++; process.stdout.write('ok - ' + name + '\n'); }
  catch (error) { process.stderr.write('not ok - ' + name + '\n' + error.stack + '\n'); process.exitCode = 1; }
}
function load(context, file) { vm.runInNewContext(fs.readFileSync(path.join(root, file), 'utf8'), context, { filename: file }); }
function fixture(name) { return JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', name + '.json'), 'utf8')); }
function extractClientFunction(source, name) {
  const match = source.match(new RegExp('  function ' + name + '\\([\\s\\S]*?\\n  \\}'));
  assert.ok(match, 'Expected client function ' + name);
  return match[0].trimStart();
}

test('all six migration fixture families are present', () => {
  ['fork-only', 'original-only', 'hybrid', 'empty', 'malformed-header', 'partially-migrated'].forEach((name) => assert.ok(fixture(name).sheets));
});

test('canonical migration remaps aliases by name and preserves unknown named columns', () => {
  const data = fixture('original-only').sheets;
  const { context, spreadsheet } = createAppsScriptContext(data);
  context.assertMigrationsSettled_ = () => true;
  load(context, 'backend/sheets.ts.js');
  const beforeRows = spreadsheet.getSheetByName('timesheet_entries').getLastRow();
  context.canonicalizeSheet_('timesheet_entries');
  const once = spreadsheet.getSheetByName('timesheet_entries').snapshot();
  assert.ok(once[0].includes('contract_id'));
  assert.ok(once[0].includes('legacy_marker'));
  assert.equal(once[1][once[0].indexOf('contract_id')], 'contract-1');
  assert.equal(spreadsheet.getSheetByName('timesheet_entries').getLastRow(), beforeRows);
  context.canonicalizeSheet_('timesheet_entries');
  assert.deepStrictEqual(spreadsheet.getSheetByName('timesheet_entries').snapshot(), once);
});

test('malformed populated header fails loudly', () => {
  const { context } = createAppsScriptContext(fixture('malformed-header').sheets);
  context.assertMigrationsSettled_ = () => true;
  load(context, 'backend/sheets.ts.js');
  assert.throws(() => context.canonicalizeSheet_('timesheet_entries'), /data but no header/);
});

test('canonical migration losslessly coalesces duplicate legacy columns', () => {
  const { context, spreadsheet } = createAppsScriptContext({
    hour_types: [
      ['id', 'name', 'icon', 'quick_fill_enabled', 'icon'],
      ['work', 'Work', '⏱️', 'TRUE', ''],
      ['leave', 'Leave', '', 'TRUE', '🌴'],
      ['sick', 'Sick', '🤒', 'TRUE', '🤒']
    ]
  });
  context.assertMigrationsSettled_ = () => true;
  load(context, 'backend/sheets.ts.js');
  const result = context.canonicalizeSheet_('hour_types');
  const output = spreadsheet.getSheetByName('hour_types').snapshot();
  const iconIndex = output[0].indexOf('icon');
  assert.deepStrictEqual(Array.from(result.coalesced_headers), ['icon']);
  assert.equal(output[0].filter((header) => header === 'icon').length, 1);
  assert.deepStrictEqual([output[1][iconIndex], output[2][iconIndex], output[3][iconIndex]], ['⏱️', '🌴', '🤒']);
});

test('canonical migration rejects conflicting duplicate legacy values', () => {
  const { context } = createAppsScriptContext({
    hour_types: [['id', 'icon', 'icon'], ['work', '⏱️', '💼']]
  });
  context.assertMigrationsSettled_ = () => true;
  load(context, 'backend/sheets.ts.js');
  assert.throws(() => context.canonicalizeSheet_('hour_types'), /duplicate header "icon" has conflicting values in row 2/);
});

const CANONICAL_HOUR_TYPE_HEADERS = ['id', 'name', 'slug', 'color', 'contributes_to_income', 'requires_contract', 'is_default', 'use_for_rate_calculation', 'auto_populate_public_holidays', 'auto_populate_hours', 'entry_mode', 'created_at', 'display_order', 'quick_fill_enabled', 'quick_fill_hours', 'icon', 'quick_fill_mode', 'counts_as_leave'];

// A copy of a production sheet as it stood before the quick-fill release: no use_for_rate_calculation,
// no entry_mode, no quick-fill columns, and contracts still on its original six columns.
function legacyProductionSheets() {
  return {
    hour_types: [
      ['id', 'name', 'slug', 'color', 'contributes_to_income', 'requires_contract', 'is_default', 'auto_populate_public_holidays', 'auto_populate_hours', 'created_at', 'display_order'],
      ['work-id', 'Work', 'work', '#3b82f6', 'TRUE', 'TRUE', 'TRUE', 'FALSE', 0, '2024-02-01T00:00:00Z', 1],
      ['leave-id', 'Annual Leave', 'annual-leave', '#22c55e', 'FALSE', 'FALSE', 'FALSE', 'TRUE', 7.5, '2024-02-01T00:00:00Z', 2]
    ],
    contracts: [
      ['id', 'name', 'start_date', 'end_date', 'hourly_rate', 'created_at'],
      ['contract-1', 'Acme', '2024-02-01', '', 150, '2024-02-01T00:00:00Z']
    ]
  };
}

// Row 1 as actually stored, not as snapshot() reports it: snapshot is clipped to the sheet's data
// region, which is exactly the window a duplicated column can hide beyond.
function sheetHeaderRow(spreadsheet, name) {
  const headers = (spreadsheet.getSheetByName(name).values[0] || []).map((value) => String(value == null ? '' : value));
  while (headers.length && !headers[headers.length - 1]) headers.pop();
  return headers;
}

function sheetHeaderDuplicates(spreadsheet, name) {
  const counts = {};
  sheetHeaderRow(spreadsheet, name).forEach((header) => {
    if (header) counts[header] = (counts[header] || 0) + 1;
  });
  return Object.keys(counts).filter((header) => counts[header] > 1);
}

// Stubs for the collaborators the sheet-backed APIs reach for, so a test can exercise the real read path.
function withSheetApiStubs(context) {
  context.cacheGet = () => null;
  context.cacheSet = () => {};
  context.cacheClearPrefix = () => {};
  context.withScriptLock_ = (label, callback) => callback();
  context.toIsoDate = (value) => String(value || '').slice(0, 10);
  context.toIsoDateTime = (value) => new Date(value || Date.now()).toISOString();
  context.api_getSettings = () => ({});
  context.api_updateSettings = () => {};
  context.api_getEntryDefaults = () => ({ basic: [] });
  return context;
}

test('a fresh fork of a production sheet upgrades and then reads back without duplicating columns', () => {
  const sheets = legacyProductionSheets();
  const { context, spreadsheet, properties } = createAppsScriptContext(sheets);
  const backup = new MockSpreadsheet(legacyProductionSheets());
  properties.setProperty('tempus_migrations_applied', '[]');
  context.SpreadsheetApp.openById = () => backup;
  context.DriveApp = {
    getRootFolder: () => ({ id: 'root' }),
    getFileById: () => ({
      getParents: () => ({ hasNext: () => false }),
      makeCopy: () => ({ getId: () => 'verified-backup', getUrl: () => 'https://drive.google.com/verified-backup', getName: () => 'Tempus pre-upgrade backup' })
    })
  };
  withSheetApiStubs(context);
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/migrations.js');
  load(context, 'backend/hourtypes.js');
  load(context, 'backend/contracts.js');

  ['backend/entries.js','backend/integrity.js','backend/settings.js','backend/invoices.js','backend/invoiceLedger.js','backend/assessments.js','backend/assessmentInvoices.js'].forEach(file=>load(context,file));
  const upgrade = context.api_runUpgrade();
  assert.equal(upgrade.required, false, 'every migration should apply on a fresh fork');
  assert.equal(upgrade.state, 'complete');

  // The upgrade gate is left in its real state: reads below prove it opened rather than being stubbed.
  assert.deepStrictEqual(sheetHeaderRow(spreadsheet, 'hour_types'), CANONICAL_HOUR_TYPE_HEADERS);

  // The read path calls getHourTypesSheet repeatedly inside one execution (api_getHourTypes, then
  // ensureWorkHourType, then again per lookup) — the shape that produced the duplicate columns.
  context.getHourTypesSheet();
  context.getHourTypesSheet();
  const hourTypes = context.api_getHourTypes();
  context.api_getHourTypes();
  context.getDefaultHourTypeId();
  context.api_getContracts();

  assert.deepStrictEqual(sheetHeaderDuplicates(spreadsheet, 'hour_types'), []);
  assert.deepStrictEqual(sheetHeaderDuplicates(spreadsheet, 'contracts'), []);
  assert.deepStrictEqual(sheetHeaderRow(spreadsheet, 'hour_types'), CANONICAL_HOUR_TYPE_HEADERS);

  // Legacy data survives, and the columns the fork added arrive with their declared defaults.
  const work = hourTypes.find((type) => type.id === 'work-id');
  const leave = hourTypes.find((type) => type.id === 'leave-id');
  assert.equal(hourTypes.length, 2, 'no phantom hour type is created on top of the legacy rows');
  assert.equal(work.name, 'Work');
  assert.equal(work.contributes_to_income, true);
  assert.equal(work.display_order, 1);
  assert.equal(work.quick_fill_mode, 'hours');
  assert.equal(leave.auto_populate_public_holidays, true);
  assert.equal(leave.auto_populate_hours, 7.5);
  assert.equal(leave.quick_fill_enabled, false);
  const contract = context.api_getContracts()[0];
  assert.equal(contract.name, 'Acme');
  assert.equal(contract.hourly_rate, 150);
  assert.equal(contract.standard_hours_per_day, 7.5);
  assert.equal(contract.include_weekends, false);
});

test('a brand new spreadsheet bootstraps hour_types complete and stays stable across reads', () => {
  const { context, spreadsheet } = createAppsScriptContext({});
  context.assertMigrationsSettled_ = () => true;
  withSheetApiStubs(context);
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/hourtypes.js');
  context.getHourTypesSheet();
  assert.deepStrictEqual(sheetHeaderRow(spreadsheet, 'hour_types'), CANONICAL_HOUR_TYPE_HEADERS);
  const created = context.api_getHourTypes();
  assert.equal(created.length, 1);
  assert.equal(created[0].slug, 'work');
  context.api_getHourTypes();
  context.getHourTypesSheet();
  assert.deepStrictEqual(sheetHeaderDuplicates(spreadsheet, 'hour_types'), []);
  assert.equal(context.api_getHourTypes().length, 1, 'repeat reads must not append another Work row');
});

test('a column holding only a header is never re-appended when the data region under-reports it', () => {
  const { context, spreadsheet } = createAppsScriptContext({
    hour_types: [CANONICAL_HOUR_TYPE_HEADERS, ['work-id', 'Work', 'work', '#3b82f6', 'TRUE', 'TRUE', 'TRUE', 'TRUE', 'FALSE', 0, '', '2024-02-01T00:00:00Z', 1, 'FALSE', '', '', '']]
  });
  context.assertMigrationsSettled_ = () => true;
  withSheetApiStubs(context);
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/hourtypes.js');
  // quick_fill_hours and icon are seeded blank, so they carry a header and nothing else. Model a data
  // region that only reaches columns with body content — the state in which the old fix-up judged those
  // columns missing and appended a second copy of each.
  const sheet = spreadsheet.getSheetByName('hour_types');
  sheet.getLastColumn = function() {
    let last = 0;
    this.values.slice(1).forEach((row) => row.forEach((value, index) => { if (value !== '' && value != null) last = Math.max(last, index + 1); }));
    return last;
  };
  assert.ok(sheet.getLastColumn() < CANONICAL_HOUR_TYPE_HEADERS.length, 'the window must fall short for this test to mean anything');
  context.getHourTypesSheet();
  context.getHourTypesSheet();
  context.api_getHourTypes();
  assert.deepStrictEqual(sheetHeaderDuplicates(spreadsheet, 'hour_types'), []);
  assert.deepStrictEqual(sheetHeaderRow(spreadsheet, 'hour_types'), CANONICAL_HOUR_TYPE_HEADERS);
});

test('upgraded hour types read lowercase text booleans as on (quick actions and leave survive the upgrade)', () => {
  const { context } = createAppsScriptContext({
    hour_types: [
      CANONICAL_HOUR_TYPE_HEADERS,
      ['leave-id', 'Annual Leave', 'annual-leave', '#0ea5e9', 'FALSE', 'FALSE', 'FALSE', 'FALSE', 'FALSE', 0, '', '2026-01-22T23:51:58Z', 1, 'true', 7.5, 'plane', 'hours', 'true'],
      ['work-id', 'Work', 'work', '#3b82f6', 'TRUE', 'TRUE', 'true', 'TRUE', 'FALSE', 0, '', '2025-11-02T23:42:37Z', 0, 'true', 8.5, 'briefcase', 'hours', 'false'],
      ['home-id', 'Home billable', 'home', '#06b6d4', 'TRUE', 'TRUE', 'FALSE', 'FALSE', 'FALSE', 0, '', '2026-07-06T12:49:37Z', 3, 'false', '', 'clock', 'hours', 'false']
    ]
  });
  withSheetApiStubs(context);
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/migrations.js');
  load(context, 'backend/hourtypes.js');
  context.assertMigrationsSettled_ = () => true;
  const byName = Object.fromEntries(context.api_getHourTypes().map((ht) => [ht.name, ht]));
  assert.equal(byName['Annual Leave'].quick_fill_enabled, true);
  assert.equal(byName['Annual Leave'].counts_as_leave, true);
  assert.equal(byName['Home billable'].quick_fill_enabled, false);
  assert.equal(byName['Home billable'].counts_as_leave, false);
  assert.equal(byName.Work.is_default, true, 'a lowercase is_default still marks the default type');
});

test('a sheet already carrying duplicate quick-fill columns is repaired by the upgrade', () => {
  const damaged = CANONICAL_HOUR_TYPE_HEADERS.concat(['quick_fill_hours', 'icon']);
  const { context, spreadsheet } = createAppsScriptContext({
    hour_types: [
      damaged,
      ['work-id', 'Work', 'work', '#3b82f6', 'TRUE', 'TRUE', 'TRUE', 'TRUE', 'FALSE', 0, '', '2024-02-01T00:00:00Z', 1, 'TRUE', '', '', 'hours', 'FALSE', 7.5, 'clock']
    ]
  });
  withSheetApiStubs(context);
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/migrations.js');
  load(context, 'backend/hourtypes.js');
  context.assertMigrationsSettled_ = () => true; // the damaged sheet's own upgrade gate is not what is under test
  assert.throws(() => context.getHourTypesSheet(), /duplicate header "quick_fill_hours"/);
  context.migrationCanonicalQuickFillColumns_();
  assert.deepStrictEqual(sheetHeaderRow(spreadsheet, 'hour_types'), CANONICAL_HOUR_TYPE_HEADERS);
  const repaired = context.api_getHourTypes();
  assert.equal(repaired.length, 1);
  assert.equal(repaired[0].quick_fill_hours, 7.5, 'the populated copy of the duplicated column wins');
  assert.equal(repaired[0].icon, 'clock');
});

test('the repair also removes a duplicate column that holds nothing but a header', () => {
  const { context, spreadsheet } = createAppsScriptContext({
    hour_types: [
      CANONICAL_HOUR_TYPE_HEADERS.concat(['quick_fill_hours', 'icon']),
      ['work-id', 'Work', 'work', '#3b82f6', 'TRUE', 'TRUE', 'TRUE', 'TRUE', 'FALSE', 0, '', '2024-02-01T00:00:00Z', 1, 'TRUE', 7.5, 'clock', 'hours', '', '']
    ]
  });
  withSheetApiStubs(context);
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/migrations.js');
  load(context, 'backend/hourtypes.js');
  context.assertMigrationsSettled_ = () => true;
  context.migrationCanonicalQuickFillColumns_();
  assert.deepStrictEqual(sheetHeaderRow(spreadsheet, 'hour_types'), CANONICAL_HOUR_TYPE_HEADERS);
  assert.deepStrictEqual(sheetHeaderDuplicates(spreadsheet, 'hour_types'), []);
  const repaired = context.api_getHourTypes();
  assert.equal(repaired[0].quick_fill_hours, 7.5, 'the populated copy survives the empty duplicate');
  assert.equal(repaired[0].icon, 'clock');
});

test('canonical sheet reads do not rewrite whole-column number formats', () => {
  const { context, spreadsheet } = createAppsScriptContext({});
  context.assertMigrationsSettled_ = () => true;
  load(context, 'backend/sheets.ts.js');
  const sheet = spreadsheet.insertSheet('timesheet_entries');
  sheet.getRange(1, 1, 1, context.TEMPUS_SHEET_SCHEMAS.timesheet_entries.headers.length)
    .setValues([context.TEMPUS_SHEET_SCHEMAS.timesheet_entries.headers]);
  const originalGetRange = sheet.getRange.bind(sheet);
  let formatWrites = 0;
  sheet.getRange = function() {
    const range = originalGetRange.apply(sheet, arguments);
    const originalSetNumberFormat = range.setNumberFormat.bind(range);
    range.setNumberFormat = function() {
      formatWrites++;
      return originalSetNumberFormat.apply(range, arguments);
    };
    return range;
  };
  context.getOrCreateSheet('timesheet_entries');
  assert.equal(formatWrites, 0);
});

test('upgrade verifies a backup, records failure, and resumes without rerunning completed work', () => {
  const { context, spreadsheet, properties } = createAppsScriptContext({});
  const backup = new MockSpreadsheet({});
  properties.setProperty('tempus_migrations_applied', '[]');
  context.SpreadsheetApp.openById = () => backup;
  context.DriveApp = {
    getRootFolder: () => ({ id: 'root' }),
    getFileById: () => ({
      getParents: () => ({ hasNext: () => false }),
      makeCopy: () => ({ getId: () => 'verified-backup', getUrl: () => 'https://drive.google.com/verified-backup', getName: () => 'Tempus pre-upgrade backup' })
    })
  };
  context.assertMigrationsSettled_ = () => true;
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/migrations.js');
  let fail = true;
  let firstRuns = 0;
  let secondRuns = 0;
  context.listMigrations = () => [
    { id: 'one', run: () => { firstRuns++; if (fail) { fail = false; throw new Error('forced interruption'); } } },
    { id: 'two', run: () => { secondRuns++; } }
  ];
  const interrupted = context.api_runUpgrade();
  assert.equal(interrupted.success, false);
  assert.equal(interrupted.status.state, 'error');
  assert.equal(JSON.parse(properties.getProperty('tempus_migration_backup_v3')).verified, true);
  const resumed = context.api_runUpgrade();
  assert.equal(resumed.required, false);
  assert.equal(firstRuns, 2);
  assert.equal(secondRuns, 1);
  context.api_runUpgrade();
  assert.equal(firstRuns, 2);
  assert.equal(secondRuns, 1);
  assert.equal(spreadsheet.getSheets().length, 0);
});

test('PAYG tables switch at the Australian financial-year boundary and warn when stale', () => {
  const { context } = createAppsScriptContext({});
  load(context, 'backend/tax.js');
  assert.equal(context.paygFinancialYear_('2026-06-30T12:00:00+10:00'), '2025-26');
  assert.equal(context.paygFinancialYear_('2026-07-01T12:00:00+10:00'), '2026-27');
  assert.equal(context.estimateTaxDetails(10000, '2026-08-01').stale, false);
  assert.equal(context.estimateTaxDetails(10000, '2027-08-01').stale, true);
  assert.ok(context.estimateTax(10000, '2026-08-01') > 0);
});

test('recurring fortnight buckets are based on civil dates across Sydney DST', () => {
  const { context } = createAppsScriptContext({});
  context.toIsoDate = (value) => typeof value === 'string' ? value : value.toISOString().slice(0, 10);
  context.toIsoDateTime = context.toIsoDate;
  context.normalizeDurationMinutes = Number;
  context.punchesTotalMinutes = () => 0;
  load(context, 'backend/recurringEntries.js');
  const schedule = { weekly_weekdays: [0], weekly_interval: 2 };
  assert.equal(context.matchesWeeklySchedule(schedule, new Date(2026, 9, 4), new Date(2026, 8, 27)), false);
  assert.equal(context.matchesWeeklySchedule(schedule, new Date(2026, 9, 11), new Date(2026, 8, 27)), true);
});

test('month-end recurrence keeps its original anchor after February', () => {
  const { context } = createAppsScriptContext({});
  load(context, 'backend/annualviews.js');
  const anchor = new Date(2024, 0, 31);
  const february = context.advanceDateByFrequency(anchor, 'monthly', anchor);
  const march = context.advanceDateByFrequency(february, 'monthly', anchor);
  assert.equal(february.getDate(), 29);
  assert.equal(march.getDate(), 31);
});

test('percentage, URL and CSV protections normalize once', () => {
  const { context } = createAppsScriptContext({});
  context.Utilities.computeDigest = () => [0];
  load(context, 'backend/integrity.js');
  assert.equal(context.normalizePercentageDecimal_(25), 0.25);
  assert.equal(context.normalizePercentageDecimal_(0.25), 0.25);
  assert.throws(() => context.requireHttpUrl_('javascript:alert(1)', 'URL', false));
  assert.ok(context.csvSafeCell_('=SUM(A1:A2)').startsWith('"\''));
});

test('importer validates calendar dates and represents full days without fake punches', () => {
  const { context } = createAppsScriptContext({});
  context.ENTRY_SHEET_TZ = 'Australia/Sydney';
  load(context, 'backend/timesheetImporter.js');
  const leap = context.normalizeSheetDate('29/02/2024');
  assert.equal(leap.getFullYear(), 2024);
  assert.equal(context.normalizeSheetDate('31/02/2024'), null);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(context.buildPunchesForDuration(1440))), []);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(context.buildPunchesForDuration(90))), [{ in: '00:00', out: '01:30' }]);
});

test('partial invoice payments allocate cash-basis GST proportionally', () => {
  const { context } = createAppsScriptContext({});
  context.cacheGet = () => null;
  context.cacheSet = () => {};
  context.getOrCreateSheet = () => ({ getDataRange: () => ({ getValues: () => [[]] }) });
  context.rowObjectFromHeaders_ = () => ({});
  context.roundMoney_ = (value) => Math.round((Number(value) + Number.EPSILON) * 100) / 100;
  context.listInvoiceLineItemsByInvoiceId = () => [{ amount: 100, gst_amount: 10 }];
  context.summarizeInvoiceLineItems = () => ({ totalWithGst: 110, gstAmount: 10 });
  load(context, 'backend/invoiceLedger.js');
  context.invoicePaymentsForInvoice_ = () => [{ payment_date: '2026-07-15', amount: 55 }, { payment_date: '2026-08-15', amount: 55 }];
  const allocation = context.invoiceCashAllocationForPeriod_({ id: 'invoice-1' }, '2026-07-01', '2026-07-31');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(allocation)), { sales: 55, gst: 5 });
});

test('BAS fixtures separate cash, accrual, prepayment, void, and unreconciled GST', () => {
  const data = fixture('financial-cases');
  const { context } = createAppsScriptContext({});
  context.cacheGet = () => null; context.cacheSet = () => {}; context.cacheClearPrefix = () => {};
  context.rowObjectFromHeaders_ = () => ({});
  context.getOrCreateSheet = () => ({ getDataRange: () => ({ getValues: () => [[]] }) });
  context.roundMoney_ = (value) => Math.round((Number(value) + Number.EPSILON) * 100) / 100;
  load(context, 'backend/invoiceLedger.js');
  load(context, 'backend/financialReporting.js');
  context.listInvoicesInternal = () => data.invoices;
  context.listInvoiceLineItemsByInvoiceId = (id) => data.invoice_lines[id] || [];
  context.listInvoiceLineItemsInternal = () => [];
  context.summarizeInvoiceLineItems = (lines) => ({ totalWithGst: lines.reduce((sum, line) => sum + line.amount + line.gst_amount, 0), gstAmount: lines.reduce((sum, line) => sum + line.gst_amount, 0) });
  context.invoicePayments_ = () => data.invoice_payments;
  context.invoicePaymentsForInvoice_ = (id) => data.invoice_payments.filter((payment) => payment.invoice_id === id).sort((a, b) => a.payment_date.localeCompare(b.payment_date));
  context.api_listExpenseTransactions = () => data.expenses;
  context.expenseReadSheet_ = (name) => ({ rows: name === 'expense_payments' ? data.expense_payments : [] });
  context.expenseClaimableGst_ = (expense) => expense.status !== 'void' && expense.reconciliation_state === 'reconciled' && expense.claimable_gst_confirmed === 'TRUE' && expense.gst_code === 'taxable' ? expense.gst_amount * expense.business_use_percentage : 0;
  context.api_listExpenseRules = () => [];
  context.api_listInvoices = () => [{ id:'invoice-part', status:'issued', invoice_date:'2026-07-10', payment_state:'part_paid', balance_due:55 }];
  context.timesheetForecastForPeriod_ = () => ({ income: 0, uninvoiced_time: 0 });
  context.scheduledExpenseForecastForPeriod_ = () => 0;
  context.sha256Hex_ = () => 'fixture-hash';
  context.stableJsonStringify_ = JSON.stringify;
  const cash = context.api_calculateBasPeriod({ financial_year:2026, period_type:'monthly', month:6, accounting_basis:'cash' });
  const accrual = context.api_calculateBasPeriod({ financial_year:2026, period_type:'monthly', month:6, accounting_basis:'accrual' });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(cash.actual)), { g1_total_sales:275, gst_on_sales:25, purchases:105, gst_on_purchases:5 });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(accrual.actual)), { g1_total_sales:330, gst_on_sales:30, purchases:215, gst_on_purchases:10 });
  assert.ok(accrual.warnings.some((warning) => warning.includes('excluded from GST credits')));
});

test('income summaries wait for complete finance data and invalidate on reference-data changes', () => {
  const scripts = fs.readFileSync(path.join(root, 'views/partials/scripts.html'), 'utf8');
  assert.match(scripts, /if \(!incomeDependenciesReady\(\)\) \{\s*state\.incomeSummary = null;/);
  assert.match(scripts, /function processMonthCalculationQueue\(\) \{\s*if \(!state\.lazyLoadQueue\) return;\s*if \(!incomeDependenciesReady\(\)\) return;/);
  assert.match(scripts, /const isDirty = shouldRecalculateMonth\(year, month\);/);
  assert.match(scripts, /metadata\.dependencyHash !== dependencyHash/);
  assert.match(scripts, /state\.incomeCacheMetadata\[monthKey\]\.dependencyHash = calculateIncomeDependencyHash\(\);/);
  ['contributes_to_income', 'hourly_rate', 'amount_value', 'lost_super_recovery_mode', 'percentage'].forEach((dependency) => {
    assert.ok(scripts.includes(dependency), 'missing income dependency: ' + dependency);
  });
  assert.match(scripts, /const entries = entriesForMonth\(year, month\)\.filter\(entryContributesToIncome\);/);
  assert.match(scripts, /return !getFeatureFlag\('hour_types'\);/);
  assert.match(scripts, /String\(candidate\.slug \|\| ''\)\.trim\(\)\.toLowerCase\(\) === legacySlug/);
  assert.equal((scripts.match(/function entryContributesToIncome\(entry\)/g) || []).length, 1);
  const helperStart = scripts.indexOf('function hourTypeContributesToIncome(');
  const helperEnd = scripts.indexOf('\nfunction ensureIncomeCacheStructures()', helperStart);
  const clientContext = {
    state: {
      hourTypeMap: {
        'work-id': { id: 'work-id', slug: 'work', contributes_to_income: true },
        'leave-id': { id: 'leave-id', slug: 'leave', contributes_to_income: false }
      },
      hourTypes: [
        { id: 'work-id', slug: 'work', contributes_to_income: true },
        { id: 'leave-id', slug: 'leave', contributes_to_income: false }
      ]
    },
    getDefaultHourTypeId: () => 'work-id',
    getFeatureFlag: () => true
  };
  vm.createContext(clientContext);
  vm.runInContext(scripts.slice(helperStart, helperEnd), clientContext);
  assert.equal(clientContext.hourTypeContributesToIncome('work'), true);
  assert.equal(clientContext.hourTypeContributesToIncome('leave-id'), false);
  assert.equal(clientContext.hourTypeContributesToIncome('unknown-type'), false);
});

test('legacy Work hour types are repaired to billable without overriding an explicit default', () => {
  const headers = ['id', 'name', 'slug', 'color', 'contributes_to_income', 'requires_contract', 'is_default', 'use_for_rate_calculation', 'auto_populate_public_holidays', 'auto_populate_hours', 'created_at', 'display_order', 'quick_fill_enabled', 'quick_fill_hours', 'icon', 'quick_fill_mode'];
  const { context, spreadsheet } = createAppsScriptContext({
    hour_types: [
      headers,
      ['work-id', 'Work', 'work', '#3b82f6', 'FALSE', 'FALSE', 'FALSE', 'FALSE', 'FALSE', 0, '', 1, 'TRUE', 7.5, 'clock', 'hours'],
      ['wfh-id', 'WFH', 'wfh', '#10b981', 'TRUE', 'TRUE', 'TRUE', 'TRUE', 'FALSE', 0, '', 2, 'FALSE', '', '', 'hours']
    ]
  });
  context.assertMigrationsSettled_ = () => true;
  context.cacheClearPrefix = () => {};
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/hourtypes.js');
  const work = context.ensureWorkHourType();
  assert.equal(work.contributes_to_income, true);
  assert.equal(work.requires_contract, true);
  assert.equal(work.is_default, false);
  const rows = spreadsheet.getSheetByName('hour_types').snapshot();
  const workRow = rows.find((row) => row[headers.indexOf('id')] === 'work-id');
  assert.equal(workRow[headers.indexOf('contributes_to_income')], 'TRUE');
  assert.equal(workRow[headers.indexOf('requires_contract')], 'TRUE');
  assert.equal(workRow[headers.indexOf('is_default')], 'FALSE');
});

test('hour type entry modes persist and invalid values fall back to the global default', () => {
  const headers = ['id', 'name', 'slug', 'color', 'contributes_to_income', 'requires_contract', 'is_default', 'use_for_rate_calculation', 'auto_populate_public_holidays', 'auto_populate_hours', 'entry_mode', 'created_at', 'display_order', 'quick_fill_enabled', 'quick_fill_hours', 'icon', 'quick_fill_mode'];
  const { context, spreadsheet } = createAppsScriptContext({
    hour_types: [
      headers,
      ['work-id', 'Work', 'work', '#3b82f6', 'TRUE', 'TRUE', 'TRUE', 'TRUE', 'FALSE', 0, '', '', 1, 'TRUE', 7.5, 'clock', 'hours']
    ]
  });
  context.assertMigrationsSettled_ = () => true;
  context.cacheClearPrefix = () => {};
  context.withScriptLock_ = (label, callback) => callback();
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/hourtypes.js');
  assert.equal(context.api_updateHourType('work-id', { entry_mode: 'detailed' }).entry_mode, 'detailed');
  assert.equal(context.api_updateHourType('work-id', { entry_mode: 'unexpected' }).entry_mode, '');
  const rows = spreadsheet.getSheetByName('hour_types').snapshot();
  assert.equal(rows[1][rows[0].indexOf('entry_mode')], '');
});

test('Annual View uses taxable income and the same FY2026 PAYG result as the server', () => {
  const scripts = fs.readFileSync(path.join(root, 'views/partials/scripts.html'), 'utf8');
  assert.match(scripts, /tax = estimateTaxLocal\(taxableIncome, periodStartIso\);/);
  assert.doesNotMatch(scripts, /tax = estimateTaxLocal\(grossIncome,/);
  const start = scripts.indexOf('  function estimateTaxLocal(');
  const end = scripts.indexOf('\n  function renderAnnualViews()', start);
  assert.ok(start >= 0 && end > start);
  const clientContext = {
    Date,
    parseIsoDate(value) {
      const parts = String(value).split('-').map(Number);
      return new Date(parts[0], parts[1] - 1, parts[2]);
    }
  };
  vm.createContext(clientContext);
  vm.runInContext(scripts.slice(start, end), clientContext);

  const { context: serverContext } = createAppsScriptContext({});
  load(serverContext, 'backend/tax.js');
  [2372.69, 4735.19, 10000].forEach((income) => {
    assert.equal(clientContext.estimateTaxLocal(income, '2026-08-01'), serverContext.estimateTax(income, '2026-08-01'));
  });
  assert.equal(clientContext.estimateTaxLocal(2372.69, '2026-08-01'), 125);
});

test('backend globals and modified HTML scripts parse without duplicate functions', () => {
  const files = fs.readdirSync(path.join(root, 'backend')).filter((file) => file.endsWith('.js'));
  const names = new Map();
  files.forEach((file) => {
    const source = fs.readFileSync(path.join(root, 'backend', file), 'utf8');
    new vm.Script(source, { filename: file });
    for (const match of source.matchAll(/^function\s+([A-Za-z_$][\w$]*)\s*\(/gm)) {
      if (!names.has(match[1])) names.set(match[1], []);
      names.get(match[1]).push(file);
    }
  });
  const duplicates = [...names.entries()].filter(([, locations]) => locations.length > 1);
  assert.deepStrictEqual(duplicates, []);
  const main = fs.readFileSync(path.join(root, 'views/partials/scripts.html'), 'utf8').replace(/^<script>\s*/, '').replace(/<\/script>\s*$/, '');
  const mobile = fs.readFileSync(path.join(root, 'views/partials/mobile-entry-scripts.html'), 'utf8').replace(/^<script>\s*/, '').replace(/<\/script>\s*$/, '');
  const operations = fs.readFileSync(path.join(root, 'views/partials/operations-scripts.html'), 'utf8').replace(/^<script>\s*/, '').replace(/<\/script>\s*$/, '');
  new vm.Script(main);
  new vm.Script(mobile);
  new vm.Script(operations);
});

test('desktop day editor stacks timed sessions, exact duration rows, and day actions', () => {
  const html = fs.readFileSync(path.join(root, 'views/index.html'), 'utf8');
  const dashboard = fs.readFileSync(path.join(root, 'views/partials/dashboard.html'), 'utf8');
  const settings = fs.readFileSync(path.join(root, 'views/partials/settings.html'), 'utf8');
  const hourTypes = fs.readFileSync(path.join(root, 'views/partials/hourtypes.html'), 'utf8');
  const scripts = fs.readFileSync(path.join(root, 'views/partials/scripts.html'), 'utf8');
  const mobile = fs.readFileSync(path.join(root, 'views/partials/mobile-entry-scripts.html'), 'utf8');
  const styles = fs.readFileSync(path.join(root, 'views/partials/head.html'), 'utf8');
  assert.match(html, /id="calendar-clock-toggle"/);
  assert.match(dashboard, /id="day-clock-toggle"/);
  assert.match(scripts, /function updateDayClockToggle\(dateIso\)/);
  assert.match(scripts, /btn\.dataset\.clockAction = showOut \? 'out' : 'in'/);
  assert.match(scripts, /clockBtn\.dataset\.clockAction === 'out'\) quickPunchOutForDate/);
  assert.match(scripts, /clockLabel\.textContent = 'Clock out'/);
  assert.match(settings, /id="set-entry-mode-source"/);
  assert.match(hourTypes, /id="hour-type-entry-mode"/);
  assert.match(hourTypes, />Clock in\/out</);
  assert.match(scripts, /'Clock out — ' : 'Clock in — '/);
  assert.match(scripts, /commitEntryPunches\(entry\.id, punches, \{ successMsg: 'Clocked out', failMsg: 'Couldn’t clock out' \}\)/);
  assert.doesNotMatch(scripts, /await goToOpenPunch\(\);\s*handlePunchToggle\(\);/);
  assert.match(scripts, /Clock out is available one minute after clock in\./);
  assert.match(scripts, /punchToggleBtn\.disabled = !contractOptional && !hasContract/);
  assert.match(scripts, /state\.settings\.entry_mode_source === 'hour_type'/);
  assert.match(mobile, /state\.settings\.entry_mode_source === 'hour_type'/);
  assert.doesNotMatch(dashboard, /data-tab="manual"|data-tab="punch"/);
  assert.ok(dashboard.indexOf('id="day-sessions-timeline"') < dashboard.indexOf('id="day-duration-list"'));
  assert.ok(dashboard.indexOf('id="day-duration-list"') < dashboard.indexOf('id="day-level-actions"'));
  assert.match(dashboard, /id="day-overall-total"/);
  assert.match(scripts, /function renderDayDurationEntries\(dateIso\)/);
  assert.match(scripts, /data-duration-entry-id/);
  assert.match(scripts, /aria-expanded=/);
  assert.match(scripts, /function dayDurationDraftDirty\(\)/);
  assert.match(scripts, /if \(dayDurationDraft\) return false;/);
  assert.match(scripts, /const currentEditingIndex = \(\) => state\.entries\.findIndex/);
  assert.match(scripts, /if \(!stashDaySessionEdit\(\)\) return; \/\/ an invalid move stays open for correction/);
  assert.match(scripts, /Discard day changes\?/);
  assert.match(scripts, /function handleCalendarClick\(dateIso, targetHourTypeId, targetEntryId\)/);
  assert.match(scripts, /if \(matches\.length === 1\) target = matches\[0\]/);
  // Simple days are one-and-done: open on Total hours, save/cancel/delete close the day.
  assert.match(scripts, /function isSimpleQuickDay\(dateIso\)[\s\S]*?if \(dayEntryModes\(dateIso\)\.detailed\) return false;/);
  assert.match(scripts, /\} else if \(daySimpleQuick\) \{\s*beginDayDurationDraft\(/);
  assert.match(scripts, /renderCalendar\(\);\s*if \(daySimpleQuick\) \{ closeDayModal\(true\); return; \}/);
  assert.match(styles, /#modal-day-entry\.is-simple-quick \.ts-day-duration-add/);
  assert.match(scripts, /hourTypeEffectiveEntryMode\(selectedHourType\) === 'detailed'/);
  assert.match(mobile, /state\.mode = hourTypeEffectiveEntryMode\(selected\) === 'detailed' \? 'punch' : 'manual'/);
  assert.match(scripts, /scheduleEntryMode\(draft\.contract_id, draft\.hour_type_id\)/);
  assert.doesNotMatch(scripts, /scheduleContractEntryMode/);
  assert.match(styles, /\.ts-pace-burndown \+ \.ts-pace-month-section/);
  assert.doesNotMatch(styles, /\.ts-pace-card-divider/);
  assert.doesNotMatch(scripts, /className = 'ts-pace-card-divider'/);
});

test('calendar preferences, resilient refresh, quick fill, and rolling holiday window are wired', () => {
  const settings = fs.readFileSync(path.join(root, 'views/partials/settings.html'), 'utf8');
  const styles = fs.readFileSync(path.join(root, 'views/partials/head.html'), 'utf8');
  const scripts = fs.readFileSync(path.join(root, 'views/partials/scripts.html'), 'utf8');
  const holidays = fs.readFileSync(path.join(root, 'views/partials/public-holidays.html'), 'utf8');
  assert.match(settings, /data-setting-key="subdue_weekends"/);
  assert.match(scripts, /subdue_weekends:[\s\S]*?defaultValue: true/);
  assert.match(scripts, /classList\.toggle\('ts-subdue-weekends', enabled\)/);
  assert.match(styles, /\.ts-subdue-weekends \.ts-calendar-cell--weekend/);
  assert.match(styles, /\.ts-subdue-weekends \.ts-agenda-mini-cell\.is-weekend/);
  assert.doesNotMatch(styles, /\.ts-subdue-weekends \.ts-calendar-cell--weekend[^{]*\{[^}]*background:/);
  assert.doesNotMatch(styles, /\.ts-subdue-weekends \.ts-agenda-mini-cell\.is-weekend[^{]*\{[^}]*background:/);
  ['networkerror', 'connection failure', 'http 0'].forEach((fragment) => assert.ok(scripts.includes(`'${fragment}'`)));
  assert.match(scripts, /updateManualContractVisibility\(defaultData\.contract_id \? String\(defaultData\.contract_id\) : ''\)/);
  assert.match(scripts, /reflectManualSessionsConflict\(\);/);
  assert.match(scripts, /Select a contract before applying this duration\./);
  assert.match(scripts, /function phHolidayInDisplayWindow\(holiday, todayStr, currentYear\)/);
  assert.match(scripts, /return \[ps\.year, ps\.year \+ 1\]/);
  assert.match(scripts, /return iso < isoDate\(cutoff\)/);
  assert.match(holidays, /rolling look-ahead window/);

  const transientContext = {};
  vm.createContext(transientContext);
  vm.runInContext(extractClientFunction(scripts, 'entriesSyncErrorText') + '\n'
    + extractClientFunction(scripts, 'isTransientEntriesSyncError'), transientContext);
  assert.equal(transientContext.isTransientEntriesSyncError({ message:'NetworkError: Connection failure due to HTTP 0' }), true);

  const holidayContext = {
    normalizeHolidayDate: (value) => /^\d{4}-\d{2}-\d{2}$/.test(String(value || '')) ? String(value) : '',
    parseIsoDate: (value) => {
      const parts = String(value || '').split('-').map(Number);
      return parts.length === 3 ? new Date(parts[0], parts[1] - 1, parts[2]) : null;
    },
    isoDate: (date) => [date.getFullYear(), String(date.getMonth() + 1).padStart(2, '0'), String(date.getDate()).padStart(2, '0')].join('-')
  };
  vm.createContext(holidayContext);
  vm.runInContext(extractClientFunction(scripts, 'phHolidayInDisplayWindow'), holidayContext);
  assert.equal(holidayContext.phHolidayInDisplayWindow({ date:'2026-12-25' }, '2026-01-27', 2026), true);
  assert.equal(holidayContext.phHolidayInDisplayWindow({ date:'2027-01-26' }, '2026-01-27', 2026), true);
  assert.equal(holidayContext.phHolidayInDisplayWindow({ date:'2027-03-08' }, '2026-01-27', 2026), false);
});

test('break windows exist only between live sessions and late creates cannot resurrect an orphan', () => {
  const scripts = fs.readFileSync(path.join(root, 'views/partials/scripts.html'), 'utf8');
  const context = {
    state: { entries: [] },
    resolveEntryType: (entry) => entry.entry_type,
    entryPunches: (entry) => entry.punches || [],
    timeToMinutes: (value) => {
      const match = /^(\d{2}):(\d{2})$/.exec(String(value || ''));
      return match ? Number(match[1]) * 60 + Number(match[2]) : null;
    }
  };
  vm.createContext(context);
  vm.runInContext(extractClientFunction(scripts, 'daySessionBreakWindows'), context);
  const session = (id, start, end) => ({ id, date:'2026-08-13', entry_type:'advanced', punches:[{ in:start, out:end }] });

  context.state.entries = [session('a', '09:00', '10:00'), session('b', '11:00', '12:00')];
  assert.deepStrictEqual(Array.from(context.daySessionBreakWindows('2026-08-13'), (pair) => Array.from(pair)), [[600, 660]]);
  context.state.entries = [session('a', '09:00', '10:00')];
  assert.deepStrictEqual(Array.from(context.daySessionBreakWindows('2026-08-13'), (pair) => Array.from(pair)), []);
  context.state.entries = [session('a', '09:00', '10:00'), session('open', '11:00', '')];
  assert.deepStrictEqual(Array.from(context.daySessionBreakWindows('2026-08-13'), (pair) => Array.from(pair)), [[600, 660]]);
  context.state.entries = [session('a', '09:00', '10:00'), session('open', '11:00', ''), session('c', '13:00', '14:00')];
  assert.deepStrictEqual(Array.from(context.daySessionBreakWindows('2026-08-13'), (pair) => Array.from(pair)), [[600, 660]]);

  assert.match(scripts, /The bounding session disappeared while this create was in flight/);
  assert.match(scripts, /Array\.from\(new Set\(\(state\.breaks \|\| \[\]\)/);
  assert.match(scripts, /const validBreakWindows = daySessionBreakWindows\(dateIso\)/);
  assert.doesNotMatch(scripts, /if \(typeof dayEntryModes === 'function' && !dayEntryModes\(dateIso\)\.detailed\) return/);
});

test('release generators stamp fixed full metadata and a valid stable manifest', () => {
  const versionGenerator = fs.readFileSync(path.join(root, 'scripts/write-version.sh'), 'utf8');
  const manifestGenerator = fs.readFileSync(path.join(root, 'scripts/write-release-manifest.sh'), 'utf8');
  const version = fs.readFileSync(path.join(root, 'backend/version.js'), 'utf8');
  assert.match(versionGenerator, /TEMPUS_BUILD_COMMIT/);
  assert.match(versionGenerator, /--timestamp/);
  assert.match(versionGenerator, /buildTimestamp:/);
  assert.match(version, /buildTimestamp: '\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z'/);
  assert.match(version, /buildCommit: '[0-9a-f]{40}'/);
  assert.match(manifestGenerator, /"schema_version": 1/);
  assert.match(manifestGenerator, /"channel": "stable"/);
  assert.match(manifestGenerator, /\{40\}/);
  assert.match(manifestGenerator, /release\.zip/);
});

test('stable release comparison handles exact, abbreviated, newer, divergent, legacy, cached, and unavailable metadata', () => {
  const sha = '0123456789abcdef0123456789abcdef01234567';
  const other = 'fedcba9876543210fedcba9876543210fedcba98';
  const manifest = { schema_version:1, channel:'stable', commit_sha:sha, published_at:'2026-08-08T10:00:00Z', display_version:'v', archive_url:'https://example.com/release.zip' };
  const makeContext = (local, responseFactory) => {
    const { context } = createAppsScriptContext({});
    let cached = null;
    context.BUILD_META = local;
    context.cacheGet = () => cached;
    context.cacheSet = (key, value) => { cached = value; };
    context.UrlFetchApp = { fetch: responseFactory };
    load(context, 'backend/releaseInfo.js');
    return context;
  };
  const response = () => ({ getResponseCode:() => 200, getContentText:() => JSON.stringify(manifest) });
  let context = makeContext({ buildCommit:sha, buildTimestamp:'2026-08-08T09:00:00Z', buildDate:'2026-08-08' }, response);
  assert.equal(context.api_checkForUpdate().relation, 'current');
  assert.equal(context.compareStableRelease_({ buildCommit:sha.slice(0, 12) }, manifest).relation, 'current');
  assert.equal(context.compareStableRelease_({ buildCommit:other, buildTimestamp:'2026-08-07T09:00:00Z' }, manifest).hasUpdate, true);
  assert.equal(context.compareStableRelease_({ buildCommit:other, buildTimestamp:'2026-08-09T09:00:00Z' }, manifest).relation, 'local_newer');
  assert.equal(context.compareStableRelease_({ buildCommit:other, buildTimestamp:'2026-08-08T10:00:00Z' }, manifest).relation, 'divergent');
  const legacy = context.compareStableRelease_({ buildDate:'2026-08-07' }, manifest);
  assert.equal(legacy.comparisonMode, 'legacy_date');
  assert.equal(legacy.hasUpdate, true);
  let fetches = 0;
  context = makeContext({ buildCommit:sha }, () => { fetches++; return response(); });
  context.api_checkForUpdate(); context.api_checkForUpdate();
  assert.equal(fetches, 1);
  context = makeContext({}, () => ({ getResponseCode:() => 200, getContentText:() => '{bad' }));
  assert.equal(context.api_checkForUpdate().recoverableError.code, 'invalid_json');
  context = makeContext({}, () => { throw new Error('offline'); });
  assert.equal(context.api_checkForUpdate().comparisonMode, 'unavailable');
});

test('updater dry run preserves clasp config and accepts only built release packages', () => {
  const updater = fs.readFileSync(path.join(root, 'scripts/update-tempus.sh'), 'utf8');
  assert.match(updater, /refs\/heads\/release/);
  assert.match(updater, /--dry-run/);
  assert.match(updater, /--archive-file/);
  assert.match(updater, /cp "\$target_dir\/\.clasp\.json" "\$clasp_backup"/);
  assert.match(updater, /cp "\$clasp_backup" "\$target_dir\/\.clasp\.json"/);
  assert.match(updater, /manifest missing/);
  assert.match(updater, /clasp push -f/);
  assert.match(fs.readFileSync(path.join(root, '.claspignore'), 'utf8'), /^release-manifest\.json$/m);
});

test('release workflow gates force-publish on tests and contains no Apps Script deployment', () => {
  const workflow = fs.readFileSync(path.join(root, '.github/workflows/publish-release.yml'), 'utf8');
  assert.match(workflow, /branches: \[main\]/);
  assert.match(workflow, /workflow_dispatch:/);
  assert.match(workflow, /contents: write/);
  // A detached HEAD is a commit, not a ref, so git cannot guess an unqualified destination branch.
  assert.ok(workflow.indexOf('npm test') < workflow.indexOf('push --force origin HEAD:refs/heads/release'));
  assert.match(workflow, /git worktree add --detach/);
  // GITHUB_TOKEN is never allowed to create or update workflow files, so a release tree that still
  // carried .github/workflows would have the force-push rejected on every run.
  assert.ok(workflow.indexOf('rm -rf "$build_dir/.github"') < workflow.indexOf('push --force origin HEAD:refs/heads/release'));
  assert.doesNotMatch(workflow, /clasp push|script\.google\.com|GOOGLE_/i);
});

test('quick actions expose menu semantics and keyboard controls', () => {
  const html = fs.readFileSync(path.join(root, 'views/index.html'), 'utf8');
  const scripts = fs.readFileSync(path.join(root, 'views/partials/scripts.html'), 'utf8');
  assert.match(html, /id="calendar-context-menu" role="menu"/);
  ['ContextMenu', 'Shift+F10', 'ArrowDown', 'ArrowUp', 'Escape'].forEach((token) => {
    if (token === 'Shift+F10') assert.match(scripts, /event\.shiftKey && event\.key === 'F10'/);
    else assert.ok(scripts.includes(token));
  });
});

test('variance details omit internal badge labels and entry sync suppresses stale failures', () => {
  const scripts = fs.readFileSync(path.join(root, 'views/partials/scripts.html'), 'utf8');
  const styles = fs.readFileSync(path.join(root, 'views/partials/head.html'), 'utf8');
  assert.doesNotMatch(scripts, /tag\.textContent\s*=\s*['"]your badge['"]/i);
  assert.doesNotMatch(styles, /\.ts-pace-stat-badge\b/);
  assert.match(scripts, /const ENTRY_SYNC_RETRY_DELAYS_MS\s*=\s*\[1500, 4000\]/);
  assert.match(scripts, /if \(!endEntriesSync\(reqSeq\)\) \{ resolve\(true\); return; \}/);
  assert.match(scripts, /console\.error\('\[Entries\] Initial sync failed:', error\)/);
});

test('hour types load once per boot and retry the upgrade guard instead of alarming the user', () => {
  const scripts = fs.readFileSync(path.join(root, 'views/partials/scripts.html'), 'utf8');
  const migrations = fs.readFileSync(path.join(root, 'backend/migrations.js'), 'utf8');
  // The retry guard has to match the message the upgrade guard actually throws.
  const guardMessage = (migrations.match(/throw new Error\('(upgrade_required:[^']+)'\)/) || [])[1];
  assert.ok(guardMessage, 'backend/migrations.js should throw an upgrade_required error');
  const matcher = scripts.match(/function isUpgradeInProgressError\(error\) \{([\s\S]*?)\n  \}/);
  assert.ok(matcher, 'isUpgradeInProgressError should exist');
  const needles = [...matcher[1].matchAll(/indexOf\('([^']+)'\)/g)].map((m) => m[1]);
  assert.ok(needles.length > 0, 'isUpgradeInProgressError should test for substrings');
  assert.ok(
    needles.some((needle) => guardMessage.toLowerCase().includes(needle)),
    'isUpgradeInProgressError matches none of: ' + guardMessage
  );
  // Concurrent boot paths must share one api_getHourTypes call — the first-load migrations inside it
  // take the script lock, so a second execution loses the race and reports a failure over good data.
  assert.match(scripts, /let hourTypesLoadPromise = null;/);
  assert.match(scripts, /if \(hourTypesLoadPromise\) return hourTypesLoadPromise;/);
  assert.match(scripts, /hourTypesLoadPromise = null;\n      return loaded;/);
  const loader = scripts.slice(scripts.indexOf('function loadHourTypes()'), scripts.indexOf('function renderHourTypes()'));
  assert.match(loader, /isUpgradeInProgressError\(error\) && upgradeRetryCount < UPGRADE_RETRY_DELAYS_MS\.length/);
  assert.match(loader, /isConcurrentUpdateError\(error\) \|\| isTransientEntriesSyncError\(error\)/);
  assert.match(scripts, /function isConcurrentUpdateError\(error\)[\s\S]*?concurrent_update/);
});

test('upgrade completion provides a top-level continuation link and refresh fallback', () => {
  const upgrade = fs.readFileSync(path.join(root, 'views/upgrade.html'), 'utf8');
  assert.match(upgrade, /id="continue"[^>]+target="_top"/);
  assert.match(upgrade, /Upgrade complete\. Refresh this page to open Tempus\./);
  assert.doesNotMatch(upgrade, /window\.top\.location\.reload/);
});

test('public holiday refresh replaces legacy blank-year rows instead of appending duplicates', () => {
  const headers = ['id', 'date', 'name', 'region', 'source', 'active', 'created_at', 'updated_at', 'local_name', 'counties', 'types', 'year', 'fetched_at'];
  const { context, spreadsheet } = createAppsScriptContext({
    public_holidays: [headers, ['legacy', '2026-01-01', "New Year's Day", '["AU"]', 'legacy', 'TRUE', '', '', "New Year's Day", '', '["Public"]', '', '']]
  });
  context.assertMigrationsSettled_ = () => true;
  context.cacheClearPrefix = () => {};
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/publicholidays.js');
  assert.equal(context.repairPublicHolidayYears_(), 1);
  context.replaceYearHolidays(2026, [{ date:'2026-01-01', name:"New Year's Day", localName:"New Year's Day", counties:null }]);
  const rows = spreadsheet.getSheetByName('public_holidays').snapshot();
  assert.equal(rows.length, 2);
  assert.equal(rows[1][headers.indexOf('year')], 2026);
});

test('public holiday time population adopts an existing date and hour type under the server lock', () => {
  const headers = ['id', 'date', 'duration_minutes', 'contract_id', 'created_at', 'punches_json', 'entry_type', 'hour_type_id', 'recurrence_id', 'note', 'assessment_id', 'source_type', 'source_id', 'source_occurrence_key', 'client_request_id'];
  const { context, spreadsheet } = createAppsScriptContext({
    timesheet_entries: [headers, ['existing', '2026-01-01', 450, '', '2026-01-01T00:00:00Z', '[]', 'basic', 'leave', '', '', '', 'manual', '', '', '']]
  });
  context.assertMigrationsSettled_ = () => true;
  context.getDefaultHourTypeId = () => 'work';
  context.api_getSettings = () => ({});
  context.cacheClearPrefix = () => {};
  load(context, 'backend/integrity.js');
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/entries.js');
  load(context, 'backend/publicholidays.js');
  const duplicate = context.api_addPublicHolidayEntry({ date:'2026-01-01', duration_minutes:450, hour_type_id:'leave' });
  assert.equal(duplicate.error, 'duplicate_entry');
  assert.equal(spreadsheet.getSheetByName('timesheet_entries').getLastRow(), 2);
  const added = context.api_addPublicHolidayEntry({ date:'2026-01-02', duration_minutes:450, hour_type_id:'leave' });
  assert.equal(added.success, true);
  assert.equal(added.entry.source_type, 'public_holiday');
  assert.equal(added.entry.source_occurrence_key, '2026-01-02');
  assert.deepStrictEqual(Array.from(added.entry.punches), []);
});

// Regression: the source-identity migration keyed every derived row on its date and deleted collisions,
// which destroyed assessment actual-hours tracking rows sharing a date with the assessment's billable
// row. 13 rows (86.5 tracked hours) were lost on the live sheet this way.
const ASSESSMENT_LEGACY_HEADERS = ['id', 'date', 'duration_minutes', 'contract_id', 'created_at', 'punches_json', 'entry_type', 'hour_type_id', 'recurrence_id', 'assessment_id'];
function assessmentLegacyRow(id, date, minutes, hourTypeId, createdAt) {
  return [id, date, minutes, 'contract-1', createdAt, '[]', 'basic', hourTypeId, '', 'assess-1'];
}
function assessmentMigrationContext(sheets) {
  const built = createAppsScriptContext(sheets);
  built.context.assertMigrationsSettled_ = () => true;
  built.context.getDefaultHourTypeId = () => 'hourtype-work';
  built.context.api_getSettings = () => ({});
  built.context.cacheClearPrefix = () => {};
  load(built.context, 'backend/integrity.js');
  load(built.context, 'backend/sheets.ts.js');
  load(built.context, 'backend/entries.js');
  load(built.context, 'backend/migrations.js');
  // Migrations run inside api_runUpgrade's context, which is what lets them write to the sheet.
  built.context.MIGRATION_CONTEXT.active = true;
  return built;
}
function timesheetSnapshot(spreadsheet) {
  const rows = spreadsheet.getSheetByName('timesheet_entries').snapshot();
  const headers = rows[0];
  return rows.slice(1).filter((row) => row[headers.indexOf('id')] !== '').map((row) => ({
    id: row[headers.indexOf('id')],
    date: row[headers.indexOf('date')],
    duration_minutes: row[headers.indexOf('duration_minutes')],
    hour_type_id: row[headers.indexOf('hour_type_id')],
    source_type: row[headers.indexOf('source_type')],
    source_id: row[headers.indexOf('source_id')],
    occurrence: row[headers.indexOf('source_occurrence_key')]
  }));
}

test('source identity migration keeps every assessment tracking row that shares a date with the billable row', () => {
  const { context, spreadsheet } = assessmentMigrationContext({
    timesheet_entries: [
      ASSESSMENT_LEGACY_HEADERS,
      assessmentLegacyRow('billable', '2026-08-17', 60, 'hourtype-work', '2026-08-17T12:25:42Z'),
      assessmentLegacyRow('track-same-day', '2026-08-17', 360, 'hourtype-track', '2026-08-17T12:26:06Z'),
      assessmentLegacyRow('track-second-same-day', '2026-08-17', 90, 'hourtype-track', '2026-08-17T12:27:00Z'),
      assessmentLegacyRow('track-other-day', '2026-08-18', 180, 'hourtype-track', '2026-08-19T12:48:22Z')
    ]
  });
  context.migrationTimesheetSources_();
  const rows = timesheetSnapshot(spreadsheet);
  assert.deepStrictEqual(rows.map((row) => row.id).sort(), ['billable', 'track-other-day', 'track-same-day', 'track-second-same-day']);
  assert.equal(rows.reduce((sum, row) => sum + Number(row.duration_minutes), 0), 690);
  rows.forEach((row) => {
    assert.equal(row.source_type, 'assessment');
    assert.equal(row.source_id, 'assess-1');
    assert.equal(row.occurrence, row.id === 'billable' ? 'billable' : '');
  });
  // Re-running must not re-key or drop anything.
  context.migrationTimesheetSources_();
  assert.deepStrictEqual(timesheetSnapshot(spreadsheet), rows);
});

test('source identity migration retains rather than deletes a genuine recurring-source collision', () => {
  const headers = ['id', 'date', 'duration_minutes', 'contract_id', 'created_at', 'punches_json', 'entry_type', 'hour_type_id', 'recurrence_id'];
  const { context, spreadsheet } = assessmentMigrationContext({
    timesheet_entries: [
      headers,
      ['rec-a', '2026-08-17', 450, 'contract-1', '2026-08-17T01:00:00Z', '[]', 'basic', 'hourtype-work', 'rule-1'],
      ['rec-b', '2026-08-17', 120, 'contract-1', '2026-08-17T02:00:00Z', '[]', 'basic', 'hourtype-work', 'rule-1']
    ]
  });
  context.migrationTimesheetSources_();
  const rows = timesheetSnapshot(spreadsheet);
  assert.equal(rows.length, 2);
  assert.equal(rows[0].occurrence, '2026-08-17');
  assert.equal(rows[1].occurrence, '', 'the colliding row keeps its data and loses only the wrong derived key');
  const archive = spreadsheet.getSheetByName('migration_archive').snapshot();
  assert.equal(archive.length, 2);
  assert.equal(archive[1][archive[0].indexOf('reason')], 'derived_source_collision_retained');
});

test('repair migration restores the archived tracking rows and re-keys the survivors', () => {
  const canonical = ['id', 'date', 'duration_minutes', 'contract_id', 'created_at', 'punches_json', 'entry_type', 'hour_type_id', 'recurrence_id', 'note', 'assessment_id', 'source_type', 'source_id', 'source_occurrence_key', 'client_request_id'];
  const deleted = { id: 'track-same-day', date: '2026-08-17', duration_minutes: 360, contract_id: 'contract-1', created_at: '2026-08-17T12:26:06Z', punches_json: '[]', entry_type: 'basic', hour_type_id: 'hourtype-track', recurrence_id: '', note: '', assessment_id: 'assess-1', source_type: 'assessment', source_id: 'assess-1', source_occurrence_key: '2026-08-17', client_request_id: '' };
  const { context, spreadsheet } = assessmentMigrationContext({
    timesheet_entries: [
      canonical,
      ['billable', '2026-08-17', 60, 'contract-1', '2026-08-17T12:25:42Z', '[]', 'basic', 'hourtype-work', '', '', 'assess-1', 'assessment', 'assess-1', '2026-08-17', ''],
      ['track-other-day', '2026-08-18', 180, 'contract-1', '2026-08-19T12:48:22Z', '[]', 'basic', 'hourtype-track', '', '', 'assess-1', 'assessment', 'assess-1', '2026-08-18', '']
    ],
    migration_archive: [
      ['id', 'migration_id', 'source_sheet', 'source_row_json', 'reason', 'archived_at'],
      ['archive-1', '2026-08-source-aware-timesheet-entries', 'timesheet_entries', JSON.stringify(deleted), 'duplicate_derived_source', '2026-08-20T00:00:00Z'],
      ['archive-2', '2026-08-company-expense-ledger', 'deductions', '{"id":"unrelated"}', 'migrated_to_expense_rule', '2026-08-20T00:00:00Z']
    ]
  });
  context.migrationRepairAssessmentEntryIdentity_();
  const rows = timesheetSnapshot(spreadsheet);
  assert.deepStrictEqual(rows.map((row) => row.id).sort(), ['billable', 'track-other-day', 'track-same-day']);
  const restored = rows.filter((row) => row.id === 'track-same-day')[0];
  assert.equal(restored.duration_minutes, 360, 'the lost tracked hours come back intact');
  assert.equal(restored.date, '2026-08-17');
  assert.equal(restored.hour_type_id, 'hourtype-track');
  assert.equal(restored.occurrence, '');
  assert.equal(rows.filter((row) => row.id === 'billable')[0].occurrence, 'billable');
  assert.equal(rows.filter((row) => row.id === 'track-other-day')[0].occurrence, '', 'a survivor keyed on its date is re-keyed as tracking');
  context.migrationRepairAssessmentEntryIdentity_();
  assert.deepStrictEqual(timesheetSnapshot(spreadsheet), rows, 'restore is idempotent');
});

test('assessment time entries accept many rows per day while the billable row stays unique', () => {
  const canonical = ['id', 'date', 'duration_minutes', 'contract_id', 'created_at', 'punches_json', 'entry_type', 'hour_type_id', 'recurrence_id', 'note', 'assessment_id', 'source_type', 'source_id', 'source_occurrence_key', 'client_request_id'];
  const { context, spreadsheet } = createAppsScriptContext({ timesheet_entries: [canonical] });
  context.assertMigrationsSettled_ = () => true;
  context.getDefaultHourTypeId = () => 'hourtype-work';
  context.api_getSettings = () => ({});
  context.cacheClearPrefix = () => {};
  load(context, 'backend/integrity.js');
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/entries.js');

  const tracking = { source_type: 'assessment', source_id: 'assess-1', assessment_id: 'assess-1', source_occurrence_key: '', hour_type_id: 'hourtype-track', contract_id: 'contract-1', entry_type: 'basic' };
  const first = context.api_addEntry(Object.assign({}, tracking, { date: '2026-08-17', duration_minutes: 360 }));
  const second = context.api_addEntry(Object.assign({}, tracking, { date: '2026-08-17', duration_minutes: 90 }));
  assert.equal(first.success, true);
  assert.equal(second.success, true, 'a second tracking row on the same day is not a duplicate');

  const billable = { source_type: 'assessment', source_id: 'assess-1', assessment_id: 'assess-1', source_occurrence_key: 'billable', hour_type_id: 'hourtype-work', contract_id: 'contract-1', entry_type: 'basic', date: '2026-08-17', duration_minutes: 60 };
  assert.equal(context.api_addEntry(billable).success, true);
  const repeat = context.api_addEntry(Object.assign({}, billable, { date: '2026-09-01' }));
  assert.equal(repeat.error, 'duplicate_entry', 'the billable row is unique per assessment regardless of date');
  assert.equal(repeat.entry.id, spreadsheet.getSheetByName('timesheet_entries').snapshot()[3][0]);
  assert.equal(spreadsheet.getSheetByName('timesheet_entries').getLastRow(), 4);
});

test('recovery report counts the archived rows before the repair runs and none after', () => {
  const canonical = ['id', 'date', 'duration_minutes', 'contract_id', 'created_at', 'punches_json', 'entry_type', 'hour_type_id', 'recurrence_id', 'note', 'assessment_id', 'source_type', 'source_id', 'source_occurrence_key', 'client_request_id'];
  const deleted = { id: 'track-same-day', date: '2026-08-17', duration_minutes: 360, contract_id: 'contract-1', created_at: '2026-08-17T12:26:06Z', punches_json: '[]', entry_type: 'basic', hour_type_id: 'hourtype-track', recurrence_id: '', note: '', assessment_id: 'assess-1', source_type: 'assessment', source_id: 'assess-1', source_occurrence_key: '2026-08-17', client_request_id: '' };
  const { context } = assessmentMigrationContext({
    timesheet_entries: [
      canonical,
      ['billable', '2026-08-17', 60, 'contract-1', '2026-08-17T12:25:42Z', '[]', 'basic', 'hourtype-work', '', '', 'assess-1', 'assessment', 'assess-1', '2026-08-17', '']
    ],
    migration_archive: [
      ['id', 'migration_id', 'source_sheet', 'source_row_json', 'reason', 'archived_at'],
      ['archive-1', '2026-08-source-aware-timesheet-entries', 'timesheet_entries', JSON.stringify(deleted), 'duplicate_derived_source', '2026-08-20T00:00:00Z']
    ]
  });
  const before = context.api_getTimesheetRecoveryReport();
  assert.equal(before.entry_rows, 1);
  assert.equal(before.entry_minutes, 60);
  assert.equal(before.archived, 1);
  assert.equal(before.recoverable, 1);
  assert.equal(before.recoverable_minutes, 360);
  assert.equal(before.rows[0].id, 'track-same-day');

  context.migrationRepairAssessmentEntryIdentity_();
  const after = context.api_getTimesheetRecoveryReport();
  assert.equal(after.entry_rows, 2);
  assert.equal(after.entry_minutes, 420, 'the recovered minutes land in the sheet');
  assert.equal(after.recoverable, 0);
  assert.equal(after.already_present, 1);
});

// ---- sole trader tax provisioning and the monthly transfer split ----

function taxContext() {
  const { context } = createAppsScriptContext({});
  context.assertMigrationsSettled_ = () => true;
  context.api_getSettings = () => ({});
  context.listBasSubmissionsInternal = () => [];
  load(context, 'backend/integrity.js');
  load(context, 'backend/tax.js');
  return context;
}

test('annual individual tax applies progressive brackets, the medicare shade-in and the 2026-27 rate cut', () => {
  const context = taxContext();
  assert.equal(context.annualIncomeTax(18200, 2024), 0, 'tax free threshold');
  // 2024-25: 16% on 45000-18200 = 4288, plus 2% medicare on 45000 = 900.
  assert.equal(context.annualIncomeTax(45000, 2024), 4288 + 900);
  // 2026-27 drops that bracket to 15%: 26800 * 0.15 = 4020.
  assert.equal(context.annualIncomeTax(45000, 2026), 4020 + 900);
  // Shade-in: just above the threshold the levy is 10% of the excess, not 2% of income.
  const shaded = context.annualIncomeTaxDetails(28000, 2024);
  assert.equal(shaded.medicare_levy, Math.round((28000 - 27222) * 0.10 * 100) / 100);
  assert.ok(shaded.medicare_levy < 28000 * 0.02);
  // Well above the shade-in range it is a flat 2%.
  assert.equal(context.annualIncomeTaxDetails(120000, 2024).medicare_levy, 2400);
  // Progressive across the third bracket: 4288 + (135000-45000)*0.30 = 31288, medicare 2700.
  assert.equal(context.annualIncomeTax(135000, 2024), 31288 + 2700);
});

test('annual tax table reports staleness beyond the bundled years and honours a settings override', () => {
  const context = taxContext();
  const stale = context.api_getIndividualTaxTableStatus(2031);
  assert.equal(stale.stale, true);
  assert.equal(stale.applied_financial_year, '2026-27');
  assert.match(stale.warning, /individual_tax_table_override/);
  assert.equal(context.api_getIndividualTaxTableStatus(2026).stale, false);

  context.api_getSettings = () => ({ individual_tax_table_override: JSON.stringify({ brackets: [[10000, 0], [null, 0.5]], medicare: { rate: 0, lower_threshold: 0, shade_in_rate: 0 } }) });
  assert.equal(context.annualIncomeTax(20000, 2026), 5000, 'override replaces the bundled brackets');
  assert.equal(context.api_getIndividualTaxTableStatus(2026).applied_financial_year, 'settings_override');
});

test('instalment rate method applies to gross income; cumulative method self-corrects on lumpy income', () => {
  const context = taxContext();
  const byRate = context.soleTraderTaxProvision_({ instalment_rate: 0.15, period_instalment_income: 20000, financial_year: 2026 });
  assert.equal(byRate.method, 'instalment_rate');
  assert.equal(byRate.amount, 3000, 'rate applies to gross business income, not profit');

  // Run a year of monthly profits through the cumulative method and return each month's set-aside.
  const runYear = (monthlyProfits) => {
    let provisioned = 0;
    let ytd = 0;
    return monthlyProfits.map((profit, index) => {
      ytd = Math.round((ytd + profit) * 100) / 100;
      const provision = context.soleTraderTaxProvision_({
        instalment_rate: null, ytd_taxable_profit: ytd, months_elapsed: index + 1,
        already_provisioned: provisioned, financial_year: 2026
      });
      provisioned = Math.round((provisioned + provision.amount) * 100) / 100;
      return provision.amount;
    });
  };
  const annualOn60k = context.annualIncomeTax(60000, 2026);

  // One 60k month then eleven empty ones. The whole year's tax is set aside in the month the money
  // arrived, and nothing further is taken once the profit stops.
  const frontLoaded = runYear([60000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
  assert.equal(frontLoaded[0], annualOn60k, 'month one holds the true tax on income received to date');
  assert.ok(frontLoaded.slice(1).every((amount) => amount === 0));
  // The withholding-table approach would have annualised that month to a 720k year.
  assert.ok(frontLoaded[0] < context.annualIncomeTax(60000 * 12, 2026) / 12);

  // The annual total is independent of how the income arrived — the property that makes it correct.
  const shapes = [
    [60000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    [5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000],
    [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 60000],
    [20000, 0, 15000, 0, 0, 9000, 0, 16000, 0, 0, 0, 0]
  ];
  shapes.forEach((shape, index) => {
    const total = runYear(shape).reduce((sum, amount) => Math.round((sum + amount) * 100) / 100, 0);
    assert.equal(total, annualOn60k, 'shape ' + index + ' totals the exact annual tax on 60k profit');
  });

  // A loss month releases previously set-aside tax rather than stranding it.
  const withLoss = runYear([40000, 30000, -20000, 10000, 0, 0, 0, 0, 0, 0, 0, 0]);
  assert.ok(withLoss[2] < 0, 'the loss month hands money back');
  assert.equal(
    withLoss.reduce((sum, amount) => Math.round((sum + amount) * 100) / 100, 0),
    annualOn60k,
    'and the year still totals the tax on the 60k of actual profit'
  );
});

function splitContext(sheets, settings) {
  const built = createAppsScriptContext(sheets);
  const { context } = built;
  context.assertMigrationsSettled_ = () => true;
  context.api_getSettings = () => settings || {};
  context.listBasSubmissionsInternal = () => [];
  context.cacheGet = () => null;
  context.cacheSet = () => {};
  context.cacheClearPrefix = () => {};
  context.getDefaultHourTypeId = () => 'hourtype-work';
  load(context, 'backend/integrity.js');
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/entries.js');
  load(context, 'backend/tax.js');
  load(context, 'backend/settings.js');
  // The split only needs the four BAS actuals; stub the ledger reads so the arithmetic is what is tested.
  context.basActualsForRange_ = (basis, from, to) => (built.actuals[from] || { g1_total_sales: 0, gst_on_sales: 0, purchases: 0, gst_on_purchases: 0 });
  load(context, 'backend/financialReporting.js');
  context.basActualsForRange_ = (basis, from, to) => (built.actuals[from] || { g1_total_sales: 0, gst_on_sales: 0, purchases: 0, gst_on_purchases: 0 });
  load(context, 'backend/transferSplit.js');
  // settings.js and tax.js define their own api_getSettings/lock helpers, so re-stub after every load.
  context.api_getSettings = () => settings || {};
  context.listBasSubmissionsInternal = () => [];
  built.actuals = {};
  return built;
}

test('monthly transfer split reaches every month of the financial year', () => {
  const built = splitContext({});
  const report = built.context.api_getMonthlyTransferSplit({ financial_year: 2026 });
  assert.equal(report.months.length, 12);
  assert.deepStrictEqual(Array.from(report.months.map((month) => month.period.from)), [
    '2026-07-01', '2026-08-01', '2026-09-01', '2026-10-01', '2026-11-01', '2026-12-01',
    '2027-01-01', '2027-02-01', '2027-03-01', '2027-04-01', '2027-05-01', '2027-06-01'
  ]);
  assert.deepStrictEqual(Array.from(report.months.map((month) => month.bas_quarter)), [1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4]);
});

test('transfer split sets aside net GST, super and tax, and the offset is her after-tax income', () => {
  const built = splitContext({}, { payg_instalment_rate: 15 });
  // One month: $11,000 received including $1,000 GST, $2,200 of costs including $200 claimable GST.
  built.actuals['2026-08-01'] = { g1_total_sales: 11000, gst_on_sales: 1000, purchases: 2200, gst_on_purchases: 200 };
  const report = built.context.api_getMonthlyTransferSplit({ financial_year: 2026 });
  const august = report.months.filter((month) => month.calendar_month === 7)[0];

  assert.equal(august.received_ex_gst, 10000);
  assert.equal(august.expenses_ex_gst, 2000);
  assert.equal(august.profit_before_super, 8000);
  assert.equal(august.super_rate, 0.12);
  assert.equal(august.super, 960);
  assert.equal(august.taxable_profit, 7040);
  assert.equal(august.net_gst, 800);
  assert.equal(august.tax_method, 'instalment_rate');
  assert.equal(august.tax, 1500, '15% of gross business income of 10000');
  assert.equal(august.to_business_account, 800 + 1500 + 960);
  assert.equal(august.to_offset, 7040 - 1500);
  // The identity: money in is fully accounted for.
  assert.equal(august.received, august.to_business_account + august.to_offset + august.expenses_paid);
});

test('the split never leaks money, in any month or for the year, on either tax method', () => {
  ['instalment', 'cumulative'].forEach((mode) => {
    const built = splitContext({}, mode === 'instalment' ? { payg_instalment_rate: 0.155 } : {});
    built.actuals['2026-07-01'] = { g1_total_sales: 4400, gst_on_sales: 400, purchases: 1100, gst_on_purchases: 100 };
    built.actuals['2026-09-01'] = { g1_total_sales: 22000, gst_on_sales: 2000, purchases: 550, gst_on_purchases: 50 };
    built.actuals['2027-02-01'] = { g1_total_sales: 8800, gst_on_sales: 800, purchases: 3300, gst_on_purchases: 300 };
    // A month with costs but no income: net GST is a refund and the offset transfer goes negative.
    built.actuals['2027-05-01'] = { g1_total_sales: 0, gst_on_sales: 0, purchases: 1100, gst_on_purchases: 100 };
    const report = built.context.api_getMonthlyTransferSplit({ financial_year: 2026 });

    report.months.forEach((month) => {
      assert.equal(
        month.received,
        Math.round((month.to_business_account + month.to_offset + month.expenses_paid) * 100) / 100,
        mode + ': ' + month.period.from + ' must account for every dollar received'
      );
      assert.equal(month.to_offset, Math.round((month.taxable_profit - month.tax) * 100) / 100, mode + ': ' + month.period.from + ' offset equals after-tax income');
    });
    const may = report.months.filter((month) => month.calendar_month === 4)[0];
    assert.equal(may.net_gst, -100, 'a costs-only month yields a GST refund');
    assert.ok(may.to_offset < 0, 'and a negative offset transfer, rather than being silently clamped');
    assert.equal(report.totals.received, Math.round((report.totals.to_business_account + report.totals.to_offset + report.totals.expenses_paid) * 100) / 100);
    assert.equal(report.totals.to_offset, Math.round((report.totals.taxable_profit - report.totals.tax) * 100) / 100);
  });
});

test('transfer split warns when no instalment rate is configured and exports a reconciling csv', () => {
  const built = splitContext({});
  built.actuals['2026-08-01'] = { g1_total_sales: 11000, gst_on_sales: 1000, purchases: 0, gst_on_purchases: 0 };
  const report = built.context.api_getMonthlyTransferSplit({ financial_year: 2026 });
  assert.equal(report.tax_method, 'cumulative');
  assert.equal(report.instalment_rate, null);
  assert.ok(report.warnings.some((warning) => /payg_instalment_rate/.test(warning)));

  const exported = built.context.api_exportMonthlyTransferSplitCsv({ financial_year: 2026 });
  assert.equal(exported.filename, 'tempus-transfer-split-fy2026-27.csv');
  const lines = exported.csv.split('\r\n');
  assert.equal(lines.length, 14, 'header, twelve months, total');
  assert.match(lines[0], /To business account/);
  assert.match(lines[13], /^"Total"/);
});

test('a configured instalment rate is read as either a percentage or a decimal, else from the last BAS', () => {
  const context = taxContext();
  context.api_getSettings = () => ({ payg_instalment_rate: 15 });
  assert.equal(context.resolvePaygInstalmentRate_(), 0.15);
  context.api_getSettings = () => ({ payg_instalment_rate: 0.15 });
  assert.equal(context.resolvePaygInstalmentRate_(), 0.15);
  context.api_getSettings = () => ({});
  context.listBasSubmissionsInternal = () => [{ t2_instalment_rate: 0 }, { t2_instalment_rate: 12.5 }];
  assert.equal(context.resolvePaygInstalmentRate_(), 0.125, 'falls back to the most recent lodged T2 rate');
  context.listBasSubmissionsInternal = () => [];
  assert.equal(context.resolvePaygInstalmentRate_(), null, 'null puts the provision on the cumulative method');
});

test('bas monthly periods resolve inside the requested financial year, not the previous calendar year', () => {
  const { context } = createAppsScriptContext({});
  context.assertMigrationsSettled_ = () => true;
  load(context, 'backend/integrity.js');
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/entries.js');
  load(context, 'backend/financialReporting.js');
  // Jul-Dec sit in the FY's first calendar year, Jan-Jun in its second. Before the fix Jan-Jun silently
  // resolved 12 months early, making half of every financial year unreachable through the monthly path.
  // Values cross the vm realm boundary, so rehome them before comparing structurally.
  const range = (payload) => ({ ...context.basPeriodRange_(payload) });
  assert.deepStrictEqual(range({ financial_year: 2026, period_type: 'monthly', month: 6 }), { from: '2026-07-01', to: '2026-07-31' });
  assert.deepStrictEqual(range({ financial_year: 2026, period_type: 'monthly', month: 0 }), { from: '2027-01-01', to: '2027-01-31' });
  assert.deepStrictEqual(range({ financial_year: 2026, period_type: 'monthly', month: 5 }), { from: '2027-06-01', to: '2027-06-30' });
  // Every month of the year lands inside it, and the twelve are contiguous and non-overlapping.
  const ranges = [6, 7, 8, 9, 10, 11, 0, 1, 2, 3, 4, 5].map((month) => range({ financial_year: 2026, period_type: 'monthly', month: month }));
  assert.equal(ranges[0].from, '2026-07-01');
  assert.equal(ranges[11].to, '2027-06-30');
  ranges.forEach((range, index) => {
    assert.ok(range.from >= '2026-07-01' && range.to <= '2027-06-30', 'month index ' + index + ' is inside FY2027');
    if (index) assert.ok(range.from > ranges[index - 1].to, 'periods do not overlap');
  });
  // Quarterly behaviour is unchanged.
  assert.deepStrictEqual(range({ financial_year: 2026, period_type: 'quarterly', quarter: 1 }), { from: '2026-07-01', to: '2026-09-30' });
  assert.deepStrictEqual(range({ financial_year: 2026, period_type: 'quarterly', quarter: 3 }), { from: '2027-01-01', to: '2027-03-31' });
});

test('transfer split is wired into the bas page, gated on the assessments flag, and parses', () => {
  const markup = fs.readFileSync(path.join(root, 'views/partials/bas.html'), 'utf8');
  const client = fs.readFileSync(path.join(root, 'views/partials/operations-scripts.html'), 'utf8');
  const shell = fs.readFileSync(path.join(root, 'views/partials/scripts.html'), 'utf8');
  const styles = fs.readFileSync(path.join(root, 'views/partials/operations.html'), 'utf8');
  const served = fs.readFileSync(path.join(root, 'views/index.html'), 'utf8');

  // Section exists, starts hidden, and carries the table and export control.
  assert.match(markup, /id="bas-transfer-split-section"[^>]*display:none/);
  ['bas-transfer-split-body', 'bas-transfer-split-foot', 'bas-transfer-split-export', 'bas-transfer-split-method', 'bas-transfer-split-warnings']
    .forEach((id) => assert.ok(markup.includes('id="' + id + '"'), 'expected #' + id));
  assert.match(markup, /To business account[\s\S]*To offset/);

  // Gated on the assessments flag, and hidden rather than left populated when the flag is off.
  assert.match(client, /function transferSplitEnabled\(\)[\s\S]*enable_lil_assessments_mode/);
  assert.match(client, /if\(!transferSplitEnabled\(\)\)\{ section\.style\.display='none'; return; \}/);

  // Calls the real endpoints, and the export reuses the basis the table was built with.
  assert.ok(client.includes("'api_getMonthlyTransferSplit'"));
  assert.ok(client.includes("'api_exportMonthlyTransferSplitCsv'"));
  assert.match(client, /bas-transfer-split-export[\s\S]{0,600}operationsState\.transferSplitFy/);
  // Changing the accounting basis must refresh the split, not just the ledger card above it.
  assert.match(client, /bas-accounting-basis'\)\?\.addEventListener[\s\S]{0,300}renderTransferSplit\(\)/);
  // And the page render path invokes it.
  assert.match(shell, /typeof renderTransferSplit === 'function'\) renderTransferSplit\(\)/);

  // Negative transfers are styled, and the token they use is themed.
  assert.match(styles, /\.ts-amount--negative \{ color:var\(--danger\); \}/);
  assert.match(fs.readFileSync(path.join(root, 'views/partials/head.html'), 'utf8'), /--danger:/);

  // Both partials reach the served page, and the client script parses.
  assert.ok(served.includes("include('views/partials/bas')"));
  assert.ok(served.includes("include('views/partials/operations-scripts')"));
  const body = client.slice(client.indexOf('<script>') + 8, client.lastIndexOf('</script>'));
  new vm.Script(body, { filename: 'operations-scripts.html' });
});

test('every financial_year in the app means the same thing: the start year', () => {
  const shell = fs.readFileSync(path.join(root, 'views/partials/scripts.html'), 'utf8');
  const client = fs.readFileSync(path.join(root, 'views/partials/operations-scripts.html'), 'utf8');
  const markup = fs.readFileSync(path.join(root, 'views/partials/operations.html'), 'utf8');

  const { context } = createAppsScriptContext({});
  context.assertMigrationsSettled_ = () => true;
  load(context, 'backend/integrity.js');
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/entries.js');
  load(context, 'backend/tax.js');
  load(context, 'backend/financialReporting.js');

  // The client labels 2026 as "FY 2026-27" and enumerates its months starting July 2026. The backend
  // must resolve the same 2026 to the same span; a mismatch here is what blanked the BAS ledger card.
  assert.match(shell, /function financialYearLabel\(startYear\)[\s\S]*?endYear = startYear \+ 1/);
  assert.match(shell, /function financialYearStartYear\(date\)[\s\S]*?getFullYear\(\) - 1/);
  const span = { ...context.basPeriodRange_({ financial_year: 2026, period_type: 'quarterly', quarter: 1 }) };
  assert.equal(span.from, '2026-07-01', 'backend Q1 of 2026 starts July 2026, matching FY 2026-27');
  const lastMonth = { ...context.basPeriodRange_({ financial_year: 2026, period_type: 'monthly', month: 5 }) };
  assert.equal(lastMonth.to, '2027-06-30', 'and the year ends June 2027');

  // Tax tables key off the same number.
  assert.equal(context.api_getIndividualTaxTableStatus(2026).applied_financial_year, '2026-27');

  // No caller may reintroduce the end-year form. The old bug was a single expression that mixed both:
  // state.basCurrentFy (a start year) with a getFullYear()+1 fallback (an end year).
  assert.doesNotMatch(client, /basCurrentFy\)\s*\|\|\s*\(now\.getMonth\(\)>=6\?now\.getFullYear\(\)\+1/);
  assert.equal((client.match(/financialYearStartYear\(/g) || []).length, 3, 'both bas renders and the expense report derive the year the same way');
  assert.doesNotMatch(markup, /Financial year ending/);

  // Exported filenames spell the span out, so "fy2026" can never be read as the wrong half.
  assert.equal(context.financialYearLabelSlug_(2026), '2026-27');
});

function featureFlagContext(sheets) {
  const built = createAppsScriptContext(sheets);
  const { context } = built;
  context.assertMigrationsSettled_ = () => true;
  context.cacheGet = () => null;
  context.cacheSet = () => {};
  context.cacheClearPrefix = () => {};
  load(context, 'backend/integrity.js');
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/entries.js');
  load(context, 'backend/settings.js');
  load(context, 'backend/migrations.js');
  context.MIGRATION_CONTEXT.active = true;
  return built;
}
function flagState(spreadsheet, feature) {
  const rows = spreadsheet.getSheetByName('feature_flags').snapshot();
  const headers = rows[0];
  const match = rows.slice(1).filter((row) => String(row[headers.indexOf('feature')]).trim() === feature)[0];
  return match ? String(match[headers.indexOf('enabled')]) : null;
}

test('migrated expenses are revealed by enabling the ledger flag, so relocated data is never hidden', () => {
  const FLAG_HEADERS = ['feature', 'enabled', 'name', 'description'];
  const TX_HEADERS = ['id', 'vendor', 'vendor_abn', 'description', 'category', 'purchase_date', 'supplier_invoice_date', 'amount', 'gst_code', 'gst_amount', 'business_use_percentage', 'claimable_gst_confirmed', 'gst_override_amount', 'status', 'reconciliation_state', 'source_rule_id', 'source_occurrence_key', 'attachments_json', 'notes', 'created_at', 'updated_at'];
  const transaction = ['legacy-expense-1', 'Insurer', '', 'Insurance', 'insurance', '2026-07-01', '', 355.52, 'taxable', 32.32, 1, 'FALSE', '', 'legacy_unreconciled', 'legacy_unreconciled', 'legacy-rule-1', '2026-07-01', '[]', '', '', ''];

  // The exact state the earlier migration left behind: ledger rows present, flag off, page unreachable.
  const stranded = featureFlagContext({
    feature_flags: [FLAG_HEADERS, ['enable_company_tracking_features', 'TRUE', '', ''], ['enable_expenses', 'FALSE', '', '']],
    expense_transactions: [TX_HEADERS, transaction]
  });
  assert.equal(flagState(stranded.spreadsheet, 'enable_expenses'), 'FALSE');
  stranded.context.migrationRevealExpenseLedger_();
  assert.equal(flagState(stranded.spreadsheet, 'enable_expenses'), 'TRUE', 'the page that shows the rows is switched on');
  // Company tracking feeds the income calculation, so the migration must not touch it.
  assert.equal(flagState(stranded.spreadsheet, 'enable_company_tracking_features'), 'TRUE');
  stranded.context.migrationRevealExpenseLedger_();
  assert.equal(flagState(stranded.spreadsheet, 'enable_expenses'), 'TRUE', 'idempotent');

  // A flag row that does not exist yet is appended.
  const missingRow = featureFlagContext({
    feature_flags: [FLAG_HEADERS],
    expense_rules: [['id', 'vendor', 'amount', 'frequency', 'start_date', 'active'], ['rule-1', 'Insurer', 355.52, 'monthly', '2026-07-01', 'TRUE']]
  });
  missingRow.context.migrationRevealExpenseLedger_();
  assert.equal(flagState(missingRow.spreadsheet, 'enable_expenses'), 'TRUE');

  // No ledger data means nothing to reveal; a spreadsheet that never had company expenses is untouched.
  const empty = featureFlagContext({
    feature_flags: [FLAG_HEADERS, ['enable_expenses', 'FALSE', '', '']],
    expense_transactions: [TX_HEADERS]
  });
  empty.context.migrationRevealExpenseLedger_();
  assert.equal(flagState(empty.spreadsheet, 'enable_expenses'), 'FALSE');
});

test('enabling a flag from a migration never overwrites a deliberate off switch for another feature', () => {
  const FLAG_HEADERS = ['feature', 'enabled', 'name', 'description'];
  const built = featureFlagContext({
    feature_flags: [FLAG_HEADERS, ['enable_expenses', 'TRUE', '', ''], ['enable_invoices', 'FALSE', '', '']]
  });
  // Already on: reports no change rather than rewriting.
  assert.equal(built.context.migrationEnableFeatureFlag_('enable_expenses'), false);
  assert.equal(flagState(built.spreadsheet, 'enable_expenses'), 'TRUE');
  // And an unrelated feature the user switched off stays off.
  assert.equal(flagState(built.spreadsheet, 'enable_invoices'), 'FALSE');
  assert.equal(built.context.migrationEnableFeatureFlag_('enable_invoices'), true, 'only when explicitly asked');
  assert.equal(flagState(built.spreadsheet, 'enable_invoices'), 'TRUE');
});

const DEDUCTION_HEADERS = ['id', 'name', 'category_id', 'company_expense', 'deduction_type', 'amount_type', 'amount_value', 'gst_inclusive', 'gst_amount', 'frequency', 'start_date', 'end_date', 'notes', 'active', 'created_at', 'updated_at', 'display_order'];
function deductionRow(id, startDate, overrides) {
  const base = { id: id, name: 'Insurance ' + id, category_id: 'cat-1', company_expense: 'TRUE', deduction_type: 'standard', amount_type: 'flat', amount_value: 110, gst_inclusive: 'TRUE', gst_amount: 10, frequency: 'monthly', start_date: startDate, end_date: '', notes: '', active: 'TRUE', created_at: '2026-01-01T00:00:00Z', updated_at: '2026-01-01T00:00:00Z', display_order: 1 };
  Object.assign(base, overrides || {});
  return DEDUCTION_HEADERS.map((header) => base[header]);
}
function expenseMigrationContext(sheets) {
  const built = featureFlagContext(sheets);
  built.context.migrationToday_ = () => '2026-03-31';
  return built;
}
function sheetRows(spreadsheet, name) {
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) return { headers: [], rows: [] };
  const all = sheet.snapshot();
  const headers = all[0] || [];
  return { headers: headers, rows: all.slice(1).filter((row) => String(row[0] || '') !== '') };
}

test('migrated company expenses arrive paid and recorded, so they are usable on a cash basis', () => {
  const built = expenseMigrationContext({
    feature_flags: [['feature', 'enabled', 'name', 'description'], ['enable_expenses', 'FALSE', '', '']],
    deductions: [DEDUCTION_HEADERS, deductionRow('d-1', '2026-01-15')]
  });
  built.context.migrationCompanyExpenses_();

  const transactions = sheetRows(built.spreadsheet, 'expense_transactions');
  const payments = sheetRows(built.spreadsheet, 'expense_payments');
  assert.equal(transactions.rows.length, 3, 'Jan, Feb and Mar occurrences up to the migration date');
  transactions.rows.forEach((row) => {
    // 'recorded' is the lifecycle status; the review signal lives on reconciliation_state.
    assert.equal(row[transactions.headers.indexOf('status')], 'recorded');
    assert.equal(row[transactions.headers.indexOf('reconciliation_state')], 'legacy_unreconciled');
    // GST stays unconfirmed: a migration must not claim a credit against an unchecked tax invoice.
    assert.equal(String(row[transactions.headers.indexOf('claimable_gst_confirmed')]), 'FALSE');
  });
  assert.equal(payments.rows.length, 3, 'each occurrence carries the payment the legacy deduction asserted');
  const paidTotal = payments.rows.reduce((sum, row) => sum + Number(row[payments.headers.indexOf('amount')]), 0);
  assert.equal(paidTotal, 330);
  assert.deepStrictEqual(
    Array.from(payments.rows.map((row) => String(row[payments.headers.indexOf('payment_date')]))).sort(),
    ['2026-01-15', '2026-02-15', '2026-03-15']
  );
  // The source row is gone, its replacement exists, and the ledger page is switched on.
  assert.equal(sheetRows(built.spreadsheet, 'deductions').rows.length, 0);
  assert.equal(sheetRows(built.spreadsheet, 'expense_rules').rows.length, 1);
  assert.equal(flagState(built.spreadsheet, 'enable_expenses'), 'TRUE');

  // Re-running must not double-pay.
  built.context.migrationCompanyExpenses_();
  assert.equal(sheetRows(built.spreadsheet, 'expense_payments').rows.length, 3);
});

test('a company expense with no parseable start date is kept, not deleted for an inert rule', () => {
  const built = expenseMigrationContext({
    feature_flags: [['feature', 'enabled', 'name', 'description']],
    deductions: [DEDUCTION_HEADERS, deductionRow('d-keep', '', { created_at: '', updated_at: '' }), deductionRow('d-move', '2026-02-10')]
  });
  built.context.migrationCompanyExpenses_();
  const deductions = sheetRows(built.spreadsheet, 'deductions');
  assert.equal(deductions.rows.length, 1, 'the row with nothing usable to replace it survives');
  assert.equal(deductions.rows[0][deductions.headers.indexOf('id')], 'd-keep');
  assert.match(String(deductions.rows[0][deductions.headers.indexOf('notes')]), /no parseable start date/);
  // The one that migrated cleanly did leave.
  assert.ok(!deductions.rows.some((row) => row[deductions.headers.indexOf('id')] === 'd-move'));
});

test('a month-end company expense migrates on anchored dates and keeps an exception on its drifted date', () => {
  const built = expenseMigrationContext({
    deductions: [DEDUCTION_HEADERS, deductionRow('d-end', '2026-01-31'), deductionRow('d-skip', '2026-01-31')],
    // The Deductions page drifted after February, so it keyed March's occurrence to the 28th.
    deduction_occurrence_exceptions: [['id', 'deduction_id', 'original_date', 'exception_type', 'new_date', 'new_amount', 'notes'], ['x-1', 'd-skip', '2026-03-28', 'skip', '', '', '']]
  });
  built.context.migrationCompanyExpenses_();
  const transactions = sheetRows(built.spreadsheet, 'expense_transactions');
  const keysFor = (ruleId) => transactions.rows.filter((row) => row[transactions.headers.indexOf('source_rule_id')] === ruleId).map((row) => String(row[transactions.headers.indexOf('source_occurrence_key')]));
  assert.deepStrictEqual(Array.from(keysFor('legacy-rule-d-end')), ['2026-01-31', '2026-02-28', '2026-03-31']);
  assert.deepStrictEqual(Array.from(keysFor('legacy-rule-d-skip')), ['2026-01-31', '2026-02-28'], 'March was skipped on the Deductions page');
});

test('company expenses left in deductions after the ledger migration move on upgrade, exactly once', () => {
  const TX = ['id', 'vendor', 'vendor_abn', 'description', 'category', 'purchase_date', 'supplier_invoice_date', 'amount', 'gst_code', 'gst_amount', 'business_use_percentage', 'claimable_gst_confirmed', 'gst_override_amount', 'status', 'reconciliation_state', 'source_rule_id', 'source_occurrence_key', 'attachments_json', 'notes', 'created_at', 'updated_at'];
  const RULES = ['id', 'vendor', 'vendor_abn', 'description', 'category', 'amount', 'gst_code', 'gst_amount', 'business_use_percentage', 'claimable_gst_confirmed', 'frequency', 'start_date', 'end_date', 'active', 'notes', 'created_at', 'updated_at'];
  // d-old was already moved by the original migration (on the dates it drifted to) but its source row survived.
  const moved = (date) => ['legacy-expense-d-old-' + date, 'Insurance d-old', '', 'Insurance d-old', 'cat-1', date, date, 110, 'taxable', 10, 1, 'FALSE', '', 'recorded', 'legacy_unreconciled', 'legacy-rule-d-old', date, '[]', '', '', ''];
  const built = expenseMigrationContext({
    feature_flags: [['feature', 'enabled', 'name', 'description'], ['enable_expenses', 'FALSE', '', '']],
    deductions: [DEDUCTION_HEADERS, deductionRow('d-old', '2026-01-31'), deductionRow('d-late', '2026-02-10'), deductionRow('d-none', '', { created_at: '', updated_at: '' }), deductionRow('d-personal', '2026-02-01', { company_expense: 'FALSE' })],
    expense_rules: [RULES, ['legacy-rule-d-old', 'Insurance d-old', '', 'Insurance d-old', 'cat-1', 110, 'taxable', 10, 1, 'FALSE', 'monthly', '2026-01-31', '', 'TRUE', '', '', '']],
    expense_transactions: [TX, moved('2026-01-31'), moved('2026-02-28'), moved('2026-03-28')]
  });
  assert.ok(Array.from(built.context.listMigrations().map((m) => m.id)).includes('2026-09-late-company-expense-deductions'));
  const run = () => built.context.migrationLateCompanyExpenses_();
  const counts = () => ['deductions', 'expense_rules', 'expense_transactions', 'expense_payments', 'migration_archive'].map((name) => sheetRows(built.spreadsheet, name).rows.length);
  run();

  const deductions = sheetRows(built.spreadsheet, 'deductions');
  const ids = deductions.rows.map((row) => row[deductions.headers.indexOf('id')]);
  assert.deepStrictEqual(Array.from(ids).sort(), ['d-none', 'd-personal'], 'dated company expenses leave; the undated one and personal deductions stay');
  assert.match(String(deductions.rows.find((row) => row[0] === 'd-none')[deductions.headers.indexOf('notes')]), /no parseable start date/);

  const transactions = sheetRows(built.spreadsheet, 'expense_transactions');
  const keysFor = (ruleId) => transactions.rows.filter((row) => row[transactions.headers.indexOf('source_rule_id')] === ruleId).map((row) => String(row[transactions.headers.indexOf('source_occurrence_key')])).sort();
  assert.deepStrictEqual(Array.from(keysFor('legacy-rule-d-old')), ['2026-01-31', '2026-02-28', '2026-03-28'], 'nothing already moved is imported again');
  assert.deepStrictEqual(Array.from(keysFor('legacy-rule-d-late')), ['2026-02-10', '2026-03-10']);
  const rules = sheetRows(built.spreadsheet, 'expense_rules');
  assert.deepStrictEqual(Array.from(rules.rows.map((row) => row[0])).sort(), ['legacy-rule-d-late', 'legacy-rule-d-none', 'legacy-rule-d-old'], 'the undated row gets the same inert rule the original migration wrote');
  const payments = sheetRows(built.spreadsheet, 'expense_payments');
  assert.deepStrictEqual(Array.from(payments.rows.map((row) => String(row[payments.headers.indexOf('expense_transaction_id')]))).sort(), ['legacy-expense-d-late-2026-02-10', 'legacy-expense-d-late-2026-03-10'], 'the new occurrences arrive paid');
  const archive = sheetRows(built.spreadsheet, 'migration_archive');
  assert.deepStrictEqual(Array.from(archive.rows.map((row) => String(row[archive.headers.indexOf('migration_id')]))), ['2026-09-late-company-expense-deductions', '2026-09-late-company-expense-deductions']);
  assert.equal(flagState(built.spreadsheet, 'enable_expenses'), 'TRUE');

  const after = counts();
  run();
  assert.deepStrictEqual(counts(), after, 'idempotent');
});

test('repair migration makes already-migrated expenses usable without double-paying', () => {
  const TX_HEADERS = ['id', 'vendor', 'vendor_abn', 'description', 'category', 'purchase_date', 'supplier_invoice_date', 'amount', 'gst_code', 'gst_amount', 'business_use_percentage', 'claimable_gst_confirmed', 'gst_override_amount', 'status', 'reconciliation_state', 'source_rule_id', 'source_occurrence_key', 'attachments_json', 'notes', 'created_at', 'updated_at'];
  const legacy = (id, date) => ['legacy-expense-' + id, 'Insurer', '', 'Insurance', 'cat-1', date, date, 110, 'taxable', 10, 1, 'FALSE', '', 'legacy_unreconciled', 'legacy_unreconciled', 'legacy-rule-1', date, '[]', '', '', ''];
  const built = expenseMigrationContext({
    feature_flags: [['feature', 'enabled', 'name', 'description']],
    expense_transactions: [TX_HEADERS, legacy('a', '2026-01-15'), legacy('b', '2026-02-15')]
  });
  built.context.migrationUsableLegacyExpenses_();

  const transactions = sheetRows(built.spreadsheet, 'expense_transactions');
  transactions.rows.forEach((row) => assert.equal(row[transactions.headers.indexOf('status')], 'recorded'));
  transactions.rows.forEach((row) => assert.equal(row[transactions.headers.indexOf('reconciliation_state')], 'legacy_unreconciled', 'still needs review'));
  const payments = sheetRows(built.spreadsheet, 'expense_payments');
  assert.equal(payments.rows.length, 2);
  assert.equal(payments.rows.reduce((sum, row) => sum + Number(row[payments.headers.indexOf('amount')]), 0), 220);

  built.context.migrationUsableLegacyExpenses_();
  assert.equal(sheetRows(built.spreadsheet, 'expense_payments').rows.length, 2, 'idempotent');
});

test('missing expense rules are rebuilt from the transactions that reference them', () => {
  const TX = ['id', 'vendor', 'vendor_abn', 'description', 'category', 'purchase_date', 'supplier_invoice_date', 'amount', 'gst_code', 'gst_amount', 'business_use_percentage', 'claimable_gst_confirmed', 'gst_override_amount', 'status', 'reconciliation_state', 'source_rule_id', 'source_occurrence_key', 'attachments_json', 'notes', 'created_at', 'updated_at'];
  const tx = (id, name, date, amount, gstCode, ruleId) => [id, name, '', name, 'cat-1', date, date, amount, gstCode, gstCode === 'taxable' ? Math.round((amount / 11) * 100) / 100 : 0, 1, 'FALSE', '', 'recorded', 'legacy_unreconciled', ruleId, date, '[]', '', '', ''];
  const archivedDeduction = { id: 'd3', name: 'Psychology Insurance', category_id: 'cat-2', company_expense: 'TRUE', deduction_type: 'standard', amount_type: 'flat', amount_value: 308, gst_inclusive: 'TRUE', gst_amount: 28, frequency: 'monthly', start_date: '2026-08-08', end_date: '', notes: '', active: 'TRUE', created_at: '2026-08-08T00:00:00Z', updated_at: '2026-08-08T00:00:00Z' };

  // Exactly the reported state: transactions naming rules that were never written.
  const built = expenseMigrationContext({
    expense_transactions: [TX,
      tx('legacy-expense-d1-2026-08-08', 'Backpack Bed Donation', '2026-08-08', 10.56, 'out_of_scope', 'legacy-rule-d1'),
      tx('legacy-expense-d3-2026-08-08', 'Psychology Insurance', '2026-08-08', 308, 'taxable', 'legacy-rule-d3'),
      tx('legacy-expense-d3-2026-09-08', 'Psychology Insurance', '2026-09-08', 308, 'taxable', 'legacy-rule-d3')
    ],
    migration_archive: [['id', 'migration_id', 'source_sheet', 'source_row_json', 'reason', 'archived_at'],
      ['arch-1', '2026-08-company-expense-ledger', 'deductions', JSON.stringify(archivedDeduction), 'moved_to_expense_ledger', '2026-08-08T00:00:00Z']
    ]
  });
  assert.equal(sheetRows(built.spreadsheet, 'expense_rules').rows.length, 0);
  built.context.migrationRebuildMissingExpenseRules_();

  const rules = sheetRows(built.spreadsheet, 'expense_rules');
  assert.equal(rules.rows.length, 2, 'one rule per distinct source_rule_id');
  const byId = {};
  rules.rows.forEach((row) => { byId[row[rules.headers.indexOf('id')]] = row; });
  const field = (row, name) => row[rules.headers.indexOf(name)];

  // Archived deduction present: the schedule is restored faithfully and the rule is live.
  const restored = byId['legacy-rule-d3'];
  assert.equal(field(restored, 'frequency'), 'monthly');
  assert.equal(field(restored, 'start_date'), '2026-08-08');
  assert.equal(field(restored, 'amount'), 308);
  assert.equal(field(restored, 'gst_code'), 'taxable');
  assert.equal(String(field(restored, 'active')), 'TRUE');
  assert.match(String(field(restored, 'notes')), /Restored from the archived deduction/);

  // No archive: reconstructed from its transaction, and left archived so an inferred
  // schedule cannot start generating expenses on its own.
  const inferred = byId['legacy-rule-d1'];
  assert.equal(field(inferred, 'vendor'), 'Backpack Bed Donation');
  assert.equal(field(inferred, 'amount'), 10.56);
  assert.equal(field(inferred, 'gst_code'), 'out_of_scope');
  assert.equal(field(inferred, 'frequency'), 'once', 'a single occurrence is not assumed to recur');
  assert.equal(String(field(inferred, 'active')), 'FALSE');
  assert.match(String(field(inferred, 'notes')), /schedule is inferred/);

  built.context.migrationRebuildMissingExpenseRules_();
  assert.equal(sheetRows(built.spreadsheet, 'expense_rules').rows.length, 2, 'idempotent');
});

test('rule rebuild leaves a healthy ledger untouched', () => {
  const TX = ['id', 'vendor', 'purchase_date', 'amount', 'gst_code', 'gst_amount', 'business_use_percentage', 'status', 'reconciliation_state', 'source_rule_id'];
  const RULES = ['id', 'vendor', 'vendor_abn', 'description', 'category', 'amount', 'gst_code', 'gst_amount', 'business_use_percentage', 'claimable_gst_confirmed', 'frequency', 'start_date', 'end_date', 'active', 'notes', 'created_at', 'updated_at'];
  const built = expenseMigrationContext({
    expense_rules: [RULES, ['legacy-rule-d1', 'Insurer', '', 'Insurance', 'cat-1', 308, 'taxable', 28, 1, 'FALSE', 'monthly', '2026-08-08', '', 'TRUE', '', '', '']],
    expense_transactions: [TX, ['legacy-expense-d1-2026-08-08', 'Insurer', '2026-08-08', 308, 'taxable', 28, 1, 'recorded', 'legacy_unreconciled', 'legacy-rule-d1']]
  });
  built.context.migrationRebuildMissingExpenseRules_();
  const rules = sheetRows(built.spreadsheet, 'expense_rules');
  assert.equal(rules.rows.length, 1, 'no duplicate rule appended');
  assert.equal(String(rules.rows[0][rules.headers.indexOf('notes')]), '', 'the existing rule is not rewritten');
});

test('the reported three-deduction sheet migrates to rules, transactions and payments end to end', () => {
  // Verbatim from the live sheet: two monthly donations (no GST) and a one-off insurance premium
  // whose GST the user had already stated as 28.00.
  const rows = [
    ['a8efd4e0-c1a4-4bdf-8d0d-5fe04bd562cd', 'Backpack Bed Donation', '', 'TRUE', 'standard', 'flat', 10.56, 'FALSE', 0, 'monthly', '2026-08-08', '', '', 'TRUE', '2026-08-08T06:57:13Z', '2026-08-08T06:57:13Z', 1],
    ['7ddb6359-cc77-41fd-838a-b2f7d41edeb5', 'Share the Dignity Donation', '', 'TRUE', 'standard', 'flat', 5.28, 'FALSE', 0, 'monthly', '2026-08-08', '', '', 'TRUE', '2026-08-08T06:57:37Z', '2026-08-08T06:57:37Z', 2],
    ['7b27c85d-34d2-4772-a2c3-6dff6b923da6', 'Psychology Insurance', '', 'TRUE', 'standard', 'flat', 308, 'TRUE', 28, 'once', '2026-08-08', '', '', 'TRUE', '2026-08-08T06:57:59Z', '2026-08-08T06:57:59Z', 3]
  ];
  const built = expenseMigrationContext({ deductions: [DEDUCTION_HEADERS, ...rows] });
  built.context.migrationToday_ = () => '2026-08-23';
  built.context.migrationCompanyExpenses_();

  const rules = sheetRows(built.spreadsheet, 'expense_rules');
  const field = (set, row, name) => row[set.headers.indexOf(name)];
  assert.equal(rules.rows.length, 3, 'a rule per deduction, which is the row that went missing in the wild');
  assert.deepStrictEqual(
    Array.from(rules.rows.map((row) => field(rules, row, 'frequency'))),
    ['monthly', 'monthly', 'once'],
    'each schedule survives the move'
  );
  rules.rows.forEach((row) => assert.equal(String(field(rules, row, 'active')), 'TRUE'));

  const transactions = sheetRows(built.spreadsheet, 'expense_transactions');
  assert.equal(transactions.rows.length, 3, 'one occurrence each between 8 Aug and the migration date');
  const insurance = transactions.rows.filter((row) => field(transactions, row, 'vendor') === 'Psychology Insurance')[0];
  assert.equal(field(transactions, insurance, 'gst_code'), 'taxable');
  assert.equal(field(transactions, insurance, 'gst_amount'), 28, "the user's own GST figure is carried over, not recomputed");
  const donation = transactions.rows.filter((row) => field(transactions, row, 'vendor') === 'Backpack Bed Donation')[0];
  assert.equal(field(transactions, donation, 'gst_code'), 'out_of_scope', 'a donation claims no GST credit');

  const payments = sheetRows(built.spreadsheet, 'expense_payments');
  assert.equal(payments.rows.length, 3);
  assert.equal(payments.rows.reduce((sum, row) => sum + Number(field(payments, row, 'amount')), 0), 323.84);
  assert.equal(sheetRows(built.spreadsheet, 'deductions').rows.length, 0, 'and the source rows are released');
});

test('rules are rebuilt from the archive even when no transaction references them', () => {
  const archivedDeduction = (id, name, amount, frequency, gstInclusive, gstAmount) => ({
    id: id, name: name, category_id: '', company_expense: 'TRUE', deduction_type: 'standard', amount_type: 'flat',
    amount_value: amount, gst_inclusive: gstInclusive, gst_amount: gstAmount, frequency: frequency,
    start_date: '2026-08-08', end_date: '', notes: '', active: 'TRUE',
    created_at: '2026-08-08T06:57:13Z', updated_at: '2026-08-08T06:57:13Z'
  });
  const archiveRows = [
    archivedDeduction('a8efd4e0-c1a4-4bdf-8d0d-5fe04bd562cd', 'Backpack Bed Donation', 10.56, 'monthly', 'FALSE', 0),
    archivedDeduction('7ddb6359-cc77-41fd-838a-b2f7d41edeb5', 'Share the Dignity Donation', 5.28, 'monthly', 'FALSE', 0),
    archivedDeduction('7b27c85d-34d2-4772-a2c3-6dff6b923da6', 'Psychology Insurance', 308, 'once', 'TRUE', 28)
  ].map((record, index) => ['arch-' + index, '2026-08-company-expense-ledger', 'deductions', JSON.stringify(record), 'moved_to_expense_ledger', '2026-08-08T00:00:00Z']);

  // Transactions written by an older release: no source_rule_id at all, so nothing points at a rule.
  const TX = ['id', 'vendor', 'purchase_date', 'amount', 'gst_code', 'gst_amount', 'status', 'reconciliation_state'];
  const built = expenseMigrationContext({
    expense_transactions: [TX, ['legacy-expense-1', 'Backpack Bed Donation', '2026-08-08', 10.56, 'out_of_scope', 0, 'recorded', 'legacy_unreconciled']],
    migration_archive: [['id', 'migration_id', 'source_sheet', 'source_row_json', 'reason', 'archived_at'], ...archiveRows]
  });

  const before = built.context.api_getExpenseLedgerDiagnostic();
  assert.equal(before.expense_rules.rows, 0);
  assert.equal(before.transactions_with_source_rule_id, 0, 'nothing to key off on the transaction side');
  assert.equal(before.archived_company_expense_deductions, 3);
  assert.equal(before.rules_rebuildable, 3, 'but the archive still owes three rules');

  built.context.migrationRebuildMissingExpenseRules_();
  const rules = sheetRows(built.spreadsheet, 'expense_rules');
  assert.equal(rules.rows.length, 3);
  const field = (row, name) => row[rules.headers.indexOf(name)];
  assert.deepStrictEqual(Array.from(rules.rows.map((row) => field(row, 'frequency'))), ['monthly', 'monthly', 'once']);
  rules.rows.forEach((row) => assert.equal(String(field(row, 'active')), 'TRUE', 'archive-restored rules are live'));
  assert.equal(field(rules.rows.filter((row) => field(row, 'vendor') === 'Psychology Insurance')[0], 'gst_code'), 'taxable');

  const after = built.context.api_getExpenseLedgerDiagnostic();
  assert.equal(after.rules_rebuildable, 0);
  assert.equal(after.expense_rules.rows, 3);
  built.context.migrationRebuildMissingExpenseRules_();
  assert.equal(sheetRows(built.spreadsheet, 'expense_rules').rows.length, 3, 'idempotent');
});

test('the ledger diagnostic reports flags and applied migrations without creating sheets', () => {
  const built = expenseMigrationContext({
    feature_flags: [['feature', 'enabled', 'name', 'description'], ['enable_expenses', 'TRUE', '', ''], ['enable_company_tracking_features', 'FALSE', '', '']]
  });
  built.properties.setProperty('tempus_migrations_applied', JSON.stringify(['2026-08-company-expense-ledger']));
  const report = built.context.api_getExpenseLedgerDiagnostic();
  assert.equal(report.enable_expenses, true);
  assert.equal(report.enable_company_tracking_features, false);
  assert.deepStrictEqual(Array.from(report.applied_migrations), ['2026-08-company-expense-ledger']);
  // Nothing was conjured into existence just by asking.
  assert.equal(report.expense_rules.exists, false);
  assert.equal(report.expense_transactions.exists, false);
  assert.equal(built.spreadsheet.getSheetByName('expense_rules'), null);
});

test('the diagnostic can be written to a sheet while an upgrade is still pending', () => {
  const built = expenseMigrationContext({
    feature_flags: [['feature', 'enabled', 'name', 'description'], ['enable_expenses', 'FALSE', '', '']]
  });
  // The whole point of this path: it must work when getOrCreateSheet would refuse.
  built.context.MIGRATION_CONTEXT.active = false;
  built.context.assertMigrationsSettled_ = () => { throw new Error('upgrade_required'); };
  built.context.getOrCreateSheet = () => { throw new Error('upgrade_required'); };

  const result = built.context.writeExpenseLedgerDiagnostic();
  assert.match(String(result), /_diagnostic sheet/);
  const rows = built.spreadsheet.getSheetByName('_diagnostic').snapshot();
  assert.deepStrictEqual(Array.from(rows[0]), ['field', 'value']);
  const map = {};
  rows.slice(1).forEach((row) => { map[row[0]] = row[1]; });
  assert.equal(map['enable_expenses'], 'false');
  assert.equal(map['expense_rules.rows'], '0');
  // The harness seeds an applied-migration marker; the report must reflect the property, not invent it.
  assert.equal(map['applied_migrations'], 'tests-ready');
  built.properties.setProperty('tempus_migrations_applied', JSON.stringify([]));
  built.context.writeExpenseLedgerDiagnostic();
  const rerun = {};
  built.spreadsheet.getSheetByName('_diagnostic').snapshot().slice(1).forEach((row) => { rerun[row[0]] = row[1]; });
  assert.equal(rerun['applied_migrations'], '(none)');
  assert.ok(Object.prototype.hasOwnProperty.call(map, 'rules_rebuildable'));

  // Re-running replaces rather than appends.
  const before = built.spreadsheet.getSheetByName('_diagnostic').snapshot().length;
  built.context.writeExpenseLedgerDiagnostic();
  assert.equal(built.spreadsheet.getSheetByName('_diagnostic').snapshot().length, before);
});

test('the agenda week total honours the billable-only flag without touching the day rows', () => {
  const scripts = fs.readFileSync(path.join(root, 'views/partials/scripts.html'), 'utf8');

  // The flag exists, is grouped with the other time-entry flags, and is gated on hour types (with no
  // hour types every hour is billable, so the toggle would do nothing).
  assert.match(scripts, /weekly_hours_billable_only: \{\s*name: 'Weekly hours: count billable only'/);
  assert.equal((scripts.match(/if \(key === 'weekly_hours_billable_only' && !getFeatureFlag\('hour_types'\)\) return;/g) || []).length, 2,
    'both the settings workbench and the legacy flag list must hide the flag');

  // The week total switches source; the day rows keep reading di.minutes either way.
  assert.match(scripts, /weekMinutes \+= billableOnly \? di\.billableMinutes : di\.minutes;/);
  assert.match(scripts, /const billableOnly = getFeatureFlag\('weekly_hours_billable_only'\);/);
  assert.match(scripts, /const hasHours = info\.primary && info\.minutes > 0;/);

  const context = {
    state: {
      entries: [
        { date: '2026-09-01', hour_type_id: 'work-id', duration_minutes: 450 },
        { date: '2026-09-01', hour_type_id: 'leave-id', duration_minutes: 60 },
        { date: '2026-09-02', hour_type_id: 'leave-id', duration_minutes: 450 },
        { date: '2026-09-03', duration_minutes: 120 }
      ],
      hourTypeMap: {
        'work-id': { id: 'work-id', name: 'Work', contributes_to_income: true },
        'leave-id': { id: 'leave-id', name: 'Annual leave', contributes_to_income: false }
      },
      publicHolidayMap: {}
    },
    getDefaultHourTypeId: () => 'work-id',
    getFeatureFlag: () => true,
    hourTypeContributesToIncome: (id) => {
      const ht = context.state.hourTypeMap[id || 'work-id'];
      return !!(ht && ht.contributes_to_income);
    }
  };
  vm.createContext(context);
  vm.runInContext(extractClientFunction(scripts, 'agendaDayInfo'), context);

  // A mixed day reports both figures, so the week can narrow while the row still shows the full day.
  const mixed = context.agendaDayInfo('2026-09-01');
  assert.equal(mixed.minutes, 510);
  assert.equal(mixed.billableMinutes, 450);
  // A leave-only day drops out of a billable-only week entirely.
  const leaveOnly = context.agendaDayInfo('2026-09-02');
  assert.equal(leaveOnly.minutes, 450);
  assert.equal(leaveOnly.billableMinutes, 0);
  // An entry with no hour type falls back to the default type, exactly as `minutes` does.
  const untyped = context.agendaDayInfo('2026-09-03');
  assert.equal(untyped.billableMinutes, 120);
  // An empty day is 0, not NaN.
  assert.equal(context.agendaDayInfo('2026-09-04').billableMinutes, 0);
});

test('contract client-timesheet fields round-trip and survive a partial update', () => {
  const { context, spreadsheet } = createAppsScriptContext({});
  context.assertMigrationsSettled_ = () => true;
  withSheetApiStubs(context);
  load(context, 'backend/sheets.ts.js');
  load(context, 'backend/contracts.js');

  const created = context.api_addContract({
    name: 'Acme', start_date: '2026-07-01', hourly_rate: 150,
    specified_personnel: '  Geoff V  ',
    work_order_number: 'WO-12345',
    contract_reference: 'C-2026-088',
    timesheet_statement: 'Line one.\nLine two.'
  }).contract;
  assert.equal(created.specified_personnel, 'Geoff V', 'surrounding whitespace is trimmed');
  assert.equal(created.work_order_number, 'WO-12345');
  assert.equal(created.timesheet_statement, 'Line one.\nLine two.', 'the declaration keeps its line breaks');

  // Declared in the canonical schema, so getOrCreateSheet appends them by name — no migration id, and no
  // existing spreadsheet is gated behind an upgrade screen just to gain these columns.
  const headers = sheetHeaderRow(spreadsheet, 'contracts');
  ['specified_personnel', 'work_order_number', 'contract_reference', 'timesheet_statement']
    .forEach((column) => assert.ok(headers.indexOf(column) !== -1, 'missing contracts column: ' + column));

  const read = context.api_getContracts()[0];
  assert.equal(read.work_order_number, 'WO-12345');
  assert.equal(read.contract_reference, 'C-2026-088');

  // A partial update — the shape used when saving a line-item template or archiving — must not silently
  // blank the buyer-facing fields, or the next client timesheet prints without its work order.
  context.api_updateContract({ id: created.id, name: 'Acme', start_date: '2026-07-01', hourly_rate: 150 });
  const preserved = context.api_getContracts()[0];
  assert.equal(preserved.specified_personnel, 'Geoff V');
  assert.equal(preserved.work_order_number, 'WO-12345');
  assert.equal(preserved.timesheet_statement, 'Line one.\nLine two.');

  // An explicit empty string still clears a field (the form always sends all four).
  context.api_updateContract({ id: created.id, name: 'Acme', start_date: '2026-07-01', hourly_rate: 150, work_order_number: '' });
  assert.equal(context.api_getContracts()[0].work_order_number, '');
  assert.equal(context.api_getContracts()[0].contract_reference, 'C-2026-088', 'clearing one field leaves the others alone');
});

test('client timesheet is company-gated, contract-driven, and shares the print builders', () => {
  const navbar = fs.readFileSync(path.join(root, 'views/partials/navbar.html'), 'utf8');
  const dashboard = fs.readFileSync(path.join(root, 'views/partials/dashboard.html'), 'utf8');
  const contractsHtml = fs.readFileSync(path.join(root, 'views/partials/contracts.html'), 'utf8');
  const scripts = fs.readFileSync(path.join(root, 'views/partials/scripts.html'), 'utf8');

  assert.match(navbar, /id="calendar-client-timesheet-btn"/);
  assert.match(dashboard, /id="modal-client-timesheet"/);
  assert.match(dashboard, /id="client-timesheet-contract"/);
  // The client modal offers the same bulk "Unselect all" for non-income types as the internal print modal.
  assert.match(dashboard, /id="btn-client-timesheet-unselect-other"/);
  assert.match(scripts, /btnClientTimesheetUnselectOther\.addEventListener\('click', clientTimesheetUnselectAllOther\)/);
  // The internal print modal's own bulk control must stay wired.
  assert.match(dashboard, /id="btn-unselect-all-other"/);
  assert.match(scripts, /btnUnselectAllOther\.addEventListener\('click', unselectAllOtherHourTypes\)/);
  ['contract-specified-personnel', 'contract-work-order-number', 'contract-contract-reference', 'contract-timesheet-statement']
    .forEach((id) => assert.ok(contractsHtml.indexOf('id="' + id + '"') !== -1, 'contract form is missing ' + id));

  // Company tracking gates the button; it deliberately has no feature flag of its own.
  assert.match(scripts, /clientTimesheetBtn\.style\.display = visible \? 'inline-flex' : 'none'/);
  assert.match(scripts, /const visible = !!state\.companyTrackingEnabled;/);
  assert.ok(scripts.indexOf("'calendar-client-timesheet-btn'") !== -1, 'button must join the Tools disclosure list');

  // Both printed surfaces must run through the shared builders (defined once, called by each generator)
  // rather than keeping private copies of the calendar/totals/CSS that could drift apart.
  ['printCalendarHtml', 'printTotalsHtml', 'printSummaryHtml', 'printPaperCss', 'printCollectMonthTotals', 'printRoundedHours']
    .forEach((fn) => {
      const uses = (scripts.match(new RegExp(fn + '\\(', 'g')) || []).length;
      assert.ok(uses >= 3, fn + ' should be defined once and used by both print generators (found ' + uses + ')');
    });

  // Only entries booked to a DIFFERENT contract are excluded. Leave, sick days and public holidays are
  // stored with contract_id '' (see backend/publicholidays.js), so filtering on equality alone silently
  // erased every non-billable hour from the sheet.
  assert.match(scripts, /const entryContractId = String\(entry\.contract_id \|\| ''\);/);
  assert.match(scripts, /if \(contractId && entryContractId && entryContractId !== String\(contractId\)\) return;/);
  // The sign-off block is unconditional — the buyer signs whether or not a declaration is configured.
  assert.match(scripts, /BUYER’S REPRESENTATIVE/);
  assert.match(scripts, /\['Signature', 'Date', 'Printed name', 'Position'\]/);
  // The whole sheet has to land on one A4 page, so the client variant renders the calendar with the
  // compact tile metrics while the internal view keeps the roomy ones.
  assert.match(scripts, /function printTileMetrics\(compact\)/);
  assert.match(scripts, /printCalendarHtml\(year, month, selectedTypes, dailySummaries, incomeAccent, roundInterval, true\)/);
  assert.match(scripts, /printCalendarHtml\(year, month, selectedTypes, dailySummaries, incomeAccent, roundInterval, false\)/);
  // Identification is one horizontal band, and the contract's friendly name stays out of the printed body.
  assert.match(scripts, /class="ct-facts"/);
  assert.match(scripts, /\['Work order #', contract\.work_order_number\]/);
  assert.match(scripts, /\['Contract ref', contract\.contract_reference\]/);
});

require('./session-updates').run(test);
require('./sync-guard').run(test);
require('./expenses').run(test);
require('./recurrence-dst').run(test);

if (!process.exitCode) process.stdout.write('\n' + passed + ' tests passed.\n');

require('./lil-assessments').run(test);

// Async (promise-driven) mobile sync tests report through the same test() helper once they settle.
require('./mobile-sync').run(test);

require('./bas-ledger').run(test);
