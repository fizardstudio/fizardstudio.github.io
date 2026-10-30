/**
 * Order Sheet Automation - Google Apps Script
 *
 * Sheets used (created automatically if missing):
 *   Raw      - paste or import the raw export here (row 1 = header)
 *   Orders   - cleaned, de-duplicated orders
 *   Errors   - rows that need a human, with the original row number and the reason
 *   Summary  - revenue by month / product / status + a chart
 *   Settings - B1: e-mail address for the report (blank = your own address)
 */

var SHEET = { RAW: 'Raw', ORDERS: 'Orders', ERRORS: 'Errors', SUMMARY: 'Summary', SETTINGS: 'Settings' };
var HEADER = ['Order ID', 'Order Date', 'Customer', 'Email', 'Product', 'Qty', 'Unit Price', 'Status'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Data Tools')
    .addItem('Import CSV from Google Drive...', 'importCsvFromDrive')
    .addItem('Clean & validate now', 'runCleaning')
    .addItem('Send report e-mail now', 'sendReport')
    .addSeparator()
    .addItem('Run every morning at 07:00', 'installDailyTrigger')
    .addItem('Stop the daily run', 'removeDailyTrigger')
    .addToUi();
}

function sheet_(name) {
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** Calendar-day UTC Date (from Logic.js) -> local Date so Sheets shows the same day. */
function toSheetDate_(d) {
  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}

/** Ask for a Drive file link or id, then load the CSV into the Raw sheet. */
function importCsvFromDrive() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Import CSV', 'Paste the Google Drive link or file id of the CSV:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var m = res.getResponseText().match(/[-\w]{25,}/);
  if (!m) { ui.alert('That does not look like a Drive file link or id.'); return; }

  var text = DriveApp.getFileById(m[0]).getBlob().getDataAsString('UTF-8').replace(/^﻿/, '');
  var rows = Utilities.parseCsv(text);
  if (rows.length < 2) { ui.alert('The file has no data rows.'); return; }

  var width = HEADER.length;
  rows = rows.map(function (r) { r = r.slice(0, width); while (r.length < width) r.push(''); return r; });
  var raw = sheet_(SHEET.RAW);
  raw.clearContents();
  // keep everything as text so Sheets does not "helpfully" re-interpret dates
  raw.getRange(1, 1, rows.length, width).setNumberFormat('@').setValues(rows);
  ui.alert('Imported ' + (rows.length - 1) + ' rows. Choose "Clean & validate now" next.');
}

/** Read Raw -> write Orders, Errors and Summary. Returns the summary object. */
function runCleaning() {
  var raw = sheet_(SHEET.RAW);
  var values = raw.getDataRange().getValues();
  if (values.length < 2) throw new Error('The Raw sheet is empty - import or paste data first.');

  var good = [], errors = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i].join('').trim() === '') continue;           // skip blank lines
    var src = values[i].slice(0, HEADER.length);
    while (src.length < HEADER.length) src.push('');
    var res = cleanRow(src);
    if (res.errors.length) errors.push([i + 1].concat(src, res.errors.join('; ')));
    else good.push(res.row);
  }

  var dd = dedupeById(good);
  dd.removed.forEach(function (r) {
    errors.push(['', r.id, '', '', '', '', '', '', '', 'duplicate order id (first occurrence kept)']);
  });

  writeTable_(sheet_(SHEET.ORDERS), HEADER, dd.rows.map(function (r) {
    return [r[0], toSheetDate_(r[1])].concat(r.slice(2));
  }), function (sh, n) {
    sh.getRange(2, 2, n, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, 7, n, 1).setNumberFormat('#,##0.00');
  });

  writeTable_(sheet_(SHEET.ERRORS), ['Raw row'].concat(HEADER, 'Problem'), errors, function (sh, n) {
    sh.getRange(2, HEADER.length + 2, n, 1).setFontColor('#b00020');
  });

  var summary = summarize(dd.rows);
  writeSummary_(summary, dd.rows.length, errors.length);
  SpreadsheetApp.getActive().toast(dd.rows.length + ' clean orders, ' + errors.length + ' to review', 'Done', 5);
  return summary;
}

function writeTable_(sh, header, rows, format) {
  sh.clearContents().clearFormats();
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    if (format) format(sh, rows.length);
  }
  sh.autoResizeColumns(1, header.length);
}

function writeSummary_(s, okCount, errCount) {
  var sh = sheet_(SHEET.SUMMARY);
  sh.clear();
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });

  var out = [
    ['Last run', new Date()],
    ['Clean orders', okCount],
    ['Rows needing review', errCount],
    ['Revenue (excl. cancelled/refunded)', s.total],
    ['', ''],
    ['Month', 'Revenue']
  ];
  var months = Object.keys(s.byMonth).sort();
  months.forEach(function (m) { out.push([m, s.byMonth[m]]); });
  sh.getRange(1, 1, out.length, 2).setValues(out);
  sh.getRange(6, 1, 1, 2).setFontWeight('bold');
  sh.getRange(1, 1, 4, 1).setFontWeight('bold');
  sh.getRange(4, 2).setNumberFormat('#,##0.00');
  if (months.length) sh.getRange(7, 2, months.length, 1).setNumberFormat('#,##0.00');

  var prod = Object.keys(s.byProduct).sort(function (a, b) { return s.byProduct[b] - s.byProduct[a]; })
    .map(function (p) { return [p, s.byProduct[p]]; });
  sh.getRange(6, 4, 1, 2).setValues([['Product', 'Revenue']]).setFontWeight('bold');
  if (prod.length) sh.getRange(7, 4, prod.length, 2).setValues(prod).offset(0, 1, prod.length, 1).setNumberFormat('#,##0.00');

  var st = Object.keys(s.byStatus).map(function (k) { return [k, s.byStatus[k]]; });
  sh.getRange(6, 7, 1, 2).setValues([['Status', 'Orders']]).setFontWeight('bold');
  if (st.length) sh.getRange(7, 7, st.length, 2).setValues(st);

  if (months.length) {
    sh.insertChart(sh.newChart().asColumnChart()
      .addRange(sh.getRange(6, 1, months.length + 1, 2))
      .setOption('title', 'Revenue by month').setOption('legend', { position: 'none' })
      .setPosition(2, 10, 0, 0).build());
  }
  sh.autoResizeColumns(1, 8);
}

/** E-mail a short HTML summary. */
function sendReport() {
  var s = runCleaning();
  var to = String(sheet_(SHEET.SETTINGS).getRange('B1').getValue() || Session.getEffectiveUser().getEmail());
  var ss = SpreadsheetApp.getActive();
  var errCount = Math.max(sheet_(SHEET.ERRORS).getLastRow() - 1, 0);
  var fmt = function (n) { return Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); };

  var rows = Object.keys(s.byProduct).sort(function (a, b) { return s.byProduct[b] - s.byProduct[a]; })
    .slice(0, 10)
    .map(function (p) { return '<tr><td>' + escape_(p) + '</td><td align="right">' + fmt(s.byProduct[p]) + '</td></tr>'; })
    .join('');
  var html = '<h2>Order report - ' + escape_(ss.getName()) + '</h2>' +
    '<p><b>' + s.orders + '</b> orders, revenue <b>' + fmt(s.total) + '</b>. ' +
    (errCount ? '<span style="color:#b00020">' + errCount + ' rows need review.</span>' : 'No rows need review.') + '</p>' +
    '<table border="1" cellpadding="4" cellspacing="0"><tr><th>Top products</th><th>Revenue</th></tr>' + rows + '</table>' +
    '<p><a href="' + ss.getUrl() + '">Open the spreadsheet</a></p>';
  MailApp.sendEmail({ to: to, subject: 'Order report: ' + s.orders + ' orders, ' + fmt(s.total), htmlBody: html });
}

function escape_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function dailyJob() { sendReport(); }

function installDailyTrigger() {
  removeDailyTrigger();
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getActive().toast('The report will run every day around 07:00.', 'Scheduled', 5);
}

function removeDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyJob') ScriptApp.deleteTrigger(t);
  });
}
