// ===== Order Sheet Automation - single-file version (Logic.js + Code.js) =====
// Paste this whole file into Extensions > Apps Script > Code.gs

/**
 * Pure data logic - no SpreadsheetApp / MailApp calls in this file, so it can
 * be unit-tested in Node (see ../test) and reused in any Apps Script project.
 *
 * Row shape (array, in sheet column order):
 *   [order_id, order_date, customer, email, product, qty, unit_price, status]
 */

var COLUMNS = ['order_id', 'order_date', 'customer', 'email', 'product', 'qty', 'unit_price', 'status'];
var STATUSES = ['pending', 'paid', 'shipped', 'cancelled', 'refunded'];
var STATUS_ALIASES = { 'done': 'shipped', 'sent': 'shipped', 'canceled': 'cancelled', 'unpaid': 'pending', 'new': 'pending' };
var EMAIL_RE = /^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$/;

function squash_(v) {
  return String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim();
}

function titleCase_(s) {
  return squash_(s).toLowerCase().replace(/(^|[\s'-])(\S)/g, function (m, p, c) { return p + c.toUpperCase(); });
}

/** Accepts a Date, an ISO string, dd/mm/yyyy or "5 Mar 2024". Returns a Date or null. */
function parseDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    // Sheets hands back dates at local midnight; keep the calendar day, store as UTC
    return isNaN(v.getTime()) ? null : mkDate_(v.getFullYear(), v.getMonth() + 1, v.getDate());
  }
  var s = squash_(v);
  if (!s) return null;
  var m;
  if ((m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/))) return mkDate_(+m[1], +m[2], +m[3]);
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) return mkDate_(+m[3], +m[2], +m[1]);
  var months = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  if ((m = s.match(/^(\d{1,2}) ([A-Za-z]{3})[a-z]* (\d{4})$/)) && months[m[2].toLowerCase()]) {
    return mkDate_(+m[3], months[m[2].toLowerCase()], +m[1]);
  }
  return null;
}

function mkDate_(y, mo, d) {
  var dt = new Date(Date.UTC(y, mo - 1, d));
  // reject 31/02 etc. instead of letting JS roll it over to March
  return dt.getUTCFullYear() === y && dt.getUTCMonth() === mo - 1 && dt.getUTCDate() === d ? dt : null;
}

function parseNumber_(v) {
  if (typeof v === 'number') return v;
  var s = squash_(v).replace(/[^0-9.\-]/g, '');
  if (s === '' || s === '-' || s === '.') return null;
  var n = Number(s);
  return isNaN(n) ? null : n;
}

/**
 * Clean one row and list its problems.
 * @return {{row: Array, errors: string[]}}
 */
function cleanRow(raw) {
  var errors = [];
  var id = squash_(raw[0]).toUpperCase();
  var date = parseDate_(raw[1]);
  var customer = titleCase_(raw[2]);
  var email = squash_(raw[3]).toLowerCase();
  var product = squash_(raw[4]);
  var qty = parseNumber_(raw[5]);
  var price = parseNumber_(raw[6]);
  var status = squash_(raw[7]).toLowerCase();
  status = STATUS_ALIASES[status] || status;

  if (!id) errors.push('missing order id');
  if (!date) errors.push('invalid date');
  if (!customer) errors.push('missing customer');
  if (email && !EMAIL_RE.test(email)) errors.push('invalid email');
  if (!product) errors.push('missing product');
  if (qty === null || qty <= 0 || Math.floor(qty) !== qty) errors.push('qty must be a whole number > 0');
  if (price === null || price < 0) errors.push('invalid price');
  if (STATUSES.indexOf(status) === -1) errors.push('unknown status "' + status + '"');

  return { row: [id, date, customer, email, product, qty, price, status], errors: errors };
}

/**
 * Remove repeated order ids, keeping the first occurrence.
 * @return {{rows: Array[], removed: Array<{index:number, id:string}>}}
 */
function dedupeById(rows) {
  var seen = {};
  var keep = [];
  var removed = [];
  rows.forEach(function (r, i) {
    var id = r[0];
    if (id && seen[id]) { removed.push({ index: i, id: id }); return; }
    seen[id] = true;
    keep.push(r);
  });
  return { rows: keep, removed: removed };
}

/** Revenue summary of valid, non-cancelled orders. Totals are rounded to cents. */
function summarize(rows) {
  var byMonth = {}, byProduct = {}, byStatus = {};
  var total = 0, orders = 0;
  rows.forEach(function (r) {
    var status = r[7];
    byStatus[status] = (byStatus[status] || 0) + 1;
    if (status === 'cancelled' || status === 'refunded') return;
    var amount = Math.round(r[5] * r[6] * 100) / 100;
    var d = r[1];
    var month = d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2);
    byMonth[month] = round2_((byMonth[month] || 0) + amount);
    byProduct[r[4]] = round2_((byProduct[r[4]] || 0) + amount);
    total = round2_(total + amount);
    orders++;
  });
  return { total: total, orders: orders, byMonth: byMonth, byProduct: byProduct, byStatus: byStatus };
}

function round2_(n) { return Math.round(n * 100) / 100; }


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

  var text = DriveApp.getFileById(m[0]).getBlob().getDataAsString('UTF-8');
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);   // drop a UTF-8 byte-order mark
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
