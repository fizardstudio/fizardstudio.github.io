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

if (typeof module !== 'undefined') {
  module.exports = { cleanRow: cleanRow, dedupeById: dedupeById, summarize: summarize,
                     parseDate_: parseDate_, COLUMNS: COLUMNS };
}
