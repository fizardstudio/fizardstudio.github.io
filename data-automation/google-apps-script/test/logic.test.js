// Run with:  node --test
const test = require('node:test');
const assert = require('node:assert');
const { cleanRow, dedupeById, summarize, parseDate_ } = require('../src/Logic.js');

const iso = (d) => d.toISOString().slice(0, 10);

test('dates in several formats, impossible dates rejected', () => {
  assert.equal(iso(parseDate_('2024-03-05')), '2024-03-05');
  assert.equal(iso(parseDate_('05/03/2024')), '2024-03-05');
  assert.equal(iso(parseDate_('5 Mar 2024')), '2024-03-05');
  assert.equal(iso(parseDate_('5 March 2024')), '2024-03-05');
  assert.equal(iso(parseDate_(new Date(2024, 2, 5))), '2024-03-05'); // local-midnight Date from Sheets
  assert.equal(parseDate_('31/02/2024'), null);
  assert.equal(parseDate_('soon'), null);
});

test('cleanRow normalises a messy but valid row', () => {
  const r = cleanRow([' ord-1 ', '05/03/2024', 'jane  DOE', 'Jane@Example.com ', 'Mug', '2', '$1,250.50', 'Done']);
  assert.deepEqual(r.errors, []);
  assert.equal(r.row[0], 'ORD-1');
  assert.equal(r.row[2], 'Jane Doe');
  assert.equal(r.row[3], 'jane@example.com');
  assert.equal(r.row[6], 1250.5);
  assert.equal(r.row[7], 'shipped');
});

test('cleanRow lists every problem', () => {
  const r = cleanRow(['', '31/02/2024', '', 'bad@', '', '1.5', 'abc', 'lost']);
  assert.deepEqual(r.errors, [
    'missing order id', 'invalid date', 'missing customer', 'invalid email',
    'missing product', 'qty must be a whole number > 0', 'invalid price', 'unknown status "lost"',
  ]);
});

test('dedupeById keeps the first occurrence', () => {
  const rows = [['A'], ['B'], ['A'], ['C'], ['B']];
  const d = dedupeById(rows);
  assert.deepEqual(d.rows, [['A'], ['B'], ['C']]);
  assert.deepEqual(d.removed, [{ index: 2, id: 'A' }, { index: 4, id: 'B' }]);
});

test('summarize skips cancelled/refunded and avoids float drift', () => {
  const rows = [
    ['1', '2024-01-10', 'X', '', 'Mug', '3', '0.1', 'paid'],
    ['2', '2024-01-20', 'Y', '', 'Mug', '1', '0.2', 'shipped'],
    ['3', '2024-02-01', 'Z', '', 'Cap', '1', '15', 'pending'],
    ['4', '2024-02-02', 'Z', '', 'Cap', '9', '15', 'cancelled'],
  ].map((r) => cleanRow(r).row);
  const s = summarize(rows);
  assert.equal(s.total, 15.5);
  assert.equal(s.orders, 3);
  assert.deepEqual(s.byMonth, { '2024-01': 0.5, '2024-02': 15 });
  assert.deepEqual(s.byProduct, { Mug: 0.5, Cap: 15 });
  assert.deepEqual(s.byStatus, { paid: 1, shipped: 1, pending: 1, cancelled: 1 });
});
