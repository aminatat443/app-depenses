const test = require('node:test');
const assert = require('node:assert/strict');
const { getRentPaymentStatus } = require('../src/rent-status');

test('rent status requires collected payments to cover the obligation', () => {
  assert.equal(getRentPaymentStatus(50_000, 50_000), 'PAYÉ');
  assert.equal(getRentPaymentStatus(50_000, 0), 'IMPAYÉ');
  assert.equal(getRentPaymentStatus(50_000, 30_000), 'PARTIELLEMENT PAYÉ');
  assert.equal(getRentPaymentStatus(100_000, 40_000 + 60_000), 'PAYÉ');
  assert.equal(getRentPaymentStatus(50_000, 60_000), 'PAYÉ');
});

test('rent status rejects invalid amounts instead of guessing', () => {
  assert.throws(() => getRentPaymentStatus(Number.NaN, 0), /Montants de loyer invalides/);
  assert.throws(() => getRentPaymentStatus(50_000, -1), /Montants de loyer invalides/);
});
