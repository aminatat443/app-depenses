const PAID = 'PAYÉ';
const UNPAID = 'IMPAYÉ';

function getRentPaymentStatus(expectedAmount, paidAmount) {
  const expected = Number(expectedAmount);
  const paid = Number(paidAmount);
  if (!Number.isFinite(expected) || expected < 0 || !Number.isFinite(paid) || paid < 0) {
    throw new TypeError('Montants de loyer invalides.');
  }
  return paid >= expected ? PAID : UNPAID;
}

module.exports = { getRentPaymentStatus };
