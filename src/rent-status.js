const PAID = 'PAYÉ';
const UNPAID = 'IMPAYÉ';
const PARTIALLY_PAID = 'PARTIELLEMENT PAYÉ';

function getRentPaymentStatus(expectedAmount, paidAmount) {
  const expected = Number(expectedAmount);
  const paid = Number(paidAmount);
  if (!Number.isFinite(expected) || expected < 0 || !Number.isFinite(paid) || paid < 0) {
    throw new TypeError('Montants de loyer invalides.');
  }
  if (paid >= expected) return PAID;
  return paid > 0 ? PARTIALLY_PAID : UNPAID;
}

module.exports = { getRentPaymentStatus };
