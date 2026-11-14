const test = require('node:test');
const assert = require('node:assert/strict');
const {
  calculateMonthlyBalance,
  buildMonthlyBalanceFormulas,
} = require('../src/monthly-balance');

test('monthly balance uses income plus one carryover minus recorded expenses', () => {
  const cases = [
    [680_000, 0, 630_000, 50_000],
    [680_000, 0, 600_000, 80_000],
    [760_000, 0, 650_000, 110_000],
    [500_000, 0, 550_000, -50_000],
    [680_000, 10_650, 630_000, 60_650],
  ];

  for (const [incomeTotal, openingBalance, totalExpenses, expectedClosingBalance] of cases) {
    assert.deepEqual(
      calculateMonthlyBalance({ incomeTotal, openingBalance, totalExpenses }),
      {
        resourcesAvailable: incomeTotal + openingBalance,
        closingBalance: expectedClosingBalance,
      },
    );
  }
});

test('Excel balance formulas use the displayed resources total', () => {
  assert.deepEqual(buildMonthlyBalanceFormulas({
    incomeCell: 'B21',
    openingCell: 'B23',
    resourcesCell: 'B24',
    expensesCell: 'E42',
    includeCarryover: true,
  }), {
    resourcesAvailable: 'B21+B23',
    closingBalance: 'B24-E42',
  });
  assert.deepEqual(buildMonthlyBalanceFormulas({
    incomeCell: 'B21',
    openingCell: 'B23',
    resourcesCell: 'B24',
    expensesCell: 'E42',
    includeCarryover: false,
  }), {
    resourcesAvailable: 'B21',
    closingBalance: 'B24-E42',
  });
});
