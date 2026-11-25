function calculateMonthlyBalance({ incomeTotal, openingBalance, totalExpenses }) {
  const resourcesAvailable = Number(incomeTotal || 0) + Number(openingBalance || 0);
  const closingBalance = resourcesAvailable - Number(totalExpenses || 0);
  return { resourcesAvailable, closingBalance };
}

function buildMonthlyBalanceFormulas({
  incomeCell,
  openingCell,
  resourcesCell,
  expensesCell,
  includeCarryover,
}) {
  return {
    resourcesAvailable: includeCarryover ? `${incomeCell}+${openingCell}` : incomeCell,
    closingBalance: includeCarryover
      ? `${resourcesCell}-${expensesCell}`
      : `${incomeCell}-${expensesCell}`,
  };
}

module.exports = { calculateMonthlyBalance, buildMonthlyBalanceFormulas };
