const path = require('node:path');
const ExcelJS = require('exceljs');
const {
  buildMonthlyBalanceFormulas,
  calculateMonthlyBalance,
} = require('./monthly-balance');
const {
  PRIORITY_CATEGORIES,
  resolvePriorityCategory,
} = require('./category-presentation');
const { getRentPaymentStatus } = require('./rent-status');

const EXPENSE_HEADER_GREEN = 'FF385623';
const EXPENSE_HEADER_GRAY = 'FF595959';
const REFERENCE_TEMPLATE = path.join(__dirname, '..', 'templates', 'depenses-maison-reference.xlsx');
const MONTH_NAMES = [
  'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin',
  'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre',
];
const FLOOR_ORDER = ['Rez-de-chaussée', '1er étage', '2e étage', '3e étage'];
const CURRENCY_FORMAT = '#,##0 [$CFA-280C]';
const MONTHLY_COLUMN_COUNT = 14;
const EXPENSE_HEADERS = ['Désignation', 'Montant'];
const normalize = (value) => String(value || '')
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g, '')
  .toLocaleLowerCase('fr')
  .trim();

function clone(value) {
  return value === undefined ? undefined : structuredClone(value);
}

function getSheetName(month) {
  return MONTH_NAMES[month - 1];
}

function withDe(monthLabel) {
  return /^[aeiouyàâäéèêëîïôöùûü]/i.test(monthLabel) ? `d'${monthLabel}` : `de ${monthLabel}`;
}

function getMonthKey(year, month) {
  const date = new Date(Date.UTC(year, month - 1, 1));
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}`;
}

function cloneWorksheet(workbook, source, name, columnCount = 7) {
  const sheet = workbook.addWorksheet(name, {
    properties: clone(source.properties),
    pageSetup: clone(source.pageSetup),
    views: clone(source.views),
    headerFooter: clone(source.headerFooter),
  });
  sheet.pageMargins = clone(source.pageMargins);
  sheet.state = source.state;
  sheet.autoFilter = clone(source.autoFilter);

  for (let column = 1; column <= columnCount; column += 1) {
    if (column > 7) continue;
    const sourceColumn = source.getColumn(column);
    const targetColumn = sheet.getColumn(column);
    if (sourceColumn.width !== undefined) targetColumn.width = sourceColumn.width;
    targetColumn.hidden = sourceColumn.hidden;
    targetColumn.outlineLevel = sourceColumn.outlineLevel;
    targetColumn.style = clone(sourceColumn.style);
  }

  for (let rowNumber = 1; rowNumber <= source.rowCount; rowNumber += 1) {
    const sourceRow = source.getRow(rowNumber);
    const targetRow = sheet.getRow(rowNumber);
    targetRow.height = sourceRow.height;
    targetRow.hidden = sourceRow.hidden;
    targetRow.outlineLevel = sourceRow.outlineLevel;
    targetRow.style = clone(sourceRow.style);
    for (let column = 1; column <= columnCount; column += 1) {
      const sourceCell = sourceRow.getCell(column);
      const targetCell = targetRow.getCell(column);
      if (column > 7) continue;
      targetCell.value = clone(sourceCell.value);
      targetCell.style = clone(sourceCell.style);
      targetCell.dataValidation = clone(sourceCell.dataValidation);
      targetCell.protection = clone(sourceCell.protection);
    }
  }

  source.model.merges.forEach((range) => sheet.mergeCells(range));
  return sheet;
}

function copyRowAppearance(sheet, sourceNumber, targetNumber, columnCount = 7) {
  const source = sheet.getRow(sourceNumber);
  const target = sheet.getRow(targetNumber);
  target.height = source.height;
  target.hidden = source.hidden;
  target.outlineLevel = source.outlineLevel;
  for (let column = 1; column <= columnCount; column += 1) {
    target.getCell(column).style = clone(source.getCell(column).style);
    target.getCell(column).dataValidation = clone(source.getCell(column).dataValidation);
    target.getCell(column).protection = clone(source.getCell(column).protection);
  }
}

function findExpenseTotalRow(sheet) {
  for (let rowNumber = sheet.rowCount; rowNumber >= 21; rowNumber -= 1) {
    if (normalize(sheet.getCell(rowNumber, 4).value) === 'total') return rowNumber;
  }
  throw new Error(`Le modèle de la feuille ${sheet.name} ne contient pas le total des dépenses.`);
}

function unmergeEditableRegions(sheet) {
  for (const range of [...sheet.model.merges]) {
    const [start, end = start] = range.split(':');
    const startRow = Number(start.match(/\d+/)?.[0]);
    const endRow = Number(end.match(/\d+/)?.[0]);
    if ((startRow >= 3 && startRow <= 11 && endRow <= 11) || range === 'D20:E20') {
      sheet.unMergeCells(range);
    }
  }
}

function setFormula(cell, formula, result) {
  cell.value = { formula, result: Number(result || 0) };
}

function categoryPriorityIndex(category) {
  const priorityCategory = resolvePriorityCategory(category);
  const index = PRIORITY_CATEGORIES.indexOf(priorityCategory);
  return index < 0 ? PRIORITY_CATEGORIES.length : index;
}

function groupExpensesByCategory(expenses, { categories, includeEmptyPriority = false } = {}) {
  const groups = new Map();
  const groupKey = (categoryId, category) => (
    categoryId === null || categoryId === undefined
      ? `name:${normalize(category)}`
      : `id:${categoryId}`
  );
  if (Array.isArray(categories)) {
    categories.forEach((category) => {
      const name = String(category.name || '').trim();
      if (!name) return;
      groups.set(groupKey(category.id, name), {
        category: name,
        categoryId: category.id,
        priorityIndex: categoryPriorityIndex(name),
        expenses: [],
        total: 0,
      });
    });
  } else if (includeEmptyPriority) {
    PRIORITY_CATEGORIES.forEach((category) => {
      groups.set(groupKey(null, category), {
        category,
        categoryId: null,
        priorityIndex: categoryPriorityIndex(category),
        expenses: [],
        total: 0,
      });
    });
  }
  for (const expense of expenses) {
    const rawCategory = String(expense.category || '').trim() || 'Autres dépenses';
    const category = Array.isArray(categories)
      ? rawCategory
      : resolvePriorityCategory(rawCategory) || rawCategory;
    let key = groupKey(expense.categoryId, category);
    if (!groups.has(key) && Array.isArray(categories)) {
      const currentCategory = categories.find(
        (candidate) => normalize(candidate.name) === normalize(category),
      );
      if (currentCategory) key = groupKey(currentCategory.id, currentCategory.name);
    }
    const group = groups.get(key) || {
      category,
      categoryId: expense.categoryId ?? null,
      priorityIndex: categoryPriorityIndex(category),
      expenses: [],
      total: 0,
    };
    group.expenses.push(expense);
    group.total += Number(expense.amount);
    groups.set(key, group);
  }
  return [...groups.values()].sort((left, right) => {
    if (left.priorityIndex !== right.priorityIndex) return left.priorityIndex - right.priorityIndex;
    return left.category.localeCompare(right.category, 'fr', { sensitivity: 'base' });
  });
}

function clearExpenseCellsAfterAmount(sheet, rowNumber, amountColumn) {
  for (let columnNumber = amountColumn + 1; columnNumber <= MONTHLY_COLUMN_COUNT; columnNumber += 1) {
    const cell = sheet.getCell(rowNumber, columnNumber);
    cell.value = null;
    cell.fill = { type: 'pattern', pattern: 'none' };
    cell.border = {};
  }
}

function styleCategoryTotalRow(sheet, rowNumber, amountColumn = 6) {
  for (let columnNumber = 4; columnNumber <= amountColumn; columnNumber += 1) {
    const cell = sheet.getCell(rowNumber, columnNumber);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE2F0E9' },
      bgColor: { argb: 'FFE2F0E9' },
    };
    cell.font = { ...cell.font, bold: true, color: { argb: 'FF264B47' } };
  }
  clearExpenseCellsAfterAmount(sheet, rowNumber, amountColumn);
}

function styleCategoryHeadingRow(sheet, rowNumber, category) {
  const heading = sheet.getCell(rowNumber, 4);
  heading.value = category;
  for (let columnNumber = 4; columnNumber <= 5; columnNumber += 1) {
    const cell = sheet.getCell(rowNumber, columnNumber);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: EXPENSE_HEADER_GREEN },
      bgColor: { argb: EXPENSE_HEADER_GREEN },
    };
    cell.font = { ...cell.font, bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFB8C6D9' } } };
  }
  sheet.getRow(rowNumber).height = 23;
  sheet.mergeCells(rowNumber, 4, rowNumber, 5);
  clearExpenseCellsAfterAmount(sheet, rowNumber, 5);
}

function styleExpenseSectionHeadingRow(sheet, rowNumber, title) {
  const heading = sheet.getCell(rowNumber, 4);
  heading.value = title;
  for (let columnNumber = 4; columnNumber <= MONTHLY_COLUMN_COUNT; columnNumber += 1) {
    const cell = sheet.getCell(rowNumber, columnNumber);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: EXPENSE_HEADER_GREEN },
      bgColor: { argb: EXPENSE_HEADER_GREEN },
    };
    cell.font = { ...cell.font, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFB8C6D9' } } };
  }
  sheet.getRow(rowNumber).height = 23;
  sheet.mergeCells(rowNumber, 4, rowNumber, MONTHLY_COLUMN_COUNT);
}

function styleExpenseHeaderRow(sheet, rowNumber, firstColumn = 5) {
  for (let columnNumber = 4; columnNumber <= MONTHLY_COLUMN_COUNT; columnNumber += 1) {
    sheet.getCell(rowNumber, columnNumber).value = null;
  }
  EXPENSE_HEADERS.forEach((header, index) => {
    const cell = sheet.getCell(rowNumber, index + firstColumn);
    cell.value = header;
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: EXPENSE_HEADER_GRAY },
      bgColor: { argb: EXPENSE_HEADER_GRAY },
    };
    cell.font = { ...cell.font, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFB8C6D9' } } };
  });
  clearExpenseCellsAfterAmount(sheet, rowNumber, firstColumn + EXPENSE_HEADERS.length - 1);
  sheet.getRow(rowNumber).height = 28;
}

function styleExpenseDataRow(sheet, rowNumber, amountColumn = 6, firstColumn = 5) {
  for (let columnNumber = firstColumn; columnNumber <= amountColumn; columnNumber += 1) {
    const cell = sheet.getCell(rowNumber, columnNumber);
    cell.fill = { type: 'pattern', pattern: 'none' };
    cell.border = { bottom: { style: 'hair', color: { argb: 'FFDDE4EA' } } };
    cell.alignment = { vertical: 'top', wrapText: true };
  }
  sheet.getCell(rowNumber, amountColumn).numFmt = CURRENCY_FORMAT;
  clearExpenseCellsAfterAmount(sheet, rowNumber, amountColumn);
}

function clearBlankExpenseSideCells(sheet, firstRow, lastRow) {
  for (let rowNumber = firstRow; rowNumber <= lastRow; rowNumber += 1) {
    for (const columnNumber of [1, 2]) {
      const cell = sheet.getCell(rowNumber, columnNumber);
      if (cell.value !== null && cell.value !== undefined) continue;
      cell.fill = { type: 'pattern', pattern: 'none' };
      cell.border = {};
    }
  }
}

function clearUnusedExpenseCells(sheet, firstRow, lastRow) {
  for (let rowNumber = firstRow; rowNumber <= lastRow; rowNumber += 1) {
    for (let columnNumber = 4; columnNumber <= MONTHLY_COLUMN_COUNT; columnNumber += 1) {
      const cell = sheet.getCell(rowNumber, columnNumber);
      cell.value = null;
      cell.fill = { type: 'pattern', pattern: 'none' };
      cell.border = {};
    }
  }
}

function applyStatusStyle(cell, status) {
  const styles = {
    PAYÉ: { fill: 'FFE2F0E9', color: 'FF264B47' },
    IMPAYÉ: { fill: 'FFF4CCCC', color: 'FF9C0006' },
    'PARTIELLEMENT PAYÉ': { fill: 'FFFFE699', color: 'FF7F6000' },
  };
  const style = styles[status];
  if (!style) return;
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: style.fill },
    bgColor: { argb: style.fill },
  };
  cell.font = { ...cell.font, bold: true, color: { argb: style.color } };
}

function growColumnToFit(sheet, columnNumber, values, maximumWidth) {
  const column = sheet.getColumn(columnNumber);
  const currentWidth = Number(column.width) || 9.14;
  const contentWidth = values.reduce((width, value) => {
    const longestLine = String(value || '').split(/\r?\n/).reduce(
      (longest, line) => Math.max(longest, line.length),
      0,
    );
    return Math.max(width, longestLine + 2);
  }, 14);
  column.width = Math.min(maximumWidth, Math.max(currentWidth, contentWidth));
}

function setWrappedContent(sheet, rowNumber, columnNumber, endColumnNumber = columnNumber) {
  const cell = sheet.getCell(rowNumber, columnNumber);
  cell.alignment = { ...cell.alignment, vertical: 'middle', wrapText: true };
  if (typeof cell.value !== 'string' || !cell.value) return;

  const width = Array.from(
    { length: endColumnNumber - columnNumber + 1 },
    (_, index) => Number(sheet.getColumn(columnNumber + index).width) || 9.14,
  ).reduce((sum, columnWidth) => sum + columnWidth, 0);
  const fontSize = Number(cell.font?.size) || 11;
  const charactersPerLine = Math.max(1, Math.floor(width * 11 / fontSize));
  const lineCount = cell.value.split(/\r?\n/).reduce(
    (count, line) => count + Math.max(1, Math.ceil(line.length / charactersPerLine)),
    0,
  );
  if (lineCount > 1) {
    const row = sheet.getRow(rowNumber);
    row.height = Math.max(Number(row.height) || 0, lineCount * fontSize * 1.25 + 4);
  }
}

function fitMonthlyColumns(sheet, data, expenseStartRow, expenseTotalRow, lastContentRow) {
  growColumnToFit(sheet, 1, ['NIVEAU', ...data.tenants.map((tenant) => tenant.floor)], 36);
  const financialValues = [
    data.pensionPere,
    data.pape,
    data.incomeTotal,
    data.openingBalance,
    data.closingBalance,
    data.savingsBalance,
    ...(data.otherIncomes || []).map((income) => income.amount),
  ].map((amount) => `${Number(amount ?? 0).toLocaleString('fr-FR')} CFA`);
  growColumnToFit(sheet, 2, financialValues, 24);
  sheet.getColumn(2).width = Math.max(sheet.getColumn(2).width, 18);
  const allExpenses = [...data.expenses, ...(data.protectedExpenses || [])];
  growColumnToFit(sheet, 4, [
    'TYPE DE LOGEMENT',
    'DATE',
    'DÉSIGNATION',
    ...data.tenants.map((tenant) => tenant.housingType),
    ...allExpenses.map((expense) => expense.title),
  ], 50);
  growColumnToFit(sheet, 5, [
    'NOMBRE DE PERSONNES',
    'MONTANT',
    ...allExpenses.map((expense) => expense.amount),
  ], 44);
  growColumnToFit(sheet, 6, [
    'MONTANT À PAYER',
    ...data.tenants.map((tenant) => tenant.expectedAmount),
  ], 22);
  growColumnToFit(sheet, 7, ['STATUT DU PAIEMENT', 'QUANTITÉ', ...data.tenants.map((tenant) => tenant.status)], 22);
  growColumnToFit(sheet, 8, ['UNITÉ', ...allExpenses.map((expense) => expense.unit)], 18);
  growColumnToFit(sheet, 9, ['PRIX UNITAIRE', ...allExpenses.map((expense) => expense.unitPrice)], 22);
  growColumnToFit(sheet, 10, ['BÉNÉFICIAIRE', ...allExpenses.map((expense) => expense.beneficiary)], 24);
  growColumnToFit(sheet, 11, ['PERSONNE AYANT PAYÉ', ...allExpenses.map((expense) => expense.paidBy)], 24);
  growColumnToFit(sheet, 12, ['MODE DE PAIEMENT', ...allExpenses.map((expense) => expense.paymentMethod)], 22);
  growColumnToFit(sheet, 13, ['RÉFÉRENCE', ...allExpenses.map((expense) => expense.reference)], 24);
  growColumnToFit(sheet, 14, ['COMMENTAIRE', ...allExpenses.map((expense) => expense.comment)], 44);

  const mergedNameWidth = (Number(sheet.getColumn(2).width) || 9.14)
    + (Number(sheet.getColumn(3).width) || 9.14);
  const nameWidth = Math.max(
    'NOM ET PRÉNOM'.length + 2,
    ...data.tenants.map((tenant) => String(tenant.name || '').length + 2),
  );
  if (nameWidth > mergedNameWidth) {
    const nameColumn = sheet.getColumn(2);
    nameColumn.width = Math.min(48, (Number(nameColumn.width) || 9.14) + nameWidth - mergedNameWidth);
  }

  for (const column of [1, 4, 5, 6, 7]) {
    for (let rowNumber = 2; rowNumber <= lastContentRow; rowNumber += 1) {
      setWrappedContent(sheet, rowNumber, column);
    }
  }
  setWrappedContent(sheet, 2, 2, 3);
  for (let index = 0; index < data.tenants.length; index += 1) {
    const rowNumber = index + 3;
    setWrappedContent(sheet, rowNumber, 2, 3);
  }
  for (let rowNumber = expenseStartRow; rowNumber < expenseTotalRow; rowNumber += 1) {
    for (let column = 4; column <= MONTHLY_COLUMN_COUNT; column += 1) {
      setWrappedContent(sheet, rowNumber, column);
    }
  }
}

function setMonthlyFinancialCell(sheet, styleSource, rowNumber, sourceRowNumber, value) {
  const labelCell = sheet.getCell(rowNumber, 1);
  const sourceLabelCell = styleSource.getCell(sourceRowNumber, 1);
  labelCell.value = sourceLabelCell.value;
  labelCell.style = clone(sourceLabelCell.style);

  const amountCell = sheet.getCell(rowNumber, 2);
  amountCell.value = Number(value ?? 0);
  amountCell.style = clone(styleSource.getCell(sourceRowNumber, 2).style);
  amountCell.numFmt = CURRENCY_FORMAT;
}

function fillMonthlySheet(sheet, data, styleSource, { includeCarryover }) {
  const templateTotalRow = findExpenseTotalRow(sheet);
  const tenantInsertions = Math.max(0, data.tenants.length - 8);
  const extraIncomeSources = data.otherIncomes || [];
  const incomeInsertions = Math.max(0, extraIncomeSources.length - 3);
  const categorizedExpenses = Array.isArray(data.categories) || data.monthKey >= '2026-09';
  const boutiqueExpenses = data.expenses.filter(
    (expense) => resolvePriorityCategory(expense.category) === 'Compte boutique',
  );
  const displayedExpenses = boutiqueExpenses.length
    ? [
      ...data.expenses.filter(
        (expense) => resolvePriorityCategory(expense.category) !== 'Compte boutique',
      ),
      {
        title: 'Compte boutique',
        amount: boutiqueExpenses.reduce((total, expense) => total + Number(expense.amount || 0), 0),
        category: 'Compte boutique',
      },
    ]
    : data.expenses;
  const expenseGroups = categorizedExpenses
    ? groupExpensesByCategory(displayedExpenses, {
      categories: data.categories,
      includeEmptyPriority: !Array.isArray(data.categories),
    }).filter((group) => group.expenses.length > 0)
    : [];
  const protectedExpenseGroups = categorizedExpenses
    ? groupExpensesByCategory(data.protectedExpenses || [], {
      categories: Array.isArray(data.categories) ? data.categories : undefined,
    }).filter((group) => group.expenses.length)
    : [];
  const groupRows = (groups) => groups.reduce((count, group) => count + group.expenses.length + 3, 0);
  const expenseRowCount = categorizedExpenses
    ? groupRows(expenseGroups)
    : displayedExpenses.length;
  const protectedExpenseRowCount = categorizedExpenses
    ? groupRows(protectedExpenseGroups)
    : (data.protectedExpenses || []).length;

  unmergeEditableRegions(sheet);
  if (tenantInsertions) {
    sheet.spliceRows(11, 0, ...Array.from({ length: tenantInsertions }, () => []));
  }

  const firstIncomeTitleRow = 20 + tenantInsertions;
  if (incomeInsertions) {
    sheet.spliceRows(firstIncomeTitleRow, 0, ...Array.from({ length: incomeInsertions }, () => []));
  }
  const rowShift = tenantInsertions + incomeInsertions;
  const expenseStartRow = (categorizedExpenses ? 21 : 22) + rowShift;
  const initialTotalRow = templateTotalRow + rowShift;
  const expenseCapacity = initialTotalRow - expenseStartRow;
  const expenseInsertions = Math.max(0, expenseRowCount - expenseCapacity);
  if (expenseInsertions) {
    sheet.spliceRows(initialTotalRow, 0, ...Array.from({ length: expenseInsertions }, () => []));
    for (let index = 0; index < expenseInsertions; index += 1) {
      const templateRow = expenseStartRow + (index % expenseCapacity);
      copyRowAppearance(sheet, templateRow, initialTotalRow + index, MONTHLY_COLUMN_COUNT);
    }
  }
  const expenseTotalRow = expenseStartRow + expenseRowCount;
  const lastTemplateExpenseRow = initialTotalRow + expenseInsertions;
  if (expenseTotalRow < lastTemplateExpenseRow) {
    clearUnusedExpenseCells(sheet, expenseTotalRow + 1, lastTemplateExpenseRow);
  }
  const protectedExpenses = data.protectedExpenses || [];
  const protectedSectionRow = protectedExpenseRowCount ? expenseTotalRow + 1 : null;
  const protectedTotalRow = protectedExpenseRowCount
    ? protectedSectionRow + protectedExpenseRowCount + 1
    : null;
  if (protectedExpenseRowCount) {
    const addedRows = Math.max(0, protectedTotalRow - lastTemplateExpenseRow);
    if (addedRows) {
      sheet.spliceRows(lastTemplateExpenseRow + 1, 0, ...Array.from({ length: addedRows }, () => []));
    }
    for (let index = 0; index < addedRows; index += 1) {
      copyRowAppearance(sheet, expenseStartRow, lastTemplateExpenseRow + 1 + index, MONTHLY_COLUMN_COUNT);
    }
  }

  [
    'NIVEAU',
    'NOM ET PRÉNOM',
    null,
    'TYPE DE LOGEMENT',
    'NOMBRE DE PERSONNES',
    'MONTANT À PAYER',
    'STATUT DU PAIEMENT',
  ].forEach((header, index) => {
    if (header !== null) sheet.getCell(2, index + 1).value = header;
  });

  for (let rowNumber = 3; rowNumber <= Math.max(10, data.tenants.length + 2); rowNumber += 1) {
    sheet.getCell(rowNumber, 7).value = null;
  }
  data.tenants.forEach((tenant, index) => {
    const rowNumber = index + 3;
    if (rowNumber > 10) copyRowAppearance(sheet, 3 + ((index) % 8), rowNumber);
    sheet.getCell(rowNumber, 1).value = null;
    sheet.getCell(rowNumber, 2).value = tenant.name;
    sheet.getCell(rowNumber, 4).value = tenant.housingType || '';
    sheet.getCell(rowNumber, 5).value = Number(tenant.peopleCount || 0);
    sheet.getCell(rowNumber, 6).value = Number(tenant.expectedAmount || 0);
    sheet.getCell(rowNumber, 6).numFmt = CURRENCY_FORMAT;
    const status = getRentPaymentStatus(tenant.expectedAmount, tenant.paidAmount);
    sheet.getCell(rowNumber, 7).value = status;
    applyStatusStyle(sheet.getCell(rowNumber, 7), status);
    sheet.mergeCells(rowNumber, 2, rowNumber, 3);
  });

  let index = 0;
  while (index < data.tenants.length) {
    const floor = data.tenants[index].floor;
    let end = index + 1;
    while (end < data.tenants.length && data.tenants[end].floor === floor) end += 1;
    const firstRow = index + 3;
    const lastRow = end + 2;
    sheet.getCell(firstRow, 1).value = floor;
    if (lastRow > firstRow) sheet.mergeCells(firstRow, 1, lastRow, 1);
    index = end;
  }

  const tenantTotalRow = 11 + tenantInsertions;
  sheet.mergeCells(tenantTotalRow, 2, tenantTotalRow, 4);
  sheet.getCell(tenantTotalRow, 2).value = 'Total';
  setFormula(
    sheet.getCell(tenantTotalRow, 5),
    data.tenants.length ? `SUM(E3:E${data.tenants.length + 2})` : '0',
    data.totalPeople,
  );
  setFormula(
    sheet.getCell(tenantTotalRow, 6),
    data.tenants.length ? `SUM(F3:F${data.tenants.length + 2})` : '0',
    data.rentExpected,
  );
  sheet.getCell(tenantTotalRow, 6).numFmt = CURRENCY_FORMAT;

  const paymentTotalRow = 13 + tenantInsertions;
  sheet.getCell(paymentTotalRow, 5).value = 'Total des paiements reçus';
  const cashFormula = data.rentPayments.length
    ? `SUM(${data.rentPayments.map((payment) => Number(payment.amount)).join(',')})`
    : '0';
  setFormula(sheet.getCell(paymentTotalRow, 6), cashFormula, data.rentReceived);
  sheet.getCell(paymentTotalRow, 6).numFmt = CURRENCY_FORMAT;

  const pensionRow = 15 + tenantInsertions;
  const papeRow = 17 + tenantInsertions;
  const totalIncomeRow = 21 + rowShift;
  const remainingRow = 22 + rowShift;
  const openingRow = 23 + rowShift;
  const resourcesRow = 24 + rowShift;
  const savingsRow = 30 + rowShift;
  setMonthlyFinancialCell(sheet, styleSource, pensionRow, 15, data.pensionPere);
  for (const columnNumber of [1, 2]) {
    const cell = sheet.getCell(pensionRow, columnNumber);
    cell.alignment = { ...cell.alignment, vertical: 'middle' };
  }
  setMonthlyFinancialCell(sheet, styleSource, papeRow, 17, data.pape);

  extraIncomeSources.forEach((income, extraIndex) => {
    const rowNumber = 18 + tenantInsertions + extraIndex;
    const styleSourceRow = 17 + tenantInsertions;
    sheet.getCell(rowNumber, 1).style = clone(sheet.getCell(styleSourceRow, 1).style);
    sheet.getCell(rowNumber, 2).style = clone(sheet.getCell(styleSourceRow, 2).style);
    sheet.getCell(rowNumber, 1).value = income.source;
    sheet.getCell(rowNumber, 2).value = Number(income.amount);
    sheet.getCell(rowNumber, 2).numFmt = CURRENCY_FORMAT;
  });

  const extraIncomeFormula = extraIncomeSources.length
    ? `+SUM(B${18 + tenantInsertions}:B${17 + tenantInsertions + extraIncomeSources.length})`
    : '';
  const totalFormula = `B${pensionRow}+B${papeRow}+F${paymentTotalRow}${extraIncomeFormula}`;
  setFormula(sheet.getCell(totalIncomeRow, 2), totalFormula, data.incomeTotal);
  sheet.getCell(totalIncomeRow, 2).numFmt = CURRENCY_FORMAT;
  const openingBalanceForDisplay = includeCarryover
    ? data.openingBalance
    : Number(data.openingBalanceWithoutCarryover ?? data.openingBalance);
  setMonthlyFinancialCell(sheet, styleSource, openingRow, 23, openingBalanceForDisplay);
  sheet.getRow(openingRow).hidden = false;
  setMonthlyFinancialCell(sheet, styleSource, resourcesRow, 21, data.resourcesAvailable);
  sheet.getCell(resourcesRow, 1).value = 'Total des ressources disponibles';
  sheet.getRow(resourcesRow).hidden = !includeCarryover;
  setMonthlyFinancialCell(sheet, styleSource, savingsRow, 30, data.savingsBalance);

  const categoryTotalRows = new Map();
  const protectedCategoryTotalRows = new Map();
  const groupKey = (group) => (
    group.categoryId === null || group.categoryId === undefined
      ? `name:${normalize(group.category)}`
      : `id:${group.categoryId}`
  );
  const writeExpense = (rowNumber, expense) => {
    const values = [
      expense.title || '',
      Number(expense.amount || 0),
    ];
    values.forEach((value, index) => {
      sheet.getCell(rowNumber, index + 4).value = value;
    });
    styleExpenseDataRow(sheet, rowNumber, 5, 4);
  };
  const renderExpenseGroups = (groups, startRow, totalsByCategory) => {
    let rowNumber = startRow;
    for (const group of groups) {
      styleCategoryHeadingRow(sheet, rowNumber, group.category);
      rowNumber += 1;
      styleExpenseHeaderRow(sheet, rowNumber, 4);
      rowNumber += 1;
      const firstExpenseRow = rowNumber;
      for (const expense of group.expenses) {
        writeExpense(rowNumber, expense);
        rowNumber += 1;
      }
      const totalRow = rowNumber;
      sheet.getCell(totalRow, 4).value = `Total ${group.category}`;
      if (group.expenses.length) {
        setFormula(
          sheet.getCell(totalRow, 5),
          `SUM(E${firstExpenseRow}:E${rowNumber - 1})`,
          group.total,
        );
      } else {
        sheet.getCell(totalRow, 5).value = 0;
      }
      sheet.getCell(totalRow, 5).numFmt = CURRENCY_FORMAT;
      styleCategoryTotalRow(sheet, totalRow, 5);
      totalsByCategory.set(groupKey(group), totalRow);
      rowNumber += 1;
    }
    return rowNumber;
  };

  if (categorizedExpenses) {
    renderExpenseGroups(expenseGroups, expenseStartRow, categoryTotalRows);
  } else {
    if (displayedExpenses.length) styleExpenseHeaderRow(sheet, expenseStartRow - 1, 4);
    displayedExpenses.forEach((expense, index) => {
      const rowNumber = expenseStartRow + index;
      sheet.getCell(rowNumber, 4).value = expense.title;
      sheet.getCell(rowNumber, 5).value = Number(expense.amount);
      sheet.getCell(rowNumber, 5).numFmt = CURRENCY_FORMAT;
      clearExpenseCellsAfterAmount(sheet, rowNumber, 5);
    });
  }
  sheet.getCell(expenseTotalRow, 4).value = 'Total';
  const currentExpenseTotalCell = sheet.getCell(expenseTotalRow, 5);
  const expensesFormula = categorizedExpenses
    ? `SUM(${[...categoryTotalRows.values()].map((rowNumber) => `E${rowNumber}`).join(',') || '0'})`
    : displayedExpenses.length
      ? `SUM(E${expenseStartRow}:E${expenseTotalRow - 1})`
      : '0';
  setFormula(currentExpenseTotalCell, expensesFormula, data.totalExpenses);
  currentExpenseTotalCell.numFmt = CURRENCY_FORMAT;
  styleCategoryTotalRow(sheet, expenseTotalRow, 5);

  if (protectedExpenses.length) {
    const sectionHeadingRow = protectedSectionRow;
    styleExpenseSectionHeadingRow(sheet, sectionHeadingRow, 'Dépenses financées par l’épargne protégée');
    let protectedRowsEnd;
    if (categorizedExpenses) {
      protectedRowsEnd = renderExpenseGroups(
        protectedExpenseGroups,
        sectionHeadingRow + 1,
        protectedCategoryTotalRows,
      );
    } else {
      protectedExpenses.forEach((expense, index) => {
        const rowNumber = sectionHeadingRow + 1 + index;
        sheet.getCell(rowNumber, 4).value = expense.title;
        sheet.getCell(rowNumber, 5).value = Number(expense.amount);
        sheet.getCell(rowNumber, 5).numFmt = CURRENCY_FORMAT;
        clearExpenseCellsAfterAmount(sheet, rowNumber, 5);
        setWrappedContent(sheet, rowNumber, 4);
      });
      protectedRowsEnd = sectionHeadingRow + 1 + protectedExpenses.length;
    }
    const totalLabel = sheet.getCell(protectedTotalRow, 4);
    totalLabel.value = 'Total réglé par l’épargne protégée';
    const protectedTotalCell = sheet.getCell(protectedTotalRow, 5);
    const protectedFormula = categorizedExpenses
      ? `SUM(${[...protectedCategoryTotalRows.values()].map((rowNumber) => `E${rowNumber}`).join(',') || '0'})`
      : `SUM(E${sectionHeadingRow + 1}:E${protectedRowsEnd - 1})`;
    setFormula(
      protectedTotalCell,
      protectedFormula,
      protectedExpenses.reduce((sum, expense) => sum + Number(expense.amount), 0),
    );
    protectedTotalCell.numFmt = CURRENCY_FORMAT;
    styleCategoryTotalRow(sheet, protectedTotalRow, 5);
  }

  let summaryTotalRow = protectedTotalRow || expenseTotalRow;
  if (categorizedExpenses) {
    const summaryTitleRow = summaryTotalRow + 2;
    const summaryHeaderRow = summaryTitleRow + 1;
    const summaryFirstRow = summaryHeaderRow + 1;
    const protectedGroupsByKey = new Map(protectedExpenseGroups.map((group) => [groupKey(group), group]));
    const mainGroupsByKey = new Map(expenseGroups.map((group) => [groupKey(group), group]));
    styleExpenseSectionHeadingRow(sheet, summaryTitleRow, 'Récapitulatif des dépenses par catégorie');
    ['Catégorie', 'Nombre de dépenses', 'Total'].forEach((header, index) => {
      const cell = sheet.getCell(summaryHeaderRow, index + 4);
      cell.value = header;
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: EXPENSE_HEADER_GREEN },
        bgColor: { argb: EXPENSE_HEADER_GREEN },
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFB8C6D9' } } };
      cell.alignment = { vertical: 'middle', wrapText: true };
    });
    let summaryRow = summaryFirstRow;
    for (const group of expenseGroups) {
      const key = groupKey(group);
      const protectedGroup = protectedGroupsByKey.get(key);
      const mainTotalRow = categoryTotalRows.get(key);
      const protectedTotalRowForCategory = protectedCategoryTotalRows.get(key);
      const totalCells = [
        mainTotalRow ? `E${mainTotalRow}` : null,
        protectedTotalRowForCategory ? `E${protectedTotalRowForCategory}` : null,
      ].filter(Boolean);
      sheet.getCell(summaryRow, 4).value = group.category;
      sheet.getCell(summaryRow, 5).value = group.expenses.length + (protectedGroup?.expenses.length || 0);
      setFormula(
        sheet.getCell(summaryRow, 6),
        totalCells.length ? `SUM(${totalCells.join(',')})` : '0',
        group.total + (protectedGroup?.total || 0),
      );
      sheet.getCell(summaryRow, 6).numFmt = CURRENCY_FORMAT;
      styleExpenseDataRow(sheet, summaryRow);
      summaryRow += 1;
    }
    summaryTotalRow = summaryRow;
    sheet.getCell(summaryTotalRow, 4).value = 'Total général des dépenses';
    sheet.getCell(summaryTotalRow, 5).value = displayedExpenses.length + protectedExpenses.length;
    const summaryRows = expenseGroups.length
      ? `SUM(F${summaryFirstRow}:F${summaryTotalRow - 1})`
      : '0';
    const grandTotal = data.totalExpenses
      + protectedExpenses.reduce((sum, expense) => sum + Number(expense.amount), 0);
    setFormula(sheet.getCell(summaryTotalRow, 6), summaryRows, grandTotal);
    sheet.getCell(summaryTotalRow, 6).numFmt = CURRENCY_FORMAT;
    styleCategoryTotalRow(sheet, summaryTotalRow);
  }
  clearBlankExpenseSideCells(
    sheet,
    expenseStartRow,
    protectedTotalRow || expenseTotalRow,
  );

  const balanceFormulas = buildMonthlyBalanceFormulas({
    incomeCell: `B${totalIncomeRow}`,
    openingCell: `B${openingRow}`,
    resourcesCell: `B${resourcesRow}`,
    expensesCell: `E${expenseTotalRow}`,
    includeCarryover,
  });
  const exportBalance = calculateMonthlyBalance({
    incomeTotal: data.incomeTotal,
    openingBalance: includeCarryover ? data.openingBalance : 0,
    totalExpenses: data.totalExpenses,
  });
  setFormula(
    sheet.getCell(resourcesRow, 2),
    balanceFormulas.resourcesAvailable,
    exportBalance.resourcesAvailable,
  );
  setFormula(
    sheet.getCell(remainingRow, 2),
    balanceFormulas.closingBalance,
    exportBalance.closingBalance,
  );
  sheet.getCell(remainingRow, 2).numFmt = CURRENCY_FORMAT;

  sheet.getCell(1, 1).value = `Loyers du mois ${withDe(data.previousMonthLabel)}`;
  sheet.getCell(20 + rowShift, 4).value = `Dépenses du mois ${withDe(data.monthLabel)}`;
  sheet.mergeCells(20 + rowShift, 4, 20 + rowShift, MONTHLY_COLUMN_COUNT);
  fitMonthlyColumns(
    sheet,
    data,
    expenseStartRow,
    expenseTotalRow,
    Math.max(expenseTotalRow, remainingRow, resourcesRow, savingsRow, protectedTotalRow || 0, summaryTotalRow),
  );
  for (let column = 1; column <= 7; column += 1) {
    if (column !== 3) setWrappedContent(sheet, 2, column, column === 2 ? 3 : column);
  }
  setWrappedContent(sheet, 20 + rowShift, 4, MONTHLY_COLUMN_COUNT);
  sheet.pageSetup.printArea = `A1:N${summaryTotalRow}`;
  sheet.pageSetup.orientation = 'landscape';
  sheet.pageSetup.fitToPage = true;
  sheet.pageSetup.fitToWidth = 1;
  sheet.pageSetup.fitToHeight = 0;

  return {
    exportBalance,
    name: sheet.name,
    tenantTotalRow,
    totalIncomeRow,
    remainingRow,
    openingRow,
    resourcesRow,
    savingsRow,
    paymentTotalRow,
    expenseTotalRow,
    expenseAmountColumn: 'E',
    protectedSectionRow,
    protectedTotalRow,
  };
}

function buildAnnualSummary(workbook, templateSheet, year, months) {
  const sheet = cloneWorksheet(workbook, templateSheet, 'Récapitulatif annuel');
  for (const range of [...sheet.model.merges]) {
    if (range !== 'A1:G1') sheet.unMergeCells(range);
  }
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    for (let column = 1; column <= 7; column += 1) sheet.getCell(rowNumber, column).value = null;
  }

  sheet.getCell('A1').value = `Récapitulatif annuel ${year}`;
  const headers = ['Mois', 'Recettes encaissées', '', 'Loyers attendus', 'Dépenses', 'Solde final', 'Épargne'];
  headers.forEach((header, index) => { sheet.getCell(2, index + 1).value = header; });
  sheet.mergeCells('B2:C2');
  for (let rowNumber = 3; rowNumber <= 14; rowNumber += 1) {
    copyRowAppearance(sheet, 22, rowNumber);
    sheet.mergeCells(rowNumber, 2, rowNumber, 3);
    const month = months[rowNumber - 3];
    const positions = month.positions;
    sheet.getCell(rowNumber, 1).value = month.monthLabel;
    setFormula(sheet.getCell(rowNumber, 2), `'${month.sheetName}'!B${positions.totalIncomeRow}`, month.incomeTotal);
    setFormula(sheet.getCell(rowNumber, 4), `'${month.sheetName}'!F${positions.tenantTotalRow}`, month.rentExpected);
    setFormula(
      sheet.getCell(rowNumber, 5),
      `'${month.sheetName}'!${positions.expenseAmountColumn}${positions.expenseTotalRow}`,
      month.totalExpenses,
    );
    setFormula(sheet.getCell(rowNumber, 6), `'${month.sheetName}'!B${positions.remainingRow}`, month.closingBalance);
    setFormula(sheet.getCell(rowNumber, 7), `'${month.sheetName}'!B${positions.savingsRow}`, month.savingsBalance);
    [2, 4, 5, 6, 7].forEach((column) => { sheet.getCell(rowNumber, column).numFmt = CURRENCY_FORMAT; });
  }

  copyRowAppearance(sheet, findExpenseTotalRow(templateSheet), 15);
  const annualValues = [
    months.reduce((sum, month) => sum + month.incomeTotal, 0),
    months.reduce((sum, month) => sum + month.rentExpected, 0),
    months.reduce((sum, month) => sum + month.totalExpenses, 0),
    months.at(-1)?.closingBalance || 0,
    months.at(-1)?.savingsBalance || 0,
  ];
  sheet.getCell(15, 1).value = 'Total annuel';
  setFormula(sheet.getCell(15, 2), 'SUM(B3:B14)', annualValues[0]);
  setFormula(sheet.getCell(15, 4), 'SUM(D3:D14)', annualValues[1]);
  setFormula(sheet.getCell(15, 5), 'SUM(E3:E14)', annualValues[2]);
  setFormula(sheet.getCell(15, 6), 'F14', annualValues[3]);
  setFormula(sheet.getCell(15, 7), 'G14', annualValues[4]);
  [2, 4, 5, 6, 7].forEach((column) => { sheet.getCell(15, column).numFmt = CURRENCY_FORMAT; });
  sheet.pageSetup.printArea = 'A1:G15';
}

async function buildReferenceWorkbook(months, { annual = false, includeCarryover = false } = {}) {
  if (!Array.isArray(months) || months.length === 0) {
    throw new Error('Sélectionnez au moins un mois à exporter.');
  }
  const templateWorkbook = new ExcelJS.Workbook();
  await templateWorkbook.xlsx.readFile(REFERENCE_TEMPLATE);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Gestion financière de la maison';
  workbook.created = new Date();
  workbook.calcProperties = { fullCalcOnLoad: true, forceFullCalc: true };

  const summaries = [];
  for (const data of months) {
    const monthNumber = Number(data.monthKey.slice(5, 7));
    const sheetName = getSheetName(monthNumber);
    const sourceName = sheetName === 'Août' ? 'Aout' : sheetName;
    const templateSheet = templateWorkbook.getWorksheet(sourceName)
      || templateWorkbook.getWorksheet('Janvier');
    const sheet = cloneWorksheet(workbook, templateSheet, sheetName, MONTHLY_COLUMN_COUNT);
    const { exportBalance, ...positions } = fillMonthlySheet(
      sheet,
      { ...data, monthLabel: sheetName },
      templateWorkbook.getWorksheet('Janvier'),
      { includeCarryover },
    );
    summaries.push({ ...data, ...exportBalance, sheetName, positions });
  }

  if (annual) {
    const source = templateWorkbook.getWorksheet('Janvier');
    buildAnnualSummary(workbook, source, Number(months[0].monthKey.slice(0, 4)), summaries);
  }
  return workbook;
}

module.exports = {
  MONTH_NAMES,
  REFERENCE_TEMPLATE,
  buildReferenceWorkbook,
  getMonthKey,
  getSheetName,
};
