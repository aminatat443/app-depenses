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

const EXPENSE_HEADER_GREEN = 'FF264B47';
const EXPENSE_HEADER_GRAY = EXPENSE_HEADER_GREEN;
const GRAND_TOTAL_GREEN = 'FFD9EAD3';
const BODY_TEXT_GRAY = 'FF334155';
const GRID_GRAY = 'FFE2E8E2';
const RENT_AMOUNT_ORANGE = 'FFFA7D00';
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
    if ((startRow >= 3 && startRow <= 11 && endRow <= 11) ||
      range === 'D20:E20' || range === 'D20:F20') {
      sheet.unMergeCells(range);
    }
  }
}

function setFormula(cell, formula, result) {
  cell.value = { formula, result: Number(result || 0) };
}

function categoryPriorityIndex(category) {
  const normalized = normalize(category);
  let priorityCategory = PRIORITY_CATEGORIES.find(
    (candidate) => normalize(candidate) === normalized,
  );
  if (!priorityCategory && ![
    'reparations',
    'alimentation des animaux',
    'factures',
    'autres depenses',
    'depenses exceptionnelles',
    'depenses imprevues',
  ].includes(normalized)) {
    priorityCategory = resolvePriorityCategory(category);
  }
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
    const normalizedRawCategory = normalize(rawCategory);
    const isPreservedHistoricalCategory = [
      'reparations',
      'alimentation des animaux',
      'factures',
      'autres depenses',
      'depenses exceptionnelles',
      'depenses imprevues',
    ].includes(normalizedRawCategory);
    const category = !Array.isArray(categories) && !isPreservedHistoricalCategory
      ? resolvePriorityCategory(rawCategory) || rawCategory
      : rawCategory;
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
      fgColor: { argb: GRAND_TOTAL_GREEN },
      bgColor: { argb: GRAND_TOTAL_GREEN },
    };
    cell.font = {
      ...cell.font,
      name: 'Aptos',
      bold: true,
      color: { argb: 'FF274E13' },
    };
    cell.alignment = {
      ...cell.alignment,
      horizontal: columnNumber === amountColumn ? 'right' : 'left',
      vertical: 'middle',
    };
    cell.border = {
      ...cell.border,
      top: { style: 'thin', color: { argb: EXPENSE_HEADER_GREEN } },
    };
  }
  clearExpenseCellsAfterAmount(sheet, rowNumber, amountColumn);
}

function styleExpenseGrandTotalRow(sheet, rowNumber, amountColumn = 5, firstColumn = 4) {
  for (let columnNumber = firstColumn; columnNumber <= amountColumn; columnNumber += 1) {
    const cell = sheet.getCell(rowNumber, columnNumber);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: GRAND_TOTAL_GREEN },
      bgColor: { argb: GRAND_TOTAL_GREEN },
    };
    cell.font = { ...cell.font, name: 'Aptos', bold: true, color: { argb: 'FF274E13' } };
    cell.alignment = {
      ...cell.alignment,
      horizontal: columnNumber === amountColumn ? 'right' : 'left',
      vertical: 'middle',
    };
    cell.border = {
      ...cell.border,
      top: { style: 'medium', color: { argb: EXPENSE_HEADER_GREEN } },
      bottom: { style: 'thin', color: { argb: EXPENSE_HEADER_GREEN } },
    };
  }
  sheet.getRow(rowNumber).height = Math.max(Number(sheet.getRow(rowNumber).height) || 0, 21);
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
    cell.font = { ...cell.font, name: 'Aptos', bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFB8C6D9' } } };
  }
  sheet.getRow(rowNumber).height = 23;
  sheet.mergeCells(rowNumber, 4, rowNumber, 5);
  clearExpenseCellsAfterAmount(sheet, rowNumber, 5);
}

function styleExpenseSectionHeadingRow(
  sheet,
  rowNumber,
  title,
  endColumnNumber = 5,
  startColumnNumber = 4,
) {
  const heading = sheet.getCell(rowNumber, startColumnNumber);
  heading.value = title;
  for (let columnNumber = startColumnNumber; columnNumber <= endColumnNumber; columnNumber += 1) {
    const cell = sheet.getCell(rowNumber, columnNumber);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: EXPENSE_HEADER_GREEN },
      bgColor: { argb: EXPENSE_HEADER_GREEN },
    };
    cell.font = { ...cell.font, name: 'Aptos', bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFB8C6D9' } } };
  }
  sheet.getRow(rowNumber).height = 25;
  sheet.mergeCells(rowNumber, startColumnNumber, rowNumber, endColumnNumber);
  for (let columnNumber = endColumnNumber + 1; columnNumber <= MONTHLY_COLUMN_COUNT; columnNumber += 1) {
    const cell = sheet.getCell(rowNumber, columnNumber);
    cell.value = null;
    cell.fill = { type: 'pattern', pattern: 'none' };
    cell.border = {};
  }
}

function styleFinancialSummaryHeading(sheet, rowNumber) {
  sheet.mergeCells(rowNumber, 1, rowNumber, 2);
  const heading = sheet.getCell(rowNumber, 1);
  heading.value = 'RÉCAPITULATIF FINANCIER';
  for (let columnNumber = 1; columnNumber <= 2; columnNumber += 1) {
    const cell = sheet.getCell(rowNumber, columnNumber);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: EXPENSE_HEADER_GREEN },
      bgColor: { argb: EXPENSE_HEADER_GREEN },
    };
    cell.font = { ...cell.font, name: 'Aptos', bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    const border = { bottom: { style: 'thin', color: { argb: EXPENSE_HEADER_GREEN } } };
    if (columnNumber === 1) border.left = { style: 'thin', color: { argb: EXPENSE_HEADER_GREEN } };
    if (columnNumber === 2) border.right = { style: 'thin', color: { argb: EXPENSE_HEADER_GREEN } };
    cell.border = border;
  }
  sheet.getRow(rowNumber).height = 24;
}

function styleFinancialSummaryRows(sheet, rowNumbers, highlightedRows) {
  const highlighted = new Set(highlightedRows);
  for (const rowNumber of rowNumbers) {
    const highlight = highlighted.has(rowNumber);
    for (let columnNumber = 1; columnNumber <= 2; columnNumber += 1) {
      const cell = sheet.getCell(rowNumber, columnNumber);
      if (highlight) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: GRAND_TOTAL_GREEN },
          bgColor: { argb: GRAND_TOTAL_GREEN },
        };
        cell.font = { ...cell.font, name: 'Aptos', bold: true, color: { argb: 'FF274E13' } };
      } else {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
      }
      cell.font = {
        ...cell.font,
        bold: highlight,
        size: highlight ? 11 : 10,
        name: 'Aptos',
        color: { argb: highlight ? 'FF274E13' : BODY_TEXT_GRAY },
      };
      cell.alignment = {
        ...cell.alignment,
        horizontal: columnNumber === 2 ? 'right' : 'left',
        vertical: 'middle',
        wrapText: true,
      };
      const border = {
        bottom: { style: 'thin', color: { argb: GRID_GRAY } },
        ...(highlight ? { top: { style: 'thin', color: { argb: EXPENSE_HEADER_GREEN } } } : {}),
      };
      if (columnNumber === 1) border.left = { style: 'thin', color: { argb: 'FFD7E1D2' } };
      if (columnNumber === 2) border.right = { style: 'thin', color: { argb: 'FFD7E1D2' } };
      cell.border = border;
    }
    sheet.getRow(rowNumber).height = Math.max(Number(sheet.getRow(rowNumber).height) || 0, 19);
  }
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
    cell.font = { ...cell.font, name: 'Aptos', bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFB8C6D9' } } };
  });
  clearExpenseCellsAfterAmount(sheet, rowNumber, firstColumn + EXPENSE_HEADERS.length - 1);
  sheet.getRow(rowNumber).height = 28;
}

function styleExpenseDataRow(sheet, rowNumber, amountColumn = 6, firstColumn = 5) {
  for (let columnNumber = firstColumn; columnNumber <= amountColumn; columnNumber += 1) {
    const cell = sheet.getCell(rowNumber, columnNumber);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: rowNumber % 2 ? 'FFFFFFFF' : 'FFF4F8F4' },
      bgColor: { argb: rowNumber % 2 ? 'FFFFFFFF' : 'FFF4F8F4' },
    };
    cell.font = { ...cell.font, name: 'Aptos', size: 10, color: { argb: BODY_TEXT_GRAY } };
    cell.border = { bottom: { style: 'thin', color: { argb: GRID_GRAY } } };
    cell.alignment = { vertical: 'top', wrapText: true };
  }
  const amountCell = sheet.getCell(rowNumber, amountColumn);
  amountCell.numFmt = CURRENCY_FORMAT;
  amountCell.alignment = { ...amountCell.alignment, horizontal: 'right', vertical: 'middle' };
  clearExpenseCellsAfterAmount(sheet, rowNumber, amountColumn);
}

function styleMonthlyTables(sheet, {
  tenantCount,
  tenantTotalRow,
  paymentTotalRow,
  expenseTitleRow,
  expenseHeaderRow,
  expenseStartRow,
  expenseDataEndRow,
  expenseTotalRow,
}) {
  const thinBorder = {
    top: { style: 'thin', color: { argb: GRID_GRAY } },
    bottom: { style: 'thin', color: { argb: GRID_GRAY } },
    left: { style: 'thin', color: { argb: GRID_GRAY } },
    right: { style: 'thin', color: { argb: GRID_GRAY } },
  };
  const setFill = (cell, color) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: color },
      bgColor: { argb: color },
    };
  };

  for (let column = 1; column <= 7; column += 1) {
    const title = sheet.getCell(1, column);
    setFill(title, EXPENSE_HEADER_GREEN);
    title.font = { ...title.font, name: 'Aptos', bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
    title.alignment = { ...title.alignment, horizontal: 'center', vertical: 'middle', wrapText: true };
    title.border = thinBorder;
    const header = sheet.getCell(2, column);
    setFill(header, EXPENSE_HEADER_GREEN);
    header.font = { ...header.font, name: 'Aptos', bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    header.alignment = { ...header.alignment, horizontal: 'center', vertical: 'middle', wrapText: true };
    header.border = thinBorder;
  }
  sheet.getRow(1).height = Math.max(Number(sheet.getRow(1).height) || 0, 27);
  sheet.getRow(2).height = Math.max(Number(sheet.getRow(2).height) || 0, 27);

  for (let rowNumber = 3; rowNumber <= tenantCount + 2; rowNumber += 1) {
    const fill = (rowNumber - 3) % 2 ? 'FFFFFFFF' : 'FFF4F8F4';
    for (let column = 1; column <= 7; column += 1) {
      const cell = sheet.getCell(rowNumber, column);
      if (column !== 7) setFill(cell, fill);
      cell.font = {
        ...cell.font,
        name: 'Aptos',
        size: 10,
        ...(column === 7 ? {} : { color: { argb: BODY_TEXT_GRAY } }),
      };
      cell.alignment = {
        ...cell.alignment,
        horizontal: column === 1 ? 'center' : column === 6 ? 'right' : 'left',
        vertical: 'middle',
        wrapText: true,
      };
      cell.border = thinBorder;
    }
  }
  for (const rowNumber of [tenantTotalRow, paymentTotalRow]) {
    for (let column = 1; column <= 7; column += 1) {
      const cell = sheet.getCell(rowNumber, column);
      if (cell.value === null || cell.value === undefined) continue;
      setFill(cell, GRAND_TOTAL_GREEN);
      cell.font = { ...cell.font, name: 'Aptos', bold: true, color: { argb: 'FF274E13' } };
      cell.alignment = { ...cell.alignment, vertical: 'middle' };
      cell.border = {
        ...thinBorder,
        top: { style: 'thin', color: { argb: EXPENSE_HEADER_GREEN } },
      };
    }
  }

  for (let column = 4; column <= 6; column += 1) {
    const title = sheet.getCell(expenseTitleRow, column);
    setFill(title, EXPENSE_HEADER_GREEN);
    title.font = { ...title.font, name: 'Aptos', bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
    title.alignment = { ...title.alignment, horizontal: 'center', vertical: 'middle', wrapText: true };
    title.border = thinBorder;
    const header = sheet.getCell(expenseHeaderRow, column);
    setFill(header, EXPENSE_HEADER_GREEN);
    header.font = { ...header.font, name: 'Aptos', bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    header.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    header.border = thinBorder;
  }
  sheet.getRow(expenseTitleRow).height = Math.max(Number(sheet.getRow(expenseTitleRow).height) || 0, 26);
  sheet.getRow(expenseHeaderRow).height = Math.max(Number(sheet.getRow(expenseHeaderRow).height) || 0, 26);

  for (let rowNumber = expenseStartRow; rowNumber <= expenseDataEndRow; rowNumber += 1) {
    for (let column = 4; column <= 6; column += 1) {
      const cell = sheet.getCell(rowNumber, column);
      setFill(cell, rowNumber % 2 ? 'FFFFFFFF' : 'FFF4F8F4');
      cell.font = {
        ...cell.font,
        name: 'Aptos',
        size: 10,
        ...(column === 4 ? { bold: true, color: { argb: 'FF264B47' } } : {}),
      };
      cell.alignment = {
        ...cell.alignment,
        horizontal: column === 6 ? 'right' : column === 4 ? 'center' : 'left',
        vertical: 'middle',
        wrapText: true,
      };
      cell.border = thinBorder;
    }
  }
  for (let column = 4; column <= 6; column += 1) {
    const total = sheet.getCell(expenseTotalRow, column);
    setFill(total, GRAND_TOTAL_GREEN);
    total.font = { ...total.font, name: 'Aptos', bold: true, color: { argb: 'FF274E13' } };
    total.alignment = { ...total.alignment, horizontal: column === 6 ? 'right' : 'left', vertical: 'middle' };
    total.border = {
      ...thinBorder,
      top: { style: 'medium', color: { argb: EXPENSE_HEADER_GREEN } },
    };
  }
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
  const expenseCategoryNames = allExpenses.map((expense) =>
    String(expense.category || expense.categoryName || '').trim(),
  ).filter(Boolean);
  const expenseTitles = allExpenses.map((expense) => String(expense.title || ''));
  const expenseAmounts = allExpenses.map((expense) =>
    `${Number(expense.amount || 0).toLocaleString('fr-FR')} FCFA`,
  );
  growColumnToFit(sheet, 4, [
    'CATÉGORIE',
    'TYPE DE LOGEMENT',
    ...data.tenants.map((tenant) => tenant.housingType),
    ...expenseCategoryNames,
  ], 32);
  growColumnToFit(sheet, 5, ['DÉSIGNATION', ...expenseTitles], 52);
  growColumnToFit(sheet, 6, ['MONTANT', ...expenseAmounts], 24);
  growColumnToFit(sheet, 7, ['STATUT DU PAIEMENT', 'QUANTITÉ', ...data.tenants.map((tenant) => tenant.status)], 22);
  growColumnToFit(sheet, 8, ['UNITÉ', ...allExpenses.map((expense) => expense.unit)], 18);
  growColumnToFit(sheet, 9, ['PRIX UNITAIRE', ...allExpenses.map((expense) => expense.unitPrice)], 22);
  growColumnToFit(sheet, 10, ['BÉNÉFICIAIRE', ...allExpenses.map((expense) => expense.beneficiary)], 24);
  growColumnToFit(sheet, 11, ['PERSONNE AYANT PAYÉ', ...allExpenses.map((expense) => expense.paidBy)], 24);
  growColumnToFit(sheet, 12, ['MODE DE PAIEMENT', ...allExpenses.map((expense) => expense.paymentMethod)], 22);
  growColumnToFit(sheet, 13, ['RÉFÉRENCE', ...allExpenses.map((expense) => expense.reference)], 24);
  growColumnToFit(sheet, 14, ['COMMENTAIRE', ...allExpenses.map((expense) => expense.comment)], 44);
  sheet.getColumn(4).width = Math.min(
    Math.max(Number(sheet.getColumn(4).width) || 30, 24),
    36,
  );
  sheet.getColumn(5).width = Math.min(
    Math.max(Number(sheet.getColumn(5).width) || 18, 42),
    50,
  );
  sheet.getColumn(6).width = Math.min(
    Math.max(Number(sheet.getColumn(6).width) || 18, 30),
    44,
  );

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
  const displayedExpenses = [...data.expenses];
  const protectedExpenses = data.protectedExpenses || [];
  const allTableExpenses = [...displayedExpenses, ...protectedExpenses];
  const expenseCategoryName = (expense) => {
    const rawCategory = String(expense.category || expense.categoryName || '').trim() || 'Autres dépenses';
    if (Array.isArray(data.categories)) return rawCategory;
    return resolvePriorityCategory(rawCategory) || rawCategory;
  };
  const categoryPriorityIndex = (name) => {
    const normalizedName = normalize(name);
    const resolved = resolvePriorityCategory(name);
    const orderedNames = [
      'Alimentaire',
      'Compte boutique',
      'Scolarité et frais éducatifs',
      'Santé',
      'Hygiène et beauté',
      'Eau',
      'Électricité',
      'Wifi',
      'Aumône et dons',
      'Entretien de la maison',
      'Transport',
    ];
    const presentationName = orderedNames.find(
      (category) => normalize(category) === normalize(resolved || name),
    ) || resolved;
    const index = orderedNames.findIndex(
      (category) => normalize(category) === normalize(presentationName),
    );
    return index >= 0 ? index : orderedNames.length;
  };
  const sortedTableExpenses = [...allTableExpenses].sort((left, right) => {
    const leftCategory = expenseCategoryName(left);
    const rightCategory = expenseCategoryName(right);
    const leftIndex = categoryPriorityIndex(leftCategory);
    const rightIndex = categoryPriorityIndex(rightCategory);
    if (leftIndex !== rightIndex) return leftIndex - rightIndex;
    if (leftIndex === PRIORITY_CATEGORIES.length) {
      const categoryComparison = leftCategory.localeCompare(rightCategory, 'fr', { sensitivity: 'base' });
      if (categoryComparison !== 0) return categoryComparison;
    }
    return String(left.title || '').localeCompare(String(right.title || ''), 'fr', { sensitivity: 'base' });
  });
  const expenseRowCount = sortedTableExpenses.length;
  const protectedExpenseRowCount = protectedExpenses.length;

  unmergeEditableRegions(sheet);
  if (tenantInsertions) {
    sheet.spliceRows(11, 0, ...Array.from({ length: tenantInsertions }, () => []));
  }

  const firstIncomeTitleRow = 20 + tenantInsertions;
  if (incomeInsertions) {
    sheet.spliceRows(firstIncomeTitleRow, 0, ...Array.from({ length: incomeInsertions }, () => []));
  }
  const rowShift = tenantInsertions + incomeInsertions;
  const expenseHeaderRow = 21 + rowShift;
  const expenseStartRow = 22 + rowShift;
  const initialTotalRow = templateTotalRow + rowShift;
  const expenseDataEndRow = expenseStartRow + expenseRowCount - 1;
  const expenseTotalRow = expenseDataEndRow + 1;
  const requiredEndRow = expenseTotalRow;
  const insertCount = Math.max(0, requiredEndRow - initialTotalRow);
  if (insertCount) {
    sheet.spliceRows(initialTotalRow + 1, 0, ...Array.from({ length: insertCount }, () => []));
  }
  const protectedSectionRow = null;
  const protectedTotalRow = null;

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
    const rowFill = index % 2 === 0 ? 'FFF7F9F5' : 'FFFFFFFF';
    for (let columnNumber = 1; columnNumber <= 6; columnNumber += 1) {
      const cell = sheet.getCell(rowNumber, columnNumber);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: rowFill },
        bgColor: { argb: rowFill },
      };
    }
  });
  for (let rowNumber = 2; rowNumber <= 13 + tenantInsertions; rowNumber += 1) {
    for (let columnNumber = 1; columnNumber <= 7; columnNumber += 1) {
      sheet.getCell(rowNumber, columnNumber).border = {};
    }
  }

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
  for (let columnNumber = 2; columnNumber <= 6; columnNumber += 1) {
    const cell = sheet.getCell(tenantTotalRow, columnNumber);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: GRAND_TOTAL_GREEN },
      bgColor: { argb: GRAND_TOTAL_GREEN },
    };
    cell.font = { ...cell.font, bold: true, color: { argb: 'FF274E13' } };
  }

  const paymentTotalRow = 13 + tenantInsertions;
  sheet.getCell(paymentTotalRow, 5).value = 'Total des paiements reçus';
  const cashFormula = data.rentPayments.length
    ? `SUM(${data.rentPayments.map((payment) => Number(payment.amount)).join(',')})`
    : '0';
  setFormula(sheet.getCell(paymentTotalRow, 6), cashFormula, data.rentReceived);
  sheet.getCell(paymentTotalRow, 6).numFmt = CURRENCY_FORMAT;
  for (let columnNumber = 5; columnNumber <= 6; columnNumber += 1) {
    const cell = sheet.getCell(paymentTotalRow, columnNumber);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: GRAND_TOTAL_GREEN },
      bgColor: { argb: GRAND_TOTAL_GREEN },
    };
    cell.font = { ...cell.font, bold: true, color: { argb: 'FF274E13' } };
  }

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
  const savingsAboveRow = savingsRow - 1;
  for (const columnNumber of [1, 2]) {
    const extraCell = sheet.getCell(savingsAboveRow, columnNumber);
    extraCell.value = null;
    extraCell.fill = { type: 'pattern', pattern: 'none' };
    extraCell.border = {};
  }

  for (let rowNumber = 20; rowNumber <= requiredEndRow; rowNumber += 1) {
    for (let columnNumber = 4; columnNumber <= 6; columnNumber += 1) {
      const cell = sheet.getCell(rowNumber, columnNumber);
      if (rowNumber < expenseHeaderRow || rowNumber > requiredEndRow) continue;
      cell.value = null;
      cell.fill = { type: 'pattern', pattern: 'none' };
      cell.border = {};
    }
  }
  for (let rowNumber = expenseTotalRow + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    for (let columnNumber = 4; columnNumber <= 6; columnNumber += 1) {
      const cell = sheet.getCell(rowNumber, columnNumber);
      cell.value = null;
      cell.fill = { type: 'pattern', pattern: 'none' };
      cell.border = {};
    }
  }

  styleExpenseSectionHeadingRow(
    sheet,
    21 + rowShift - 1,
    `DÉPENSES DU MOIS ${withDe(data.monthLabel).toUpperCase()}`,
    6,
    4,
  );
  for (let columnNumber = 4; columnNumber <= 6; columnNumber += 1) {
    const cell = sheet.getCell(21 + rowShift - 1, columnNumber);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: EXPENSE_HEADER_GREEN },
      bgColor: { argb: EXPENSE_HEADER_GREEN },
    };
    cell.font = {
      ...cell.font,
      name: 'Aptos',
      bold: true,
      size: 12,
      color: { argb: 'FFFFFFFF' },
    };
  }

  ['Catégorie', 'Désignation', 'Montant'].forEach((header, index) => {
    const cell = sheet.getCell(expenseHeaderRow, 4 + index);
    cell.value = header;
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: EXPENSE_HEADER_GREEN },
      bgColor: { argb: EXPENSE_HEADER_GREEN },
    };
    cell.font = { ...cell.font, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFB8C6D9' } } };
  });
  sheet.getRow(expenseHeaderRow).height = 28;

  sortedTableExpenses.forEach((expense, index) => {
    const rowNumber = expenseStartRow + index;
    const category = expenseCategoryName(expense);
    sheet.getCell(rowNumber, 4).value = category;
    sheet.getCell(rowNumber, 5).value = expense.title || '';
    sheet.getCell(rowNumber, 6).value = Number(expense.amount || 0);
    sheet.getCell(rowNumber, 6).numFmt = CURRENCY_FORMAT;
    sheet.getCell(rowNumber, 5).font = {
      ...sheet.getCell(rowNumber, 5).font,
      name: 'Aptos',
      size: 10,
      bold: false,
      color: { argb: BODY_TEXT_GRAY },
    };
    sheet.getCell(rowNumber, 6).font = {
      ...sheet.getCell(rowNumber, 6).font,
      name: 'Aptos',
      size: 10,
      bold: false,
      color: { argb: RENT_AMOUNT_ORANGE },
    };
    sheet.getCell(rowNumber, 7).value = category;
    for (let columnNumber = 4; columnNumber <= 6; columnNumber += 1) {
      const cell = sheet.getCell(rowNumber, columnNumber);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: index % 2 ? 'FFFFFFFF' : 'FFF7F9F5' },
        bgColor: { argb: index % 2 ? 'FFFFFFFF' : 'FFF7F9F5' },
      };
      cell.border = { bottom: { style: 'thin', color: { argb: GRID_GRAY } } };
      cell.alignment = {
        vertical: 'middle',
        horizontal: columnNumber === 6 ? 'right' : 'left',
        wrapText: true,
      };
    }
    sheet.getCell(rowNumber, 4).font = {
      ...sheet.getCell(rowNumber, 4).font,
      name: 'Aptos',
      bold: true,
      size: 10,
      color: { argb: 'FF264B47' },
    };
    sheet.getCell(rowNumber, 4).alignment = {
      ...sheet.getCell(rowNumber, 4).alignment,
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    };
    sheet.getRow(rowNumber).height = 21;
  });

  let categoryStartIndex = 0;
  while (categoryStartIndex < sortedTableExpenses.length) {
    const category = expenseCategoryName(sortedTableExpenses[categoryStartIndex]);
    let categoryEndIndex = categoryStartIndex + 1;
    while (
      categoryEndIndex < sortedTableExpenses.length &&
      expenseCategoryName(sortedTableExpenses[categoryEndIndex]) === category
    ) {
      categoryEndIndex += 1;
    }
    if (categoryEndIndex - categoryStartIndex >= 2) {
      const firstRow = expenseStartRow + categoryStartIndex;
      const lastRow = expenseStartRow + categoryEndIndex - 1;
      sheet.mergeCells(firstRow, 4, lastRow, 4);
      const categoryCell = sheet.getCell(firstRow, 4);
      categoryCell.value = category;
      categoryCell.font = {
        ...categoryCell.font,
        name: 'Aptos',
        bold: true,
        size: 10,
        color: { argb: 'FF264B47' },
      };
      categoryCell.alignment = {
        ...categoryCell.alignment,
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true,
      };
      categoryCell.border = {
        top: { style: 'thin', color: { argb: GRID_GRAY } },
        bottom: { style: 'thin', color: { argb: GRID_GRAY } },
        left: { style: 'thin', color: { argb: GRID_GRAY } },
        right: { style: 'thin', color: { argb: GRID_GRAY } },
      };
    } else {
      const categoryCell = sheet.getCell(expenseStartRow + categoryStartIndex, 4);
      categoryCell.font = {
        ...categoryCell.font,
        name: 'Aptos',
        bold: true,
        size: 10,
        color: { argb: 'FF264B47' },
      };
    }
    categoryStartIndex = categoryEndIndex;
  }
  sheet.getColumn(7).hidden = true;

  sheet.getCell(expenseTotalRow, 4).value = 'Total général';
  if (sortedTableExpenses.length) {
    setFormula(
      sheet.getCell(expenseTotalRow, 6),
      `SUM(F${expenseStartRow}:F${expenseDataEndRow})`,
      allTableExpenses.reduce((sum, expense) => sum + Number(expense.amount || 0), 0),
    );
  } else {
    sheet.getCell(expenseTotalRow, 6).value = 0;
  }
  sheet.getCell(expenseTotalRow, 6).numFmt = CURRENCY_FORMAT;
  styleExpenseGrandTotalRow(sheet, expenseTotalRow, 6, 4);
  sheet.getCell(expenseTotalRow, 6).font = {
    ...sheet.getCell(expenseTotalRow, 6).font,
    bold: true,
    color: { argb: RENT_AMOUNT_ORANGE },
  };

  if (sortedTableExpenses.length) {
    sheet.autoFilter = `D${expenseHeaderRow}:F${expenseDataEndRow}`;
  } else {
    sheet.autoFilter = `D${expenseHeaderRow}:F${expenseHeaderRow}`;
  }

  const balanceFormulas = buildMonthlyBalanceFormulas({
    incomeCell: `B${totalIncomeRow}`,
    openingCell: `B${openingRow}`,
    resourcesCell: `B${resourcesRow}`,
    expensesCell: data.expenses.length
      ? `SUM(F${expenseStartRow}:F${expenseStartRow + data.expenses.length - 1})`
      : '0',
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

  const financialHeadingRow = 14 + tenantInsertions;
  const otherIncomeRows = extraIncomeSources.map((_, index) => 18 + tenantInsertions + index);
  const financialSummaryRows = [
    pensionRow,
    papeRow,
    ...otherIncomeRows,
    totalIncomeRow,
    remainingRow,
    openingRow,
    resourcesRow,
    savingsRow,
  ];
  styleFinancialSummaryHeading(sheet, financialHeadingRow);
  styleFinancialSummaryRows(
    sheet,
    financialSummaryRows,
    [totalIncomeRow, remainingRow, resourcesRow, savingsRow],
  );

  sheet.getCell(1, 1).value = `Loyers du mois ${withDe(data.previousMonthLabel)}`;
  const expenseHeadingRow = 20 + rowShift;
  const expenseTitle = sheet.getCell(expenseHeadingRow, 4);
  expenseTitle.value = `DÉPENSES DU MOIS ${withDe(data.monthLabel).toUpperCase()}`;
  for (let columnNumber = 4; columnNumber <= 6; columnNumber += 1) {
    const cell = sheet.getCell(expenseHeadingRow, columnNumber);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: EXPENSE_HEADER_GREEN },
      bgColor: { argb: EXPENSE_HEADER_GREEN },
    };
    cell.font = {
      ...cell.font,
      name: 'Aptos',
      bold: true,
      size: 12,
      color: { argb: 'FFFFFFFF' },
    };
  }
  fitMonthlyColumns(
    sheet,
    data,
    expenseStartRow,
    expenseTotalRow,
    Math.max(expenseTotalRow, remainingRow, resourcesRow, savingsRow, protectedTotalRow || 0),
  );
  for (let column = 1; column <= 7; column += 1) {
    if (column !== 3) setWrappedContent(sheet, 2, column, column === 2 ? 3 : column);
  }
  setWrappedContent(sheet, expenseHeadingRow, 4, 6);
  sheet.pageSetup.printArea = `A1:F${expenseTotalRow}`;
  sheet.pageSetup.orientation = 'landscape';
  sheet.pageSetup.fitToPage = true;
  sheet.pageSetup.fitToWidth = 1;
  sheet.pageSetup.fitToHeight = 0;
  styleMonthlyTables(sheet, {
    tenantCount: data.tenants.length,
    tenantTotalRow,
    paymentTotalRow,
    expenseTitleRow: expenseHeadingRow,
    expenseHeaderRow,
    expenseStartRow,
    expenseDataEndRow,
    expenseTotalRow,
  });

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

function styleAnnualSummary(sheet, monthEndRow, annualTotalRow) {
  for (let columnNumber = 1; columnNumber <= 7; columnNumber += 1) {
    const titleCell = sheet.getCell(1, columnNumber);
    titleCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: EXPENSE_HEADER_GREEN },
      bgColor: { argb: EXPENSE_HEADER_GREEN },
    };
    titleCell.font = { ...titleCell.font, name: 'Aptos', bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    titleCell.border = { bottom: { style: 'thin', color: { argb: 'FFB8C6D9' } } };

    const headerCell = sheet.getCell(2, columnNumber);
    headerCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: EXPENSE_HEADER_GRAY },
      bgColor: { argb: EXPENSE_HEADER_GRAY },
    };
    headerCell.font = { ...headerCell.font, name: 'Aptos', bold: true, color: { argb: 'FFFFFFFF' } };
    headerCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    headerCell.border = { bottom: { style: 'thin', color: { argb: 'FFB8C6D9' } } };

    const totalCell = sheet.getCell(annualTotalRow, columnNumber);
    totalCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: GRAND_TOTAL_GREEN },
      bgColor: { argb: GRAND_TOTAL_GREEN },
    };
    totalCell.font = { ...totalCell.font, name: 'Aptos', bold: true, color: { argb: 'FF274E13' } };
    totalCell.alignment = {
      ...totalCell.alignment,
      vertical: 'middle',
      horizontal: columnNumber === 1 ? 'left' : 'right',
    };
    totalCell.border = {
      ...totalCell.border,
      top: { style: 'medium', color: { argb: EXPENSE_HEADER_GREEN } },
      bottom: { style: 'thin', color: { argb: EXPENSE_HEADER_GREEN } },
    };
  }
  for (let rowNumber = 3; rowNumber <= monthEndRow; rowNumber += 1) {
    for (let columnNumber = 1; columnNumber <= 7; columnNumber += 1) {
      const cell = sheet.getCell(rowNumber, columnNumber);
      cell.font = { ...cell.font, name: 'Aptos', size: 10, color: { argb: BODY_TEXT_GRAY } };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: rowNumber % 2 ? 'FFFFFFFF' : 'FFF4F8F4' },
        bgColor: { argb: rowNumber % 2 ? 'FFFFFFFF' : 'FFF4F8F4' },
      };
      cell.alignment = {
        ...cell.alignment,
        horizontal: columnNumber === 1 ? 'left' : 'right',
        vertical: 'middle',
        wrapText: true,
      };
      cell.border = { bottom: { style: 'thin', color: { argb: GRID_GRAY } } };
    }
    sheet.getRow(rowNumber).height = Math.max(Number(sheet.getRow(rowNumber).height) || 0, 20);
  }
  sheet.getRow(1).height = 25;
  sheet.getRow(2).height = 27;
  sheet.getRow(annualTotalRow).height = Math.max(Number(sheet.getRow(annualTotalRow).height) || 0, 22);
}

function buildAnnualSummary(workbook, templateSheet, year, months) {
  const sheet = cloneWorksheet(workbook, templateSheet, 'Récapitulatif annuel');
  for (const range of [...sheet.model.merges]) {
    if (range !== 'A1:G1') sheet.unMergeCells(range);
  }
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    for (let column = 1; column <= 7; column += 1) sheet.getCell(rowNumber, column).value = null;
  }
  for (let rowNumber = 16; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    for (let column = 1; column <= 7; column += 1) sheet.getCell(rowNumber, column).style = {};
    row.style = {};
    row.height = undefined;
  }

  sheet.getCell('A1').value = `Récapitulatif annuel ${year}`;
  const headers = ['Mois', 'Recettes encaissées', '', 'Loyers attendus', 'Dépenses', 'Solde final', 'Épargne'];
  headers.forEach((header, index) => { sheet.getCell(2, index + 1).value = header; });
  sheet.mergeCells('B2:C2');
  const monthEndRow = 2 + months.length;
  const annualTotalRow = monthEndRow + 1;
  for (let rowNumber = 3; rowNumber <= monthEndRow; rowNumber += 1) {
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

  copyRowAppearance(sheet, findExpenseTotalRow(templateSheet), annualTotalRow);
  const annualValues = [
    months.reduce((sum, month) => sum + month.incomeTotal, 0),
    months.reduce((sum, month) => sum + month.rentExpected, 0),
    months.reduce((sum, month) => sum + month.totalExpenses, 0),
    months.at(-1)?.closingBalance || 0,
    months.at(-1)?.savingsBalance || 0,
  ];
  sheet.getCell(annualTotalRow, 1).value = 'Total annuel';
  setFormula(sheet.getCell(annualTotalRow, 2), `SUM(B3:B${monthEndRow})`, annualValues[0]);
  setFormula(sheet.getCell(annualTotalRow, 4), `SUM(D3:D${monthEndRow})`, annualValues[1]);
  setFormula(sheet.getCell(annualTotalRow, 5), `SUM(E3:E${monthEndRow})`, annualValues[2]);
  setFormula(sheet.getCell(annualTotalRow, 6), `F${monthEndRow}`, annualValues[3]);
  setFormula(sheet.getCell(annualTotalRow, 7), `G${monthEndRow}`, annualValues[4]);
  [2, 4, 5, 6, 7].forEach((column) => { sheet.getCell(annualTotalRow, column).numFmt = CURRENCY_FORMAT; });
  sheet.mergeCells(annualTotalRow, 2, annualTotalRow, 3);
  styleAnnualSummary(sheet, monthEndRow, annualTotalRow);
  sheet.pageSetup.printArea = `A1:G${annualTotalRow}`;
  sheet.pageSetup.orientation = 'landscape';
  sheet.pageSetup.fitToPage = true;
  sheet.pageSetup.fitToWidth = 1;
  sheet.pageSetup.fitToHeight = 1;
}

async function buildReferenceWorkbook(
  months,
  { annual = false, includeCarryover = false, includeAnnualSummary } = {},
) {
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

  const shouldIncludeAnnualSummary = includeAnnualSummary === undefined
    ? annual
    : Boolean(includeAnnualSummary);
  if (shouldIncludeAnnualSummary) {
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
