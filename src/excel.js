const XLSX = require('xlsx');

const START_DATE = '2026-01-01';
const END_DATE = '2026-08-31';
const MONTHS = [
  'JANVIER', 'FÉVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN',
  'JUILLET', 'AOÛT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DÉCEMBRE',
];

const normalize = (value) => String(value ?? '')
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g, '')
  .trim()
  .toLowerCase()
  .replace(/[^a-z0-9]+/g, ' ')
  .trim();

const HEADER_ALIASES = new Map();
const addAliases = (field, values) => values.forEach((value) => HEADER_ALIASES.set(normalize(value), field));

addAliases('title', ['intitule', 'libelle', 'designation', 'nom', 'depense', 'article', 'marchandise', 'produit', 'item name', 'title', 'name']);
addAliases('date', ['date', 'date operation', 'date depense', 'date achat', 'purchase date', 'expense date']);
addAliases('day', ['jour', 'day']);
addAliases('amount', ['montant', 'amount', 'total', 'total ttc', 'total paye', 'montant total', 'prix total', 'valeur', 'total boutique', 'total compte boutique', 'montant compte boutique']);
addAliases('category', ['categorie', 'category', 'rubrique', 'poste', 'nature', 'type de depense']);
addAliases('type', ['type', 'operation', 'nature operation']);
addAliases('quantity', ['quantite', 'qte', 'qty', 'quantity']);
addAliases('unit', ['unite', 'unit', 'mesure']);
addAliases('unitPrice', ['prix unitaire', 'pu', 'unit price', 'prix par unite', 'prix unite']);
addAliases('notes', ['note', 'notes', 'observation', 'commentaire', 'description']);
addAliases('reference', ['reference', 'numero', 'n facture', 'facture']);

const MONTH_ALIASES = new Map([
  ['janvier', 1], ['jan', 1], ['fevrier', 2], ['fev', 2], ['mars', 3],
  ['avril', 4], ['avr', 4], ['mai', 5], ['juin', 6], ['juillet', 7],
  ['juil', 7], ['aout', 8], ['septembre', 9], ['sept', 9],
  ['octobre', 10], ['oct', 10], ['novembre', 11], ['nov', 11],
  ['decembre', 12], ['dec', 12],
]);

const CATEGORY_ALIASES = new Map([
  ['electricite', 'electricite'], ['eau', 'eau'],
  ['wifi', 'internet'], ['wi fi', 'internet'], ['internet', 'internet'],
  ['sante', 'sante'], ['ordonnance', 'sante'], ['ordonnances', 'sante'],
  ['transport', 'transport'], ['entretien', 'entretien'],
  ['reparations', 'reparations'], ['reparation', 'reparations'],
  ['scolarite', 'scolarite'], ['autre', 'autres'], ['autres', 'autres'],
]);

function parseAmount(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  const raw = String(value).trim().replace(/[\s\u00a0\u202f']/g, '').replace(/[^\d,.-]/g, '');
  if (!raw) return null;
  const comma = raw.lastIndexOf(',');
  const dot = raw.lastIndexOf('.');
  let normalized = raw;
  if (comma >= 0 && dot >= 0) {
    const decimal = Math.max(comma, dot);
    const integer = raw.slice(0, decimal).replace(/[,.]/g, '');
    normalized = `${integer}.${raw.slice(decimal + 1).replace(/[,.]/g, '')}`;
  } else if (comma >= 0 || dot >= 0) {
    const separator = comma >= 0 ? ',' : '.';
    const pieces = raw.split(separator);
    const trailingDigits = pieces.at(-1).length;
    normalized = trailingDigits > 0 && trailingDigits <= 2
      ? `${pieces.slice(0, -1).join('')}.${pieces.at(-1)}`
      : pieces.join('');
  }
  const amount = Number(normalized);
  return Number.isFinite(amount) ? amount : null;
}

function dateFromParts(year, month, day) {
  const y = Number(year);
  const m = Number(month);
  const d = Number(day);
  if (!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d)) return null;
  const date = new Date(Date.UTC(y, m - 1, d));
  if (date.getUTCFullYear() !== y || date.getUTCMonth() !== m - 1 || date.getUTCDate() !== d) return null;
  return `${String(y).padStart(4, '0')}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

function parseDate(value, options = {}) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return dateFromParts(value.getUTCFullYear(), value.getUTCMonth() + 1, value.getUTCDate());
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    if (options.month && Number.isInteger(value) && value >= 1 && value <= 31) {
      return dateFromParts(2026, options.month, value);
    }
    const parsed = XLSX.SSF.parse_date_code(value);
    return parsed ? dateFromParts(parsed.y, parsed.m, parsed.d) : null;
  }
  if (value === null || value === undefined || value === '') return null;
  const text = String(value).trim();
  const iso = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (iso) return dateFromParts(iso[1], iso[2], iso[3]);
  const french = text.match(/^(\d{1,2})[/. -](\d{1,2})[/. -](\d{2,4})$/);
  if (french) {
    const year = french[3].length === 2 ? `20${french[3]}` : french[3];
    return dateFromParts(year, french[2], french[1]);
  }
  const namedMonth = normalize(text).match(/^(\d{1,2}) ([a-z]+) (\d{4})$/);
  if (namedMonth && MONTH_ALIASES.has(namedMonth[2])) {
    return dateFromParts(namedMonth[3], MONTH_ALIASES.get(namedMonth[2]), namedMonth[1]);
  }
  const dayOnly = text.match(/^(\d{1,2})$/);
  if (dayOnly && options.month) return dateFromParts(2026, options.month, dayOnly[1]);
  return null;
}

function mapHeaders(row) {
  const columns = {};
  row.forEach((value, index) => {
    const field = HEADER_ALIASES.get(normalize(value));
    if (field && columns[field] === undefined) columns[field] = index;
  });
  const fields = Object.keys(columns);
  return fields.includes('title') && fields.some((field) => ['amount', 'date', 'quantity', 'unitPrice'].includes(field))
    ? columns
    : null;
}

function monthForSheet(sheetName) {
  const normalized = normalize(sheetName);
  for (const [name, month] of MONTH_ALIASES) {
    if (normalized === name || normalized.includes(name)) return month;
  }
  return null;
}

function categoryResolution(value, categories) {
  const normalized = normalize(value);
  const alias = CATEGORY_ALIASES.get(normalized)
    || (/wifi|wi fi|internet/.test(normalized) ? 'internet' : null)
    || (/sante|ordonnance/.test(normalized) ? 'sante' : null)
    || (/electricite/.test(normalized) ? 'electricite' : null)
    || (/transport/.test(normalized) ? 'transport' : null)
    || (/entretien|maintenance/.test(normalized) ? 'entretien' : null);
  const known = categories.find((category) => normalize(category.name) === normalized)
    || (alias && categories.find((category) => normalize(category.name) === alias))
    || (alias === 'internet' && categories.find((category) => normalize(category.name).includes('internet')))
    || (alias === 'sante' && categories.find((category) => normalize(category.name).includes('sante')));
  const fallback = categories.find((category) => ['autre', 'autres'].includes(normalize(category.name)));
  return {
    name: known?.name || fallback?.name || 'Autres',
    recognized: Boolean(known),
  };
}

function parseHistoricalWorkbook(buffer, categories = []) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const records = [];
  const excludedRows = [];
  const ignoredRows = [];
  const summaryLines = [];
  let lineNumber = 0;

  workbook.SheetNames.forEach((sheetName) => {
    if (normalize(sheetName).includes('recapitulatif')) {
      ignoredRows.push({ sheetName, reason: 'Feuille de synthèse, non importée comme opérations' });
      return;
    }

    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, blankrows: false });
    const sheetMonth = monthForSheet(sheetName);
    let columns = null;
    let section = '';

    rows.forEach((cells, rowIndex) => {
      if (!cells.some((cell) => cell !== null && String(cell).trim() !== '')) return;
      const rowText = normalize(cells.filter((cell) => cell !== null).join(' '));
      const detectedColumns = mapHeaders(cells);
      if (detectedColumns) {
        columns = detectedColumns;
        return;
      }

      const hasNumericCell = cells.some((cell) => parseAmount(cell) !== null);
      if (!hasNumericCell && /compte boutique|ravitaillement|marchandise|depenses ordinaires|depenses diverses/.test(rowText)) {
        section = /boutique|ravitaillement|marchandise/.test(rowText) ? 'grocery' : 'ordinary';
        return;
      }
      if (!columns) {
        ignoredRows.push({
          sheetName,
          row: rowIndex + 1,
          values: cells.map((cell) => cell === null || cell === undefined ? '' : String(cell)),
          reason: 'En-tête de colonnes non reconnu',
        });
        return;
      }

      const get = (field) => columns[field] === undefined ? null : cells[columns[field]];
      const title = String(get('title') ?? '').trim();
      const categoryRaw = String(get('category') ?? '').trim();
      const typeRaw = String(get('type') ?? '').trim();
      const categoryValue = categoryRaw || typeRaw;
      const titleNorm = normalize(title);
      const rowContext = normalize(`${sheetName} ${section} ${categoryRaw} ${typeRaw}`);
      const quantity = parseAmount(get('quantity'));
      const unitPrice = parseAmount(get('unitPrice'));
      const unit = String(get('unit') ?? '').trim();
      const dateValue = get('date') ?? get('day');
      const date = parseDate(dateValue, { month: sheetMonth });
      let amount = parseAmount(get('amount'));
      const itemDetails = quantity !== null || unit !== '' || unitPrice !== null;
      const isBoutique = /boutique|ravitaillement|marchandise|epicerie/.test(rowContext) || itemDetails;
      const isBoutiqueSummary = /^(compte boutique|total boutique|total compte boutique|sous total boutique)$/.test(titleNorm)
        && !itemDetails;
      const isTotalRow = /^(total|sous total|total general|total mensuel)$/.test(titleNorm);

      if (isTotalRow) return;
      if (isBoutiqueSummary) {
        summaryLines.push({ sheetName, row: rowIndex + 1, date, amount });
        return;
      }

      lineNumber += 1;
      const issues = [];
      if (!date) issues.push('Date absente ou illisible');
      if (!title) issues.push('Intitulé absent');
      if (amount === null && quantity !== null && unitPrice !== null) amount = quantity * unitPrice;
      if (amount === null) issues.push('Montant absent ou illisible');
      if (amount !== null && quantity !== null && unitPrice !== null
        && Math.abs(amount - quantity * unitPrice) > 0.01) {
        issues.push('Le montant ne correspond pas à quantité × prix unitaire');
      }

      const resolution = isBoutique
        ? { name: 'Boutique', recognized: true }
        : categoryResolution(categoryValue, categories);
      if (!isBoutique && !resolution.recognized) issues.push('Catégorie inconnue, classée dans Autres');
      const record = {
        lineNumber,
        sheetName,
        rowNumber: rowIndex + 1,
        title,
        date,
        amount,
        type: isBoutique ? 'grocery' : 'expense',
        categoryRaw: categoryValue,
        categoryName: resolution.name,
        categoryRecognized: resolution.recognized,
        quantity,
        unit: unit || null,
        unitPrice,
        quantityKnown: quantity !== null,
        unitPriceKnown: unitPrice !== null,
        notes: String(get('notes') ?? '').trim(),
        reference: String(get('reference') ?? '').trim(),
        issues,
        valid: Boolean(date && title && amount !== null),
        duplicateSource: null,
      };

      if (date && (date < START_DATE || date > END_DATE)) {
        excludedRows.push({ ...record, reason: 'Hors de la période du 01/01/2026 au 31/08/2026' });
      } else {
        records.push(record);
      }
    });
  });

  const summaryReports = summaryLines.map((summary) => {
    const details = records.filter((record) => record.valid
      && record.type === 'grocery'
      && record.date?.slice(0, 7) === summary.date?.slice(0, 7));
    const detailTotal = details.reduce((sum, record) => sum + record.amount, 0);
    const unresolved = summary.amount !== null && details.length === 0;
    const mismatch = summary.amount !== null && details.length > 0
      && Math.abs(summary.amount - detailTotal) > 0.01;
    return {
      ...summary,
      detailCount: details.length,
      detailTotal,
      unresolved,
      mismatch,
    };
  });
  return {
    records,
    excludedRows,
    ignoredRows,
    summaryLines: summaryReports,
    sheetNames: workbook.SheetNames,
  };
}

function signature(record) {
  if (record.valid === false || !record.date || record.amount === null || record.amount === undefined) return null;
  return [
    record.type,
    normalize(record.title),
    record.date,
    Math.round(Number(record.amount) * 100),
  ].join('|');
}

function markDuplicates(records, existingRecords = []) {
  const seen = new Map();
  existingRecords.forEach((record) => {
    const key = signature(record);
    if (key) seen.set(key, 'base');
  });
  return records.map((record) => {
    const key = signature(record);
    const duplicateSource = key ? (seen.get(key) || null) : null;
    if (key && !duplicateSource) seen.set(key, 'fichier');
    return { ...record, duplicateSource };
  });
}

function monthKey(year, month) {
  return `${year}-${String(month).padStart(2, '0')}`;
}

function summarizeRecords(records) {
  const months = Array.from({ length: 8 }, (_, index) => {
    const month = index + 1;
    return {
      month: monthKey(2026, month),
      label: MONTHS[index],
      ordinaryCount: 0,
      groceryCount: 0,
      totalBoutique: 0,
      totalOther: 0,
      totalGeneral: 0,
    };
  });
  const byMonth = new Map(months.map((month) => [month.month, month]));
  records.filter((record) => record.valid && !record.duplicateSource).forEach((record) => {
    const month = byMonth.get(record.date.slice(0, 7));
    if (!month) return;
    if (record.type === 'grocery') {
      month.groceryCount += 1;
      month.totalBoutique += record.amount;
    } else {
      month.ordinaryCount += 1;
      month.totalOther += record.amount;
    }
    month.totalGeneral = month.totalBoutique + month.totalOther;
  });
  return months;
}

function setNumberFormat(sheet, rows, columns, format) {
  rows.forEach((row) => {
    columns.forEach((column) => {
      const address = XLSX.utils.encode_cell({ r: row, c: column });
      if (sheet[address]) sheet[address].z = format;
    });
  });
}

function setCell(sheet, row, column, value, numberFormat) {
  const address = XLSX.utils.encode_cell({ r: row, c: column });
  const cell = sheet[address];
  if (cell && numberFormat) cell.z = numberFormat;
}

function excelDateSerial(date) {
  const [year, month, day] = date.split('-').map(Number);
  return Date.UTC(year, month - 1, day) / 86_400_000 + 25_569;
}

function buildAnnualWorkbook(year, expenses, groceryItems) {
  const workbook = XLSX.utils.book_new();
  const annual = [];

  for (let month = 1; month <= 12; month += 1) {
    const key = monthKey(year, month);
    const firstDate = `${key}-01`;
    const nextMonthDate = month === 12 ? `${year + 1}-01-01` : `${monthKey(year, month + 1)}-01`;
    const monthExpenses = expenses.filter((row) => row.expense_date >= firstDate && row.expense_date < nextMonthDate);
    const monthGroceries = groceryItems.filter((row) => row.purchase_date >= firstDate && row.purchase_date < nextMonthDate);
    const totalBoutique = monthGroceries.reduce((sum, row) => sum + Number(row.total_amount || 0), 0);
    const totalOther = monthExpenses.reduce((sum, row) => sum + Number(row.amount || 0), 0);
    const label = MONTHS[month - 1];
    annual.push({ label, totalBoutique, totalOther, totalGeneral: totalBoutique + totalOther });

    const consolidatedDate = monthGroceries.length
      ? monthGroceries.map((row) => row.purchase_date).sort()[0]
      : null;
    const mainRows = [
      ['Date', 'Intitulé de la dépense', 'Catégorie', 'Montant (FCFA)', 'Type de dépense'],
      ...monthExpenses.map((row) => [
        excelDateSerial(row.expense_date),
        row.title,
        row.category || 'Autres',
        Number(row.amount),
        'Ordinaire',
      ]),
      ...(monthGroceries.length ? [[
        excelDateSerial(consolidatedDate),
        'COMPTE BOUTIQUE',
        'Boutique',
        totalBoutique,
        'Boutique',
      ]] : []),
    ];
    const mainTotalRow = mainRows.length;
    mainRows.push(['', 'TOTAL DU MOIS', '', {
      f: `SUM(D2:D${mainTotalRow})`,
      v: totalBoutique + totalOther,
    }, '']);
    const detailTitleRow = mainRows.length + 2;
    mainRows.push([], ['DÉTAIL DU COMPTE BOUTIQUE']);
    mainRows.push(['Date', 'Marchandise', 'Quantité', 'Unité', 'Prix unitaire (FCFA)', 'Montant (FCFA)']);
    const detailHeaderRow = mainRows.length;
    monthGroceries.forEach((row) => {
      mainRows.push([
        excelDateSerial(row.purchase_date),
        row.item_name,
        row.quantityKnown === 0 || row.quantityKnown === false ? '' : row.quantity,
        row.unit || '',
        row.unitPriceKnown === 0 || row.unitPriceKnown === false ? '' : row.unit_price,
        Number(row.total_amount),
      ]);
    });
    const detailTotalRow = mainRows.length + 1;
    mainRows.push(['', 'TOTAL COMPTE BOUTIQUE', '', '', '', {
      f: monthGroceries.length ? `SUM(F${detailHeaderRow + 1}:F${detailTotalRow - 1})` : '0',
      v: totalBoutique,
    }]);

    const sheet = XLSX.utils.aoa_to_sheet(mainRows);
    sheet['!autofilter'] = { ref: `A1:E${Math.max(1, mainTotalRow)}` };
    sheet['!cols'] = [{ wch: 15 }, { wch: 34 }, { wch: 25 }, { wch: 18 }, { wch: 21 }, { wch: 18 }];
    sheet['!rows'] = [{ hpt: 24 }];
    sheet['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' };
    setNumberFormat(sheet, Array.from({ length: mainTotalRow - 1 }, (_, index) => index + 1), [3], '#,##0.## "FCFA"');
    setNumberFormat(sheet, Array.from({ length: monthExpenses.length + monthGroceries.length }, (_, index) => index + 1), [0], 'dd/mm/yyyy');
    const detailRows = Array.from({ length: monthGroceries.length }, (_, index) => detailHeaderRow + index);
    setNumberFormat(sheet, detailRows, [4, 5], '#,##0.## "FCFA"');
    setNumberFormat(sheet, detailRows, [0], 'dd/mm/yyyy');
    setCell(sheet, mainTotalRow - 1, 3, totalBoutique + totalOther, '#,##0.## "FCFA"');
    setCell(sheet, detailTotalRow - 1, 5, totalBoutique, '#,##0.## "FCFA"');
    sheet['!merges'] = [
      { s: { r: detailTitleRow - 1, c: 0 }, e: { r: detailTitleRow - 1, c: 5 } },
    ];
    XLSX.utils.book_append_sheet(workbook, sheet, label);
  }

  const summaryRows = [
    ['RÉCAPITULATIF ANNUEL'],
    ['Mois', 'Total boutique (FCFA)', 'Autres dépenses (FCFA)', 'Total général (FCFA)'],
  ];
  annual.forEach((item, index) => {
    const row = index + 3;
    summaryRows.push([
      item.label,
      item.totalBoutique,
      item.totalOther,
      { f: `SUM(B${row}:C${row})`, v: item.totalGeneral },
    ]);
  });
  const annualTotalRow = summaryRows.length + 1;
  summaryRows.push([
    'TOTAL ANNUEL',
    { f: `SUM(B3:B14)`, v: annual.reduce((sum, item) => sum + item.totalBoutique, 0) },
    { f: `SUM(C3:C14)`, v: annual.reduce((sum, item) => sum + item.totalOther, 0) },
    { f: `SUM(D3:D14)`, v: annual.reduce((sum, item) => sum + item.totalGeneral, 0) },
  ]);
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
  summarySheet['!autofilter'] = { ref: 'A2:D14' };
  summarySheet['!cols'] = [{ wch: 24 }, { wch: 24 }, { wch: 27 }, { wch: 25 }];
  summarySheet['!freeze'] = { xSplit: 0, ySplit: 2, topLeftCell: 'A3', activePane: 'bottomLeft', state: 'frozen' };
  setNumberFormat(summarySheet, Array.from({ length: 13 }, (_, index) => index + 2), [1, 2, 3], '#,##0.## "FCFA"');
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'RÉCAPITULATIF ANNUEL');
  workbook.SheetNames.unshift(workbook.SheetNames.pop());
  workbook.Workbook = { Views: [{ activeTab: 0 }] };
  workbook.Props = { Title: `Dépenses maison ${year}`, Subject: `Récapitulatif annuel ${year}`, Author: 'Maison' };
  return { workbook, annual };
}

module.exports = {
  START_DATE,
  END_DATE,
  MONTHS,
  buildAnnualWorkbook,
  categoryResolution,
  markDuplicates,
  normalize,
  parseAmount,
  parseDate,
  parseHistoricalWorkbook,
  summarizeRecords,
};
