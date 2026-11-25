const XLSX = require('xlsx-js-style');
const { parse: parseCsv } = require('csv-parse/sync');

const START_DATE = '2026-01-01';
const END_DATE = '9999-12-31';
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
addAliases('month', ['mois', 'month', 'mois de rattachement', 'month key']);
addAliases('beneficiary', ['beneficiaire', 'beneficiary', 'destinataire']);
addAliases('paidBy', ['paye par', 'personne ayant paye', 'paid by']);
addAliases('paymentMethod', ['mode de paiement', 'moyen de paiement', 'payment method']);
addAliases('title', ['intitule de la depense', 'libelle de la depense', 'designation depense']);
addAliases('category', ['type de depense', 'type de depenses', 'categorie de depense']);

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
      return dateFromParts(options.year || 2026, options.month, value);
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
  if (dayOnly && options.month) return dateFromParts(options.year || 2026, options.month, dayOnly[1]);
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

const TENANT_HEADER_ALIASES = new Map();
const addTenantAliases = (field, values) => values.forEach((value) => TENANT_HEADER_ALIASES.set(normalize(value), field));
addTenantAliases('floor', ['niveau', 'etage', 'niveau etage', 'floor']);
addTenantAliases('name', ['prenom et nom', 'nom du locataire', 'nom locataire', 'locataire', 'tenant name']);
addTenantAliases('housingType', ['categorie de logement', 'type de logement', 'logement']);
addTenantAliases('people', ['nombre de personnes', 'nombre de personne', 'personnes', 'occupants']);
addTenantAliases('rent', ['montant a payer', 'loyer', 'loyer mensuel', 'montant du loyer', 'montant loyer']);
addTenantAliases('status', ['statut de paiement', 'statut paiement', 'paiement', 'status']);
addTenantAliases('paidAmount', ['montant paye', 'montant paye', 'somme recue']);

const isEmptyCell = (value) => value === null || value === undefined || String(value).trim() === '';
const normalizeSheetCell = (value) => (typeof value === 'string' ? value.trim() : value);

function extractWorksheetRows(sheet) {
  if (!sheet['!ref']) return [];
  const range = XLSX.utils.decode_range(sheet['!ref']);
  return Array.from({ length: range.e.r - range.s.r + 1 }, (_, rowOffset) => {
    const rowIndex = range.s.r + rowOffset;
    const values = Array.from({ length: range.e.c - range.s.c + 1 }, (_, columnOffset) => {
      const columnIndex = range.s.c + columnOffset;
      const cell = sheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
      return normalizeSheetCell(cell?.v ?? (cell?.f ? `=${cell.f}` : null));
    });
    const formulas = values.map((_, columnOffset) => {
      const columnIndex = range.s.c + columnOffset;
      const cell = sheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
      return cell?.f ? { column: columnIndex, formula: cell.f, value: cell.v ?? null } : null;
    }).filter(Boolean);
    const populatedIndexes = values
      .map((value, index) => (isEmptyCell(value) ? -1 : index))
      .filter((index) => index >= 0);
    const textLine = populatedIndexes.length === 1 && typeof values[populatedIndexes[0]] === 'string'
      ? values[populatedIndexes[0]]
      : null;
    const delimitedValues = textLine?.includes('|')
      ? textLine.split('|').map((value) => value.trim())
      : null;
    return {
      rowNumber: rowIndex + 1,
      values: delimitedValues || values,
      originalValues: values,
      formulas,
      textDelimited: Boolean(delimitedValues),
    };
  });
}

function mapFields(row, aliases) {
  const columns = {};
  row.forEach((value, index) => {
    const field = aliases.get(normalize(value));
    if (field && columns[field] === undefined) columns[field] = index;
  });
  return columns;
}

function tenantHeader(row) {
  const columns = mapFields(row, TENANT_HEADER_ALIASES);
  return ['floor', 'name', 'housingType', 'people', 'rent', 'status']
    .every((field) => columns[field] !== undefined)
    ? columns
    : null;
}

function expenseHeader(row) {
  const columns = mapFields(row, HEADER_ALIASES);
  const explicitAmountColumn = row.findIndex((value) => [
    'montant', 'amount', 'montant total', 'montant fcfa', 'montant en fcfa',
  ].includes(normalize(value)));
  if (explicitAmountColumn >= 0) columns.amount = explicitAmountColumn;
  return columns.category !== undefined && columns.amount !== undefined
    ? columns
    : null;
}

function canonicalFloor(value, previousFloor) {
  const text = normalize(value);
  if (!text) return previousFloor || null;
  if (/^(rez|rdc|rez de chaussee|rez chaussee|ground)/.test(text)) return 'Rez-de-chaussée';
  if (/\b(1|1er|premier|premiere)\b/.test(text)) return '1er étage';
  if (/\b(2|2eme|deuxieme|second)\b/.test(text)) return '2e étage';
  if (/\b(3|3eme|troisieme)\b/.test(text)) return '3e étage';
  return null;
}

function splitTenantName(value) {
  const name = String(value || '').trim().replace(/\s+/g, ' ');
  const [firstName, ...lastParts] = name.split(' ');
  return { fullName: name, firstName, lastName: lastParts.join(' ') };
}

function paymentStatus(value) {
  const normalized = normalize(value);
  if (/^(paye|payee|regle|reglee|encaisse)$/.test(normalized)) return 'paid';
  if (/partiel|acompte/.test(normalized)) return 'partial';
  if (/impaye|non paye|en attente|du/.test(normalized)) return 'unpaid';
  return null;
}

function isTotalLabel(value) {
  return /^(total|sous total|total general|total mensuel|total des .+)$/.test(normalize(value));
}

function rawRowData(sheetName, row, classification, classifications = [classification], reason = null) {
  return {
    sheetName,
    row: row.rowNumber,
    classification,
    classifications: [...new Set(classifications)],
    values: row.originalValues.map((value) => (
      value instanceof Date ? value.toISOString().slice(0, 10) : value === null ? '' : String(value)
    )),
    formulas: row.formulas,
    textDelimited: row.textDelimited,
    reason,
  };
}

function categoryFromDescription(description, categories) {
  const normalized = normalize(description);
  const aliases = [
    [/compte boutique|boutique/, 'compte boutique'],
    [/sante|ticket|medoc|medicament|consultation|ordonnance|radio oeil/, 'sante'],
    [/electricite/, 'electricite'],
    [/\beau\b|tuyau|fosse|vidange/, 'eau'],
    [/wifi|wi fi|internet/, 'wifi'],
    [/transport|aeroport/, 'transport'],
    [/reparation|reparation|entretien/, 'reparation'],
    [/scolarite|cahier|pedagogique|ecole/, 'scolarite'],
    [/aumone|don/, 'aumone'],
    [/renforcement|aide|argent de poche/, 'renforcement'],
    [/exceptionnel|imprevu|korite|valise|tenue|parfum/, 'exceptionnel'],
    [/menage|femme de menage/, 'menage'],
    [/animal|mouton/, 'animal'],
    [/pain|aliment|repas|riz/, 'alimentaire'],
  ];
  const match = aliases.find(([pattern]) => pattern.test(normalized));
  if (match) {
    const category = categories.find((item) => {
      const name = normalize(item.name);
      switch (match[1]) {
        case 'compte boutique': return name === 'compte boutique';
        case 'sante': return /sante|consultation|medicament/.test(name);
        case 'electricite': return /electricite/.test(name);
        case 'eau': return /eau/.test(name);
        case 'wifi': return /wifi|internet/.test(name);
        case 'transport': return /transport/.test(name);
        case 'reparation': return /reparation|entretien/.test(name);
        case 'scolarite': return /scolarite|educatif/.test(name);
        case 'aumone': return /aumone|don/.test(name);
        case 'renforcement': return /renforcement|aide/.test(name);
        case 'exceptionnel': return /exceptionnel|imprevu/.test(name);
        case 'menage': return /menage/.test(name);
        case 'animal': return /animal|animaux/.test(name);
        case 'alimentaire': return /aliment|pain/.test(name);
        default: return false;
      }
    });
    if (category) return { name: category.name, recognized: true };
  }
  return categoryResolution(description, categories);
}

function parseHouseholdSheet(sheetName, rows, categories, year) {
  const tenantHeaderIndex = rows.findIndex((row) => tenantHeader(row.values));
  const expenseHeaderIndex = rows.findIndex((row) => expenseHeader(row.values));
  if (tenantHeaderIndex < 0 || expenseHeaderIndex < 0 || expenseHeaderIndex <= tenantHeaderIndex) return null;
  const tenantsByKey = new Map();
  const tenants = [];
  const rentPayments = [];
  const incomes = [];
  const records = [];
  const savings = [];
  const totals = [];
  const classifiedRows = [];
  const ignoredRows = [];
  const periodWarnings = [];
  const monthNumber = monthForSheet(sheetName);
  const monthKey = monthNumber ? `${year}-${String(monthNumber).padStart(2, '0')}` : null;
  const rentMonth = monthKey
    ? new Date(Date.UTC(year, monthNumber - 2, 1))
    : null;
  const rentMonthKey = rentMonth
    ? `${rentMonth.getUTCFullYear()}-${String(rentMonth.getUTCMonth() + 1).padStart(2, '0')}`
    : null;
  const tenantColumns = tenantHeader(rows[tenantHeaderIndex].values);
  const expenseColumns = expenseHeader(rows[expenseHeaderIndex].values);
  let inheritedFloor = null;

  rows.forEach((row, rowIndex) => {
    const cells = row.values;
    const nonempty = cells.some((value) => !isEmptyCell(value));
    if (!nonempty) {
      classifiedRows.push(rawRowData(sheetName, row, 'blank'));
      return;
    }

    const classes = [];
    const sectionTitle = cells.find((cell) => typeof cell === 'string'
      && /depenses? du mois/.test(normalize(cell)));
    const declaredMonth = sectionTitle ? monthForSheet(sectionTitle) : null;
    if (declaredMonth && monthNumber && declaredMonth !== monthNumber) {
      periodWarnings.push({
        sheetName,
        rowNumber: row.rowNumber,
        sheetMonth: monthKey,
        declaredMonth: `${year}-${String(declaredMonth).padStart(2, '0')}`,
        content: String(sectionTitle).trim(),
      });
    }
    if (rowIndex === tenantHeaderIndex || rowIndex === expenseHeaderIndex) classes.push('header');
    const tenantNameValue = rowIndex > tenantHeaderIndex && rowIndex < expenseHeaderIndex
      ? cells[tenantColumns.name]
      : null;
    const rentValue = rowIndex > tenantHeaderIndex && rowIndex < expenseHeaderIndex
      ? parseAmount(cells[tenantColumns.rent])
      : null;
    const tenantName = String(tenantNameValue ?? '').trim();
    const isTenantTotal = normalize(tenantName) === 'total';
    const isTenant = Boolean(tenantName && !isTenantTotal && rentValue !== null
      && parseAmount(cells[tenantColumns.people]) !== null);
    if (isTenant) {
      inheritedFloor = canonicalFloor(cells[tenantColumns.floor], inheritedFloor);
      const { fullName, firstName, lastName } = splitTenantName(tenantName);
      const peopleCount = parseAmount(cells[tenantColumns.people]);
      const statusText = String(cells[tenantColumns.status] ?? '').trim();
      const status = paymentStatus(statusText);
      const partialAmount = parseAmount(cells[tenantColumns.paidAmount]);
      const sourceKey = `${normalize(inheritedFloor)}|${normalize(fullName)}`;
      const tenant = {
        sourceKey,
        fullName,
        firstName,
        lastName,
        floor: inheritedFloor,
        housingType: String(cells[tenantColumns.housingType] ?? '').trim(),
        peopleCount,
        monthlyRent: rentValue,
        sources: [],
        valid: Boolean(inheritedFloor && firstName && lastName
          && Number.isInteger(peopleCount) && peopleCount > 0),
        issues: [],
      };
      if (!inheritedFloor) tenant.issues.push('Étage absent ou non reconnu.');
      if (!lastName) tenant.issues.push('Nom de famille absent.');
      if (!Number.isInteger(peopleCount) || peopleCount < 1) tenant.issues.push('Nombre de personnes invalide.');
      const existing = tenantsByKey.get(sourceKey);
      if (existing) {
        existing.sources.push({ sheetName, row: row.rowNumber });
        if (existing.monthlyRent !== tenant.monthlyRent
          || existing.peopleCount !== tenant.peopleCount
          || normalize(existing.housingType) !== normalize(tenant.housingType)) {
          existing.issues.push(`Informations différentes dans ${sheetName}, ligne ${row.rowNumber}.`);
        }
      } else {
        tenant.sources.push({ sheetName, row: row.rowNumber });
        tenantsByKey.set(sourceKey, tenant);
        tenants.push(tenant);
      }
      classes.push('tenant');
      rentPayments.push({
        tenantKey: sourceKey,
        tenantName: fullName,
        floor: inheritedFloor,
        monthKey: rentMonthKey,
        receivedMonthKey: monthKey,
        expectedAmount: rentValue,
        paidAmount: status === 'paid' ? rentValue : status === 'unpaid' ? 0 : partialAmount,
        status: status || 'unknown',
        statusText: String(cells[tenantColumns.status] ?? '').trim(),
        paymentDate: null,
        sheetName,
        rowNumber: row.rowNumber,
        valid: Boolean(monthKey && status && (status !== 'partial' || parseAmount(cells[tenantColumns.paidAmount]) !== null)),
        issues: status ? [] : ['Statut de paiement absent ou non reconnu.'],
      });
      if (status === 'partial' && partialAmount === null) {
        rentPayments.at(-1).issues.push('Montant partiellement payé absent.');
      }
    }
    const totalLabelIndex = cells.findIndex((cell) => isTotalLabel(cell));
    if (isTenantTotal || totalLabelIndex >= 0) classes.push('total');
    if (isTenantTotal) {
      totals.push({
        label: tenantName,
        amount: parseAmount(cells[tenantColumns.rent]),
        monthKey,
        kind: 'rent-total',
        sheetName,
        rowNumber: row.rowNumber,
        formulas: row.formulas,
      });
    } else if (totalLabelIndex >= 0) {
      const label = String(cells[totalLabelIndex]).trim();
      const amount = cells.slice(totalLabelIndex + 1).map(parseAmount).find((value) => value !== null) ?? null;
      totals.push({
        label,
        amount,
        monthKey,
        kind: /paiement/.test(normalize(label))
          ? 'rent-paid-total'
          : totalLabelIndex === 0 ? 'resources-total' : 'expense-total',
        sheetName,
        rowNumber: row.rowNumber,
        formulas: row.formulas,
      });
    }

    const leftLabel = String(cells[0] ?? '').trim();
    const leftAmount = parseAmount(cells[1]);
    const normalizedLabel = normalize(leftLabel);
    if (rowIndex > tenantHeaderIndex && leftLabel && leftAmount !== null
      && !isTenant && totalLabelIndex < 0) {
      if (/epargne/.test(normalizedLabel)) {
        savings.push({
          name: leftLabel,
          amount: leftAmount,
          monthKey,
          date: null,
          sheetName,
          rowNumber: row.rowNumber,
          classification: 'savings',
          valueType: 'balance',
          valid: Boolean(monthKey && Number.isFinite(leftAmount) && leftAmount >= 0),
          importable: true,
          reason: 'Enregistré comme solde observé du mois, sans le compter comme un versement.',
        });
        classes.push('savings');
      } else if (/restant|solde/.test(normalizedLabel)) {
        totals.push({
          label: leftLabel,
          amount: leftAmount,
          monthKey,
          kind: /dernier|precedent/.test(normalizedLabel) ? 'opening-balance' : 'balance',
          sheetName,
          rowNumber: row.rowNumber,
          formulas: row.formulas,
        });
        classes.push('total');
      } else if (rowIndex < expenseHeaderIndex) {
        incomes.push({
          title: leftLabel,
          source: leftLabel,
          categoryName: leftLabel,
          amount: leftAmount,
          monthKey,
          date: null,
          sheetName,
          rowNumber: row.rowNumber,
          valid: Boolean(monthKey),
          issues: monthKey ? [] : ['Mois de rattachement absent du nom de la feuille.'],
        });
        classes.push('income');
      }
    }

    if (rowIndex > expenseHeaderIndex) {
      const expenseLabel = String(cells[expenseColumns.category] ?? '').trim();
      const expenseAmount = parseAmount(cells[expenseColumns.amount]);
      if (expenseLabel && isTotalLabel(expenseLabel)) {
        classes.push('total');
      } else if (expenseLabel && expenseAmount !== null) {
        const category = categoryFromDescription(expenseLabel, categories);
        const isBoutique = normalize(expenseLabel) === 'compte boutique';
        const record = {
          lineNumber: records.length + 1,
          sheetName,
          rowNumber: row.rowNumber,
          title: expenseLabel,
          date: null,
          monthKey,
          amount: expenseAmount,
          type: isBoutique ? 'grocery' : 'expense',
          categoryRaw: expenseLabel,
          categoryName: isBoutique ? 'Boutique' : category.name,
          categoryRecognized: isBoutique || category.recognized,
          quantity: null,
          unit: null,
          unitPrice: null,
          quantityKnown: false,
          unitPriceKnown: false,
          notes: isBoutique ? 'Montant global du compte boutique ; le classeur ne fournit pas le détail des achats.' : '',
          reference: '',
          beneficiary: '',
          paidBy: '',
          paymentMethod: '',
          issues: [
            ...(monthKey ? [] : ['Mois absent du nom de la feuille.']),
            ...(category.recognized || isBoutique ? [] : ['Catégorie inconnue, classée dans Autres.']),
          ],
          valid: Boolean(monthKey && expenseAmount >= 0),
          duplicateSource: null,
        };
        records.push(record);
        classes.push(isBoutique ? 'expense' : 'expense');
      }
    }

    const rowText = normalize(cells.filter((value) => !isEmptyCell(value)).join(' '));
    if (classes.length === 0 && /^(loyers?|depenses?) du mois/.test(rowText)) classes.push('header');
    if (classes.length === 0 && rowIndex < tenantHeaderIndex
      && cells.filter((cell) => !isEmptyCell(cell)).length === 1
      && !cells.some((cell) => parseAmount(cell) !== null)) {
      classes.push('header');
    }
    if (classes.length === 0) {
      const reason = row.formulas.length
        ? 'Ligne avec formule sans libellé exploitable.'
        : 'Aucune structure connue ni champs financiers cohérents.';
      ignoredRows.push(rawRowData(sheetName, row, 'unrecognized', ['unrecognized'], reason));
      classes.push('unrecognized');
    }
    classifiedRows.push(rawRowData(sheetName, row, classes[0], classes));
  });

  return {
    records,
    tenants,
    rentPayments,
    incomes,
    savings,
    totals,
    classifiedRows,
    ignoredRows,
    periodWarnings,
    sheetNames: [sheetName],
  };
}

function categoryResolution(value, categories) {
  const normalized = normalize(value);
  const mappedName = [
    [/aliment|nourriture|ramadan/, 'dépenses alimentaires'],
    [/facture.*eau|eau/, 'facture d’eau'],
    [/electricite/, 'facture d’électricité'],
    [/wifi|wi fi|internet/, 'facture wi-fi'],
    [/mensualite|argent de poche|enfants/, 'mensualités et argent de poche des enfants'],
    [/compte boutique/, 'compte boutique'],
    [/menage/, 'salaire femme de ménage'],
    [/mouton|animaux?/, 'alimentation des animaux'],
    [/pain/, 'pain et achats alimentaires complémentaires'],
    [/renforcement|aide financiere/, 'renforcement et aides financières'],
    [/transport/, 'transport'],
    [/sante|consultation|medicam|ordonnance|dentaire|traitement|yeux|oeil/, 'santé, consultations et médicaments'],
    [/reparation|entretien/, 'réparations et entretien de la maison'],
    [/scolarite|educatif|ecole|pedeg|pedagog/, 'scolarité et frais éducatifs'],
    [/tissu|couture|chaussure|coiffure/, 'autres dépenses'],
    [/aumone|don/, 'aumônes et dons'],
    [/exceptionnel|imprevu/, 'dépenses exceptionnelles ou imprévues'],
  ].find(([pattern]) => pattern.test(normalized))?.[1];
  const known = categories.find((category) => normalize(category.name) === normalized)
    || categories.find((category) => mappedName && normalize(category.name) === normalize(mappedName));
  const alias = CATEGORY_ALIASES.get(normalized);
  const fallback = categories.find((category) => ['autre', 'autres', 'autres depenses'].includes(normalize(category.name)));
  return {
    name: known?.name
      || (alias && categories.find((category) => normalize(category.name).includes(alias))?.name)
      || fallback?.name
      || 'Autres',
    recognized: Boolean(known),
  };
}

function parseHistoricalWorkbook(buffer, categories = []) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true, cellFormula: true });
  const records = [];
  const tenants = [];
  const tenantsByKey = new Map();
  const rentPayments = [];
  const incomes = [];
  const savings = [];
  const totals = [];
  const classifiedRows = [];
  const excludedRows = [];
  const ignoredRows = [];
  const summaryLines = [];
  const periodWarnings = [];

  workbook.SheetNames.forEach((sheetName) => {
    if (normalize(sheetName).includes('recapitulatif')) {
      ignoredRows.push({ sheetName, reason: 'Feuille de synthèse, non importée comme opérations' });
      return;
    }

    const sheet = workbook.Sheets[sheetName];
    const rows = extractWorksheetRows(sheet);
    const sheetMonth = monthForSheet(sheetName);
    const sheetYear = Number(sheetName.match(/\b(20\d{2})\b/)?.[1]) || 2026;
    const householdRows = parseHouseholdSheet(sheetName, rows, categories, sheetYear);
    if (householdRows) {
      records.push(...householdRows.records);
      householdRows.tenants.forEach((tenant) => {
        const existing = tenantsByKey.get(tenant.sourceKey);
        if (existing) {
          existing.sources.push(...tenant.sources);
          if (existing.monthlyRent !== tenant.monthlyRent
            || existing.peopleCount !== tenant.peopleCount
            || normalize(existing.housingType) !== normalize(tenant.housingType)) {
            existing.issues.push(`Informations différentes dans ${sheetName}.`);
          }
        } else {
          tenantsByKey.set(tenant.sourceKey, tenant);
          tenants.push(tenant);
        }
      });
      rentPayments.push(...householdRows.rentPayments);
      incomes.push(...householdRows.incomes);
      savings.push(...householdRows.savings);
      totals.push(...householdRows.totals);
      classifiedRows.push(...householdRows.classifiedRows);
      ignoredRows.push(...householdRows.ignoredRows);
      periodWarnings.push(...householdRows.periodWarnings);
      return;
    }

    let columns = null;
    let section = '';

    rows.forEach((row) => {
      const cells = row.values;
      if (!cells.some((cell) => !isEmptyCell(cell))) {
        classifiedRows.push(rawRowData(sheetName, row, 'blank'));
        return;
      }
      const rowText = normalize(cells.filter((cell) => cell !== null).join(' '));
      const detectedColumns = mapHeaders(cells);
      if (detectedColumns) {
        columns = detectedColumns;
        classifiedRows.push(rawRowData(sheetName, row, 'header'));
        return;
      }

      const hasNumericCell = cells.some((cell) => parseAmount(cell) !== null);
      if (!hasNumericCell && /compte boutique|ravitaillement|marchandise|depenses ordinaires|depenses diverses/.test(rowText)) {
        section = /boutique|ravitaillement|marchandise/.test(rowText) ? 'grocery' : 'ordinary';
        classifiedRows.push(rawRowData(sheetName, row, 'header'));
        return;
      }
      if (!columns) {
        const label = cells.find((cell) => !isEmptyCell(cell));
        if (isTotalLabel(label)) {
          totals.push({
            label: String(label),
            amount: cells.map(parseAmount).find((amount) => amount !== null) ?? null,
            monthKey: sheetMonth ? `${sheetYear}-${String(sheetMonth).padStart(2, '0')}` : null,
            kind: 'total',
            sheetName,
            rowNumber: row.rowNumber,
            formulas: row.formulas,
          });
          classifiedRows.push(rawRowData(sheetName, row, 'total'));
        } else {
          const reason = rowText.includes('|')
            ? 'Ligne texte délimitée mais son schéma ne correspond à aucun tableau connu.'
            : 'Aucun en-tête ou tableau connu ne correspond à cette ligne.';
          const unrecognized = rawRowData(sheetName, row, 'unrecognized', ['unrecognized'], reason);
          ignoredRows.push(unrecognized);
          classifiedRows.push(unrecognized);
        }
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
      const date = parseDate(dateValue, { month: sheetMonth, year: sheetYear });
      const monthValue = normalize(get('month'));
      const monthKey = /^\d{4} (0[1-9]|1[0-2])$/.test(monthValue.replace('-', ' '))
        ? monthValue.replace(' ', '-')
        : date?.slice(0, 7);
      let amount = parseAmount(get('amount'));
      const itemDetails = quantity !== null || unit !== '' || unitPrice !== null;
      const isBoutiqueContext = section === 'grocery' || /boutique|ravitaillement|marchandise|epicerie/.test(rowContext);
      const isBoutique = isBoutiqueContext || itemDetails;
      const isBoutiqueSummary = isBoutiqueContext && !itemDetails
        && (!titleNorm || /^(compte boutique|boutique|total(?: .*)?|sous total(?: .*)?)$/.test(titleNorm));
      const isTotalRow = /^(total|sous total|total general|total mensuel)$/.test(titleNorm);

      if (isBoutiqueSummary) {
        summaryLines.push({ sheetName, row: row.rowNumber, date, amount });
        classifiedRows.push(rawRowData(sheetName, row, 'total'));
        return;
      }
      if (isTotalRow) {
        totals.push({
          label: title,
          amount,
          monthKey: monthKey || date?.slice(0, 7),
          kind: 'total',
          sheetName,
          rowNumber: row.rowNumber,
          formulas: row.formulas,
        });
        classifiedRows.push(rawRowData(sheetName, row, 'total'));
        return;
      }

      const issues = [];
      if (!date && !monthKey) issues.push('Date et mois de rattachement absents ou illisibles');
      if (!title) issues.push('Intitulé absent');
      if (amount === null && quantity !== null && unitPrice !== null) amount = quantity * unitPrice;
      if (amount === null) issues.push('Montant absent ou illisible');
      if (!date && monthKey) issues.push('Date exacte absente ; le mois est conservé.');
      if (amount !== null && quantity !== null && unitPrice !== null
        && Math.abs(amount - quantity * unitPrice) > 1e-9) {
        issues.push('Le montant ne correspond pas à quantité × prix unitaire');
      }

      const resolution = isBoutique
        ? { name: 'Boutique', recognized: true }
        : categoryFromDescription(categoryValue || title, categories);
      if (!isBoutique && !resolution.recognized) issues.push('Catégorie inconnue, classée dans Autres');
      const record = {
        lineNumber: records.length + 1,
        sheetName,
        rowNumber: row.rowNumber,
        title,
        date,
        monthKey,
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
        beneficiary: String(get('beneficiary') ?? '').trim(),
        paidBy: String(get('paidBy') ?? '').trim(),
        paymentMethod: String(get('paymentMethod') ?? '').trim(),
        issues,
        valid: Boolean((date || monthKey) && title && amount !== null),
        duplicateSource: null,
      };

      if (date && date < START_DATE) {
        excludedRows.push({ ...record, reason: 'Antérieure au 01/01/2026' });
        classifiedRows.push(rawRowData(sheetName, row, 'total'));
      } else {
        records.push(record);
        classifiedRows.push(rawRowData(sheetName, row, record.type));
      }
    });
  });

  for (let index = 0; index < totals.length; index += 1) {
    const total = totals[index];
    if (total.kind !== 'expense-total') continue;
    const detailTotal = records
      .filter((record) => record.valid && record.monthKey === total.monthKey)
      .reduce((sum, record) => sum + record.amount, 0);
    totals[index] = {
      ...total,
      detailTotal,
      mismatch: total.amount === null || Math.abs(total.amount - detailTotal) > 1e-9,
    };
  }

  const summaryReports = summaryLines.map((summary) => {
    const details = records.filter((record) => record.valid
      && record.type === 'grocery'
      && (record.monthKey || record.date?.slice(0, 7)) === (summary.monthKey || summary.date?.slice(0, 7)));
    const detailTotal = details.reduce((sum, record) => sum + record.amount, 0);
    const unresolved = summary.amount !== null && details.length === 0 && Math.abs(summary.amount) > 1e-9;
    const mismatch = summary.amount !== null && details.length > 0
      && Math.abs(summary.amount - detailTotal) > 1e-9;
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
    tenants,
    rentPayments,
    incomes,
    savings,
    totals,
    classifiedRows,
    excludedRows,
    ignoredRows,
    summaryLines: summaryReports,
    periodWarnings,
    sheetNames: workbook.SheetNames,
  };
}

function parseHistoricalFile(buffer, fileName, categories = []) {
  const extension = String(fileName || '').toLowerCase().split('.').pop();
  if (extension !== 'csv') return parseHistoricalWorkbook(buffer, categories);
  const firstLine = buffer.toString('utf8').split(/\r?\n/, 1)[0] || '';
  const delimiters = [',', ';', '\t'];
  const delimiter = delimiters.reduce((selected, candidate) => (
    firstLine.split(candidate).length > firstLine.split(selected).length ? candidate : selected
  ), ',');
  const rows = parseCsv(buffer, {
    bom: true,
    delimiter,
    relax_column_count: true,
    skip_empty_lines: true,
    trim: true,
  });
  if (!rows.length) throw new Error('Le fichier CSV est vide.');
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), 'Données CSV');
  return parseHistoricalWorkbook(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }), categories);
}

function signature(record) {
  const operationPeriod = record.date || record.monthKey;
  if (record.valid === false || !operationPeriod || record.amount === null || record.amount === undefined) return null;
  return [
    record.type,
    normalize(record.title),
    operationPeriod,
    record.monthKey || record.date?.slice(0, 7) || '',
    normalize(record.type === 'grocery' ? 'boutique' : (record.categoryName || record.category || '')),
    Number(record.amount).toString(),
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
  const recordYears = records
    .filter((record) => record.valid && /^\d{4}-\d{2}-\d{2}$/.test(record.date || ''))
    .map((record) => Number(record.date.slice(0, 4)));
  const years = [...new Set(recordYears)].sort((first, second) => first - second);
  if (!years.length) years.push(2026);
  const months = years.flatMap((year) => Array.from({ length: 12 }, (_, index) => ({
    month: monthKey(year, index + 1),
    label: `${MONTHS[index]} ${year}`,
    ordinaryCount: 0,
    groceryCount: 0,
    totalBoutique: 0,
    totalOther: 0,
    totalGeneral: 0,
  })));
  const byMonth = new Map(months.map((month) => [month.month, month]));
  records.filter((record) => record.valid && !record.duplicateSource).forEach((record) => {
    const month = byMonth.get(record.monthKey || record.date.slice(0, 7));
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

function styleRow(sheet, rowIndex, columnCount, style) {
  for (let column = 0; column < columnCount; column += 1) {
    const address = XLSX.utils.encode_cell({ r: rowIndex, c: column });
    if (!sheet[address]) sheet[address] = { t: 's', v: '' };
    sheet[address].s = style;
  }
}

function styleBodyRows(sheet, firstRow, lastRow, columnCount, amountColumns = []) {
  const widths = (sheet['!cols'] || []).map((column) => Number(column.wch) || 14);
  const amountColumnSet = new Set(amountColumns);
  for (let rowIndex = firstRow; rowIndex <= lastRow; rowIndex += 1) {
    let lineCount = 1;
    for (let column = 0; column < columnCount; column += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: column });
      const cell = sheet[address];
      if (!cell) continue;
      const text = String(cell.v ?? '').trim();
      const width = widths[column] || 14;
      const lines = text.split(/\r?\n/).reduce(
        (total, line) => total + Math.max(1, Math.ceil(line.length / Math.max(1, width - 2))),
        0,
      );
      lineCount = Math.max(lineCount, lines);
      cell.s = {
        font: { name: 'Aptos', sz: 10, color: { rgb: '334155' } },
        fill: {
          patternType: 'solid',
          fgColor: { rgb: rowIndex % 2 ? 'FFFFFF' : 'F4F8F4' },
        },
        alignment: {
          horizontal: amountColumnSet.has(column) ? 'right' : 'left',
          vertical: 'center',
          wrapText: true,
        },
        border: { bottom: { style: 'thin', color: { rgb: 'E2E8E2' } } },
      };
    }
    sheet['!rows'] ||= [];
    sheet['!rows'][rowIndex] = {
      ...sheet['!rows'][rowIndex],
      hpt: Math.max(Number(sheet['!rows'][rowIndex]?.hpt) || 0, lineCount * 14 + 4),
    };
  }
}

function appendDataSheet(workbook, { name, headers, rows, dateColumns = [], amountColumns = [] }) {
  const dateColumnSet = new Set(dateColumns);
  const convertedRows = rows.map((row) => row.map((value, index) => (
    dateColumnSet.has(index) && /^\d{4}-\d{2}-\d{2}$/.test(String(value || ''))
      ? excelDateSerial(String(value))
      : value
  )));
  const data = [headers, ...convertedRows];
  const sheet = XLSX.utils.aoa_to_sheet(data);
  const columnCount = headers.length;
  styleRow(sheet, 0, columnCount, HEADER_STYLE);
  sheet['!autofilter'] = {
    ref: `A1:${XLSX.utils.encode_col(columnCount - 1)}${Math.max(1, data.length)}`,
  };
  sheet['!cols'] = headers.map((header, columnIndex) => {
    const values = [header, ...convertedRows.map((row) => row[columnIndex])];
    const width = values.reduce((max, value) => Math.max(
      max,
      ...String(value ?? '').split(/\r?\n/).map((line) => line.length),
    ), 12) + 2;
    return { wch: Math.min(36, Math.max(14, width)) };
  });
  styleBodyRows(sheet, 1, data.length - 1, columnCount, amountColumns);
  sheet['!rows'] ||= [];
  sheet['!rows'][0] = { ...sheet['!rows'][0], hpt: Math.max(sheet['!rows'][0]?.hpt || 0, 30) };
  sheet['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' };
  const bodyRows = Array.from({ length: rows.length }, (_, index) => index + 1);
  setNumberFormat(sheet, bodyRows, dateColumns, 'dd/mm/yyyy');
  setNumberFormat(sheet, bodyRows, amountColumns, '#,##0.## "FCFA"');
  XLSX.utils.book_append_sheet(workbook, sheet, name);
}

const HEADER_STYLE = {
  font: { name: 'Aptos', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
  fill: { patternType: 'solid', fgColor: { rgb: '264B47' } },
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  border: { bottom: { style: 'thin', color: { rgb: 'E2E8E2' } } },
};

const SECTION_STYLE = {
  font: { name: 'Aptos', sz: 12, bold: true, color: { rgb: 'FFFFFF' } },
  fill: { patternType: 'solid', fgColor: { rgb: '264B47' } },
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  border: { bottom: { style: 'thin', color: { rgb: 'E2E8E2' } } },
};

const TOTAL_STYLE = {
  font: { name: 'Aptos', sz: 11, bold: true, color: { rgb: '264B47' } },
  fill: { patternType: 'solid', fgColor: { rgb: 'D9EAD3' } },
  border: { top: { style: 'thin', color: { rgb: '264B47' } } },
};

function excelDateSerial(date) {
  const [year, month, day] = date.split('-').map(Number);
  return Date.UTC(year, month - 1, day) / 86_400_000 + 25_569;
}

function buildAnnualWorkbook(year, expenses, groceryItems, databaseTotals = [], financialData = {}) {
  const workbook = XLSX.utils.book_new();
  const annual = [];
  const totalsByMonth = new Map(databaseTotals.map((row) => [row.month, row]));

  for (let month = 1; month <= 12; month += 1) {
    const key = monthKey(year, month);
    const monthExpenses = expenses.filter((row) => (
      row.month_key || row.expense_date?.slice(0, 7)
    ) === key);
    const monthGroceries = groceryItems.filter((row) => (
      row.month_key || row.monthKey || row.purchase_date?.slice(0, 7)
    ) === key);
    const itemBoutiqueTotal = monthGroceries.reduce((sum, row) => sum + Number(row.total_amount || 0), 0);
    const itemOtherTotal = monthExpenses.reduce((sum, row) => sum + Number(row.amount || 0), 0);
    const databaseMonth = totalsByMonth.get(key);
    const totalBoutique = databaseMonth ? Number(databaseMonth.totalBoutique) : itemBoutiqueTotal;
    const totalOther = databaseMonth ? Number(databaseMonth.totalOther) : itemOtherTotal;
    if (Math.abs(totalBoutique - itemBoutiqueTotal) > 1e-9
      || Math.abs(totalOther - itemOtherTotal) > 1e-9) {
      throw new Error(`Le total SQL de ${MONTHS[month - 1]} ne correspond pas aux lignes exportées.`);
    }
    const label = MONTHS[month - 1];
    annual.push({ label, totalBoutique, totalOther, totalGeneral: totalBoutique + totalOther });

    const consolidatedDate = monthGroceries.length
      ? monthGroceries.map((row) => row.purchase_date).sort()[0]
      : null;
    const mainDataRows = [
      ...monthExpenses.map((row) => [
        row.expense_date ? excelDateSerial(row.expense_date) : null,
        row.title,
        row.category || 'Autres',
        Number(row.amount),
        'Ordinaire',
      ]),
      ...(monthGroceries.length ? [[
        consolidatedDate ? excelDateSerial(consolidatedDate) : null,
        'COMPTE BOUTIQUE',
        'Boutique',
        totalBoutique,
        'Boutique',
      ]] : []),
    ].sort((first, second) => first[0] - second[0]);
    const mainRows = [
      ['Date', 'Intitulé de la dépense', 'Catégorie', 'Montant (FCFA)', 'Type de dépense'],
      ...mainDataRows,
    ];
    const mainTotalRow = mainRows.length;
    mainRows.push(['', 'TOTAL DU MOIS', '', {
      f: mainTotalRow > 1 ? `SUM(D2:D${mainTotalRow})` : '0',
      v: totalBoutique + totalOther,
    }, '']);
    const sheet = XLSX.utils.aoa_to_sheet(mainRows);
    styleRow(sheet, 0, 5, HEADER_STYLE);
    styleRow(sheet, mainTotalRow, 5, TOTAL_STYLE);
    sheet['!autofilter'] = { ref: `A1:E${Math.max(1, mainTotalRow)}` };
    sheet['!cols'] = [{ wch: 15 }, { wch: 34 }, { wch: 25 }, { wch: 18 }, { wch: 21 }];
    sheet['!rows'] = [];
    styleBodyRows(sheet, 1, mainTotalRow - 2, 5, [3]);
    sheet['!rows'][0] = { hpt: 30 };
    sheet['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' };
    setNumberFormat(sheet, Array.from({ length: mainTotalRow - 1 }, (_, index) => index + 1), [3], '#,##0.## "FCFA"');
    setNumberFormat(sheet, Array.from({ length: mainDataRows.length }, (_, index) => index + 1), [0], 'dd/mm/yyyy');
    setCell(sheet, mainTotalRow - 1, 3, totalBoutique + totalOther, '#,##0.## "FCFA"');
    sheet['!merges'] = [];
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
  styleRow(summarySheet, 0, 4, SECTION_STYLE);
  styleRow(summarySheet, 1, 4, HEADER_STYLE);
  styleRow(summarySheet, annualTotalRow - 1, 4, TOTAL_STYLE);
  summarySheet['!autofilter'] = { ref: 'A2:D14' };
  summarySheet['!cols'] = [{ wch: 24 }, { wch: 24 }, { wch: 27 }, { wch: 25 }];
  summarySheet['!rows'] = [];
  styleBodyRows(summarySheet, 2, annualTotalRow - 2, 4, [1, 2, 3]);
  summarySheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
  summarySheet['!rows'][0] = { hpt: 30 };
  summarySheet['!rows'][1] = { hpt: 30 };
  summarySheet['!freeze'] = { xSplit: 0, ySplit: 2, topLeftCell: 'A3', activePane: 'bottomLeft', state: 'frozen' };
  setNumberFormat(summarySheet, Array.from({ length: 13 }, (_, index) => index + 2), [1, 2, 3], '#,##0.## "FCFA"');
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'RÉCAPITULATIF ANNUEL');
  workbook.SheetNames.unshift(workbook.SheetNames.pop());
  (financialData.sheets || []).forEach((sheet) => appendDataSheet(workbook, sheet));
  workbook.Workbook = { Views: [{ activeTab: 0 }] };
  workbook.Props = { Title: `Dépenses maison ${year}`, Subject: `Récapitulatif annuel ${year}`, Author: 'Maison' };
  return { workbook, annual };
}

function buildMonthlyWorkbook(year, month, finance, sheets) {
  const workbook = XLSX.utils.book_new();
  appendDataSheet(workbook, {
    name: 'BILAN MENSUEL',
    headers: ['Indicateur', 'Montant (FCFA)'],
    rows: [
      ['Solde reporté', finance.openingBalance],
      ['Recettes encaissées', finance.incomeTotal],
      ['Ressources disponibles', finance.resourcesAvailable],
      ['Loyers attendus', finance.rentExpected],
      ['Loyers encaissés', finance.rentReceived],
      ['Loyers impayés', finance.rentOutstanding],
      ['Dépenses', finance.totalExpenses],
      ['Épargne cumulée', finance.savingsBalance],
      ['Solde final', finance.closingBalance],
    ],
    amountColumns: [1],
  });
  sheets.forEach((sheet) => appendDataSheet(workbook, sheet));
  workbook.Workbook = { Views: [{ activeTab: 0 }] };
  workbook.Props = {
    Title: `Bilan mensuel maison ${year}-${String(month).padStart(2, '0')}`,
    Subject: `Bilan financier ${year}-${String(month).padStart(2, '0')}`,
    Author: 'Maison',
  };
  return workbook;
}

module.exports = {
  START_DATE,
  END_DATE,
  MONTHS,
  buildAnnualWorkbook,
  buildMonthlyWorkbook,
  categoryResolution,
  markDuplicates,
  normalize,
  parseAmount,
  parseDate,
  parseHistoricalFile,
  parseHistoricalWorkbook,
  summarizeRecords,
};
