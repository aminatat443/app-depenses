const test = require('node:test');
const assert = require('node:assert/strict');
const XLSX = require('xlsx-js-style');
const {
  buildAnnualWorkbook,
  buildMonthlyWorkbook,
  markDuplicates,
  parseAmount,
  parseDate,
  parseHistoricalFile,
  parseHistoricalWorkbook,
  summarizeRecords,
} = require('../src/excel');

const categories = [
  { name: 'Électricité' },
  { name: 'Eau' },
  { name: 'Internet / Wi-Fi' },
  { name: 'Santé / Ordonnances' },
  { name: 'Transport' },
  { name: 'Entretien' },
  { name: 'Autres' },
];

const makeWorkbookBuffer = (sheets) => {
  const workbook = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name, rows]) => {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), name);
  });
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
};

test('monthly summary export shows the carryover and available resources separately', () => {
  const workbook = buildMonthlyWorkbook(2026, 4, {
    openingBalance: 10_650,
    incomeTotal: 680_000,
    resourcesAvailable: 690_650,
    rentExpected: 0,
    rentReceived: 0,
    rentOutstanding: 0,
    totalExpenses: 630_000,
    savingsBalance: 0,
    closingBalance: 60_650,
  }, []);
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets['BILAN MENSUEL'], {
    header: 1,
  });

  assert.deepEqual(rows.slice(1, 4), [
    ['Solde reporté', 10_650],
    ['Recettes encaissées', 680_000],
    ['Ressources disponibles', 690_650],
  ]);
  assert.deepEqual(rows.find((row) => row[0] === 'Solde final'), ['Solde final', 60_650]);
});

test('parses different sheet headers, imports all 2026 months and later, and falls back to Autres', () => {
  const buffer = makeWorkbookBuffer({
    Janvier: [
      ['Date', 'Intitulé', 'Type', 'Montant'],
      ['05/01/2026', 'Facture électricité', 'Électricité', '25 000 FCFA'],
      ['06/01/2026', 'Achat spécial', 'Catégorie inconnue', '1 500'],
      ['06/01/2026', 'Achat spécial', 'Catégorie inconnue', '1 500'],
      ['07/01/2026', 'Ligne sans montant', 'Eau', null],
      ['08/09/2025', 'Dépense antérieure', 'Eau', '700'],
    ],
    'Achats boutique': [
      ['Date achat', 'Marchandise', 'Quantité', 'Unité', 'Prix unitaire', 'Montant'],
      ['2026-01-08', 'Riz', 2, 'sac', 2_500, 5_000],
      ['2026-01-08', 'COMPTE BOUTIQUE', null, null, null, 5_000],
      ['2026-08-03', 'Huile', null, 'bidon', null, 7_000],
    ],
    'Septembre 2026': [
      ['Date', 'Désignation', 'Catégorie', 'Montant'],
      ['12/09/2026', 'Achat septembre', 'Autres', 900],
    ],
  });

  const parsed = parseHistoricalWorkbook(buffer, categories);
  const duplicatesMarked = markDuplicates(parsed.records, []);
  const records = duplicatesMarked.filter((row) => row.valid);

  assert.equal(parsed.sheetNames.length, 3);
  assert.equal(parsed.excludedRows.length, 1);
  assert.equal(parsed.summaryLines.length, 1);
  assert.equal(parsed.summaryLines[0].mismatch, false);
  assert.equal(parsed.records.find((row) => row.title === 'Achat spécial').categoryName, 'Autres');
  assert.equal(parsed.records.find((row) => row.title === 'Achat spécial').categoryRecognized, false);
  assert.equal(parsed.records.find((row) => row.title === 'Facture électricité').categoryName, 'Électricité');
  assert.equal(parsed.records.find((row) => row.title === 'Achat septembre').date, '2026-09-12');
  assert.equal(parsed.records.find((row) => row.title === 'Ligne sans montant').valid, false);
  assert.equal(records.filter((row) => row.type === 'grocery').length, 2);
  assert.equal(records.find((row) => row.title === 'Huile').quantityKnown, false);
  assert.equal(records.find((row) => row.title === 'Huile').unitPriceKnown, false);
  assert.equal(duplicatesMarked.filter((row) => row.duplicateSource === 'fichier').length, 1);

  const monthly = summarizeRecords(duplicatesMarked);
  assert.equal(monthly.length, 12);
  assert.equal(monthly[0].totalBoutique, 5_000);
  assert.equal(monthly[0].totalOther, 26_500);
  assert.equal(monthly[0].totalGeneral, monthly[0].totalBoutique + monthly[0].totalOther);
  assert.equal(monthly[7].totalBoutique, 7_000);
  assert.equal(monthly[8].totalOther, 900);
});

test('detects existing database duplicates and supports French dates and amounts', () => {
  assert.equal(parseDate('5 février 2026'), '2026-02-05');
  assert.equal(parseDate(12, { month: 8 }), '2026-08-12');
  assert.equal(parseDate(12, { month: 8, year: 2027 }), '2027-08-12');
  assert.equal(parseAmount('1 234 500 FCFA'), 1_234_500);
  assert.equal(parseAmount('1.234,50'), 1_234.5);

  const [record] = markDuplicates([{
    valid: true,
    type: 'expense',
    title: 'Eau',
    date: '2026-02-01',
    amount: 1250,
  }], [{
    type: 'expense',
    title: 'Eau',
    date: '2026-02-01',
    amount: 1250,
  }]);
  assert.equal(record.duplicateSource, 'base');
});

test('parses semicolon CSV and preserves the accounting month and expense details', () => {
  const parsed = parseHistoricalFile(Buffer.from([
    'Date;Mois;Désignation;Catégorie;Montant;Bénéficiaire;Personne ayant payé;Mode de paiement;Référence',
    '2026-01-03;2026-02;Consultation;Santé / Ordonnances;15000;Clinique;Aminata;Espèces;R-17',
  ].join('\n')), 'operations.csv', categories);
  const [record] = parsed.records;
  assert.equal(record.valid, true);
  assert.equal(record.monthKey, '2026-02');
  assert.equal(record.categoryName, 'Santé / Ordonnances');
  assert.equal(record.beneficiary, 'Clinique');
  assert.equal(record.paidBy, 'Aminata');
  assert.equal(record.paymentMethod, 'Espèces');
  assert.equal(record.reference, 'R-17');
});

test('recognizes boutique sections and blocks unitemized monthly totals', () => {
  const buffer = makeWorkbookBuffer({
    Janvier: [
      ['Compte boutique'],
      ['Date', 'Intitulé', 'Montant'],
      ['2026-01-03', 'Banane', 1000],
      ['2026-01-03', 'Total', 1000],
    ],
    Avril: [
      ['Date', 'Intitulé', 'Catégorie', 'Montant'],
      ['2026-04-03', 'Total', 'Boutique', 5000],
    ],
  });
  const parsed = parseHistoricalWorkbook(buffer, categories);
  assert.equal(parsed.records[0].type, 'grocery');
  assert.equal(parsed.summaryLines[0].mismatch, false);
  assert.equal(parsed.summaryLines[1].unresolved, true);
});

test('parses household worksheets with tenant, rent, income, savings, totals and expense tables', () => {
  const monthSheet = (expenseHeading) => [
    ['Gestion financière du foyer'],
    ['Niveau', 'Prénom et nom', 'Catégorie de logement', 'Nombre de personnes', 'Montant à payer', 'Statut de paiement'],
    ['Rez-de-chaussée', 'Irene Bassene', 'Un magasin', 1, 15_000, 'Payé'],
    ['', 'Total', '', '', 15_000],
    ['Pension père', 180_000],
    ['Pape', 300_000],
    ['', '', '', expenseHeading],
    ['', '', '', 'Type de dépenses', 'Montant'],
    ['', '', '', 'Facture WiFi', 15_000],
    ['Épargne pension père', 364_000, '', 'Pain', 2_500],
    ['', '', '', 'Compte boutique', 20_000],
    ['', '', '', 'Total', 37_500],
  ];
  const parsed = parseHistoricalWorkbook(makeWorkbookBuffer({
    Janvier: monthSheet('Dépenses du mois de Janvier'),
    Février: monthSheet('Dépenses du mois de Février'),
  }), [
    { name: 'Facture Wi-Fi' },
    { name: 'Pain et achats alimentaires complémentaires' },
    { name: 'Autres dépenses' },
  ]);

  assert.equal(parsed.sheetNames.length, 2);
  assert.equal(parsed.tenants.length, 1);
  assert.equal(parsed.tenants[0].fullName, 'Irene Bassene');
  assert.equal(parsed.tenants[0].floor, 'Rez-de-chaussée');
  assert.equal(parsed.rentPayments.length, 2);
  assert.equal(parsed.rentPayments[0].paidAmount, 15_000);
  assert.equal(parsed.rentPayments[0].paymentDate, null);
  assert.equal(parsed.rentPayments[0].monthKey, '2025-12');
  assert.equal(parsed.rentPayments[0].receivedMonthKey, '2026-01');
  assert.equal(parsed.rentPayments[1].monthKey, '2026-01');
  assert.equal(parsed.rentPayments[1].receivedMonthKey, '2026-02');
  assert.equal(parsed.incomes.length, 4);
  assert.equal(parsed.savings.length, 2);
  assert.equal(parsed.savings[0].valueType, 'balance');
  assert.equal(parsed.savings[0].valid, true);
  assert.equal(parsed.totals.length, 4);
  assert.equal(parsed.records.length, 6);
  assert.equal(parsed.records.find((row) => row.title === 'Facture WiFi').categoryName, 'Facture Wi-Fi');
  assert.equal(parsed.records.find((row) => row.title === 'Compte boutique').type, 'grocery');
  assert.equal(parsed.records.find((row) => row.title === 'Compte boutique').date, null);
  assert.equal(parsed.records.find((row) => row.title === 'Compte boutique').monthKey, '2026-01');
  assert.equal(parsed.ignoredRows.filter((row) => row.values).length, 0);
});

test('builds a 13-sheet annual workbook with consolidated boutique lines only', () => {
  const { workbook, annual } = buildAnnualWorkbook(2026, [
    { title: 'Facture Wi-Fi', expense_date: '2026-01-08', amount: 20_000, category: 'Internet / Wi-Fi' },
    { title: 'Après import manuel', expense_date: '2026-11-02', amount: 3_000, category: 'Autres' },
    { title: 'Hors année', expense_date: '2027-01-01', amount: 50_000, category: 'Autres' },
  ], [
    {
      item_name: 'Riz', purchase_date: '2026-01-05', quantity: 2, quantityKnown: 1,
      unit: 'sac', unit_price: 2_500, unitPriceKnown: 1, total_amount: 5_000,
    },
    {
      item_name: 'Huile', purchase_date: '2026-01-09', quantity: 0, quantityKnown: 0,
      unit: 'bidon', unit_price: 0, unitPriceKnown: 0, total_amount: 7_000,
    },
  ]);
  const summaryName = 'RÉCAPITULATIF ANNUEL';
  assert.equal(workbook.SheetNames.length, 13);
  assert.equal(workbook.SheetNames[0], summaryName);
  assert.equal(annual[0].totalBoutique, 12_000);
  assert.equal(annual[0].totalOther, 20_000);
  assert.equal(annual[10].totalOther, 3_000);
  assert.equal(annual[1].totalGeneral, 0);

  const january = XLSX.utils.sheet_to_json(workbook.Sheets.JANVIER, { header: 1, raw: true });
  const consolidated = january.filter((row) => row[1] === 'COMPTE BOUTIQUE');
  assert.equal(consolidated.length, 1);
  assert.equal(consolidated[0][3], 12_000);
  assert.ok(january.findIndex((row) => row[1] === 'COMPTE BOUTIQUE')
    < january.findIndex((row) => row[1] === 'Facture Wi-Fi'));
  assert.equal(january.filter((row) => row[1] === 'Riz' || row[1] === 'Huile').length, 0);
  assert.ok(!january.some((row) => row[0] === 'DÉTAIL DU COMPTE BOUTIQUE'));
  assert.equal(workbook.Sheets.JANVIER.D4.v, 32_000);
  assert.deepEqual(
    (({ y, m, d }) => ({ y, m, d }))(XLSX.SSF.parse_date_code(workbook.Sheets.JANVIER.A2.v)),
    { y: 2026, m: 1, d: 5 },
  );
  assert.equal(workbook.Sheets[summaryName].D3.v, 32_000);
  assert.equal(workbook.Sheets[summaryName].B15.v, 12_000);

  const roundTrip = XLSX.read(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer' });
  assert.equal(roundTrip.SheetNames.length, 13);
  assert.equal(roundTrip.SheetNames[0], summaryName);
  assert.equal(roundTrip.Sheets.JANVIER.D4.f, 'SUM(D2:D3)');
  assert.equal(roundTrip.Sheets[summaryName].D3.f, 'SUM(B3:C3)');
  assert.equal(roundTrip.Sheets.JANVIER.A2.w, '05/01/2026');
  const styled = XLSX.read(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }), {
    type: 'buffer',
    cellStyles: true,
  });
  assert.equal(styled.Sheets.JANVIER.A1.s.fgColor.rgb, '264B47');
  assert.equal(styled.Sheets[summaryName].A1.s.fgColor.rgb, '264B47');
  assert.equal(workbook.Sheets.JANVIER.B2.s.font.name, 'Aptos');
  assert.equal(workbook.Sheets.JANVIER.D2.s.alignment.horizontal, 'right');
  assert.equal(workbook.Sheets.JANVIER.B2.s.alignment.wrapText, true);
  assert.equal(workbook.Sheets[summaryName].A3.s.font.name, 'Aptos');
});
