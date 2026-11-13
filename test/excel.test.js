const test = require('node:test');
const assert = require('node:assert/strict');
const XLSX = require('xlsx');
const {
  buildAnnualWorkbook,
  markDuplicates,
  parseAmount,
  parseDate,
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

test('parses different sheet headers, confines imports to January-August 2026 and falls back to Autres', () => {
  const buffer = makeWorkbookBuffer({
    Janvier: [
      ['Date', 'Intitulé', 'Type', 'Montant'],
      ['05/01/2026', 'Facture électricité', 'Électricité', '25 000 FCFA'],
      ['06/01/2026', 'Achat spécial', 'Catégorie inconnue', '1 500'],
      ['06/01/2026', 'Achat spécial', 'Catégorie inconnue', '1 500'],
      ['07/01/2026', 'Ligne sans montant', 'Eau', null],
      ['08/09/2026', 'Dépense septembre', 'Eau', '700'],
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
  assert.equal(parsed.excludedRows.length, 2);
  assert.equal(parsed.summaryLines.length, 1);
  assert.equal(parsed.summaryLines[0].mismatch, false);
  assert.equal(parsed.records.find((row) => row.title === 'Achat spécial').categoryName, 'Autres');
  assert.equal(parsed.records.find((row) => row.title === 'Achat spécial').categoryRecognized, false);
  assert.equal(parsed.records.find((row) => row.title === 'Facture électricité').categoryName, 'Électricité');
  assert.equal(parsed.records.find((row) => row.title === 'Ligne sans montant').valid, false);
  assert.equal(records.filter((row) => row.type === 'grocery').length, 2);
  assert.equal(records.find((row) => row.title === 'Huile').quantityKnown, false);
  assert.equal(records.find((row) => row.title === 'Huile').unitPriceKnown, false);
  assert.equal(duplicatesMarked.filter((row) => row.duplicateSource === 'fichier').length, 1);

  const monthly = summarizeRecords(duplicatesMarked);
  assert.equal(monthly.length, 8);
  assert.equal(monthly[0].totalBoutique, 5_000);
  assert.equal(monthly[0].totalOther, 26_500);
  assert.equal(monthly[0].totalGeneral, monthly[0].totalBoutique + monthly[0].totalOther);
  assert.equal(monthly[7].totalBoutique, 7_000);
});

test('detects existing database duplicates and supports French dates and amounts', () => {
  assert.equal(parseDate('5 février 2026'), '2026-02-05');
  assert.equal(parseDate(12, { month: 8 }), '2026-08-12');
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

test('builds a 13-sheet annual workbook with a consolidated boutique line and item detail', () => {
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
  assert.equal(january.filter((row) => row[1] === 'Riz' || row[1] === 'Huile').length, 2);
  assert.equal(workbook.Sheets.JANVIER.D4.v, 32_000);
  assert.deepEqual(
    (({ y, m, d }) => ({ y, m, d }))(XLSX.SSF.parse_date_code(workbook.Sheets.JANVIER.A2.v)),
    { y: 2026, m: 1, d: 8 },
  );
  assert.equal(workbook.Sheets[summaryName].D3.v, 32_000);
  assert.equal(workbook.Sheets[summaryName].B15.v, 12_000);

  const roundTrip = XLSX.read(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer' });
  assert.equal(roundTrip.SheetNames.length, 13);
  assert.equal(roundTrip.SheetNames[0], summaryName);
});
