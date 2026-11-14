const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const { MONTH_NAMES, REFERENCE_TEMPLATE, buildReferenceWorkbook } = require('../src/template-export');

function monthData(overrides = {}) {
  return {
    monthKey: '2026-01',
    monthLabel: 'Janvier',
    previousMonthLabel: 'Décembre 2025',
    tenants: [],
    totalPeople: 0,
    rentExpected: 0,
    rentPayments: [],
    rentReceived: 0,
    pensionPere: 0,
    pape: 0,
    otherIncomes: [],
    incomeTotal: 0,
    openingBalance: 0,
    resourcesAvailable: 0,
    savingsBalance: 0,
    savingsDeposits: 0,
    savingsWithdrawals: 0,
    closingBalance: 0,
    expenses: [],
    totalExpenses: 0,
    ...overrides,
  };
}

function assertTenantHeaders(sheet) {
  assert.deepEqual(
    ['A2', 'B2', 'D2', 'E2', 'F2', 'G2'].map((address) => sheet.getCell(address).value),
    ['NIVEAU', 'NOM ET PRÉNOM', 'TYPE DE LOGEMENT', 'NOMBRE DE PERSONNES', 'MONTANT À PAYER', 'STATUT DU PAIEMENT'],
  );
}

function assertWrappedTemplateStyle(cell, templateCell) {
  const { alignment: actualAlignment, ...actualStyle } = cell.style;
  const { alignment: templateAlignment, ...templateStyle } = templateCell.style;
  assert.deepEqual(actualStyle, templateStyle);
  assert.equal(actualAlignment.vertical, 'middle');
  assert.equal(actualAlignment.wrapText, true);
  assert.equal(templateAlignment.vertical, 'bottom');
}

async function workbookFromBuffer(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
}

test('reference exports preserve the template layout and calculate empty-month totals', async () => {
  const source = new ExcelJS.Workbook();
  await source.xlsx.readFile(REFERENCE_TEMPLATE);
  const output = await buildReferenceWorkbook([monthData()]);
  const generated = await workbookFromBuffer(await output.xlsx.writeBuffer());
  const templateSheet = source.getWorksheet('Janvier');
  const sheet = generated.getWorksheet('Janvier');

  assert.deepEqual(generated.worksheets.map((worksheet) => worksheet.name), ['Janvier']);
  assert.equal(sheet.getCell('A1').value, 'Loyers du mois de Décembre 2025');
  assert.equal(sheet.getCell('D20').value, 'Dépenses du mois de Janvier');
  assert.equal(sheet.getCell('E42').value.formula, 'SUM(E22:E41)');
  assert.equal(sheet.getCell('B22').value.formula, 'B24-E42');
  assert.equal(sheet.getCell('A24').value, 'Total des ressources disponibles');
  assert.equal(sheet.getCell('B24').value.formula, 'B21');
  assert.equal(Number(sheet.getCell('B24').value.result || 0), 0);
  assert.equal(sheet.getCell('A3').value, null);
  assertTenantHeaders(sheet);
  assert.equal(sheet.getCell('A2').alignment.wrapText, true);
  const { alignment: generatedAlignment, ...generatedStyle } = sheet.getCell('A3').style;
  const { alignment: templateAlignment, ...templateStyle } = templateSheet.getCell('A3').style;
  assert.deepEqual(generatedStyle, templateStyle);
  assert.equal(generatedAlignment.horizontal, templateAlignment.horizontal);
  assert.equal(generatedAlignment.vertical, 'middle');
  assert.equal(generatedAlignment.wrapText, true);
  assert.deepEqual(sheet.getColumn(4).width, templateSheet.getColumn(4).width);
  assert.deepEqual(sheet.pageMargins, templateSheet.pageMargins);
  assert.ok(sheet.model.merges.includes('A1:G1'));
  assert.ok(sheet.model.merges.includes('D20:E20'));
});

test('monthly revenues use the original cells and exclude protected savings from the total', async () => {
  const template = new ExcelJS.Workbook();
  await template.xlsx.readFile(REFERENCE_TEMPLATE);
  const output = await buildReferenceWorkbook([monthData({
    monthKey: '2026-02',
    monthLabel: 'Février',
    previousMonthLabel: 'Janvier',
    pensionPere: 180_000,
    pape: 350_000,
    rentPayments: [{ amount: 150_000 }, { amount: 50_000 }],
    rentReceived: 200_000,
    incomeTotal: 730_000,
    openingBalance: 10_000,
    resourcesAvailable: 740_000,
    savingsBalance: 364_000,
    closingBalance: 90_000,
    expenses: [{ title: 'Dépense test', amount: 650_000 }],
    totalExpenses: 650_000,
  })], { includeCarryover: true });
  const generated = await workbookFromBuffer(await output.xlsx.writeBuffer());
  const sheet = generated.getWorksheet('Février');

  assert.equal(sheet.getCell('B15').value, 180_000);
  assert.equal(sheet.getCell('B17').value, 350_000);
  assert.equal(sheet.getCell('A15').value, 'Pension père');
  assert.equal(sheet.getCell('A17').value, 'Pape');
  assert.equal(sheet.getCell('E13').value, 'Total des paiements reçus');
  assert.equal(sheet.getCell('F13').value.result, 200_000);
  assert.equal(sheet.getCell('A21').value, 'Total');
  assert.equal(sheet.getCell('B21').value.formula, 'B15+B17+F13');
  assert.equal(sheet.getCell('B21').value.result, 730_000);
  assert.equal(sheet.getCell('A30').value.trim(), 'Epargne pension père');
  assert.equal(sheet.getCell('B30').value, 364_000);
  assert.equal(sheet.getCell('A23').value, 'Restant du mois dernier');
  assert.equal(sheet.getCell('B23').value, 10_000);
  assert.equal(sheet.getCell('A24').value, 'Total des ressources disponibles');
  assert.equal(sheet.getCell('B24').value.formula, 'B21+B23');
  assert.equal(sheet.getCell('B24').value.result, 740_000);
  assert.equal(sheet.getCell('B22').value.formula, 'B24-E43');
  assert.ok(sheet.getColumn(2).width >= 18);
  assertWrappedTemplateStyle(sheet.getCell('A30'), template.getWorksheet('Janvier').getCell('A30'));
  assert.deepEqual(sheet.getCell('B30').style, template.getWorksheet('Janvier').getCell('B30').style);
  assert.equal(sheet.getCell('B22').value.result, 90_000);
});

test('carryover is always visible but only included in the exported balance when selected', async () => {
  const month = monthData({
    monthKey: '2026-04',
    monthLabel: 'Avril',
    previousMonthLabel: 'Mars',
    incomeTotal: 680_000,
    openingBalance: 10_650,
    resourcesAvailable: 690_650,
    expenses: [{ title: 'Dépense test', amount: 630_000 }],
    totalExpenses: 630_000,
    closingBalance: 60_650,
  });
  const excludedWorkbook = await buildReferenceWorkbook([month]);
  const includedWorkbook = await buildReferenceWorkbook([month], { includeCarryover: true });
  const excluded = await workbookFromBuffer(await excludedWorkbook.xlsx.writeBuffer());
  const included = await workbookFromBuffer(await includedWorkbook.xlsx.writeBuffer());
  const excludedSheet = excluded.getWorksheet('Avril');
  const includedSheet = included.getWorksheet('Avril');

  for (const sheet of [excludedSheet, includedSheet]) {
    assert.equal(sheet.getCell('A23').value, 'Restant du mois dernier');
    assert.equal(sheet.getCell('B23').value, 10_650);
    assert.equal(sheet.getRow(23).hidden, false);
    assert.equal(sheet.getCell('B22').value.formula.startsWith('B24-E'), true);
  }
  assert.equal(excludedSheet.getCell('B24').value.formula, 'B21');
  assert.equal(excludedSheet.getCell('B24').value.result, 680_000);
  assert.equal(excludedSheet.getCell('B22').value.result, 50_000);
  assert.equal(includedSheet.getCell('B24').value.formula, 'B21+B23');
  assert.equal(includedSheet.getCell('B24').value.result, 690_650);
  assert.equal(includedSheet.getCell('B22').value.result, 60_650);

  const annualMonths = MONTH_NAMES.map((monthLabel, index) => (
    index === 3
      ? month
      : monthData({
        monthKey: `2026-${String(index + 1).padStart(2, '0')}`,
        monthLabel,
      })
  ));
  const annualExcludedOutput = await buildReferenceWorkbook(annualMonths, {
    annual: true,
    includeCarryover: false,
  });
  const annualIncludedOutput = await buildReferenceWorkbook(annualMonths, {
    annual: true,
    includeCarryover: true,
  });
  const annualExcluded = await workbookFromBuffer(await annualExcludedOutput.xlsx.writeBuffer());
  const annualIncluded = await workbookFromBuffer(await annualIncludedOutput.xlsx.writeBuffer());
  assert.equal(annualExcluded.getWorksheet('Récapitulatif annuel').getCell('F6').value.result, 50_000);
  assert.equal(annualIncluded.getWorksheet('Récapitulatif annuel').getCell('F6').value.result, 60_650);
  assert.equal(annualExcluded.getWorksheet('Avril').getCell('A23').value, 'Restant du mois dernier');
  assert.equal(annualIncluded.getWorksheet('Avril').getCell('B23').value, 10_650);
});

test('annual exports include all four financial labels and values on all twelve monthly sheets', async () => {
  const template = new ExcelJS.Workbook();
  await template.xlsx.readFile(REFERENCE_TEMPLATE);
  const months = MONTH_NAMES.map((monthLabel, index) => monthData({
    monthKey: `2026-${String(index + 1).padStart(2, '0')}`,
    monthLabel,
  }));
  const output = await buildReferenceWorkbook(months, { annual: true });
  const generated = await workbookFromBuffer(await output.xlsx.writeBuffer());
  const janTemplate = template.getWorksheet('Janvier');

  assert.equal(generated.worksheets.length, 13);
  for (const monthLabel of MONTH_NAMES) {
    const sheet = generated.getWorksheet(monthLabel);
    assert.ok(sheet, `missing monthly sheet ${monthLabel}`);
    assert.equal(sheet.getCell('A15').value, 'Pension père');
    assert.equal(sheet.getCell('B15').value, 0);
    assert.equal(sheet.getCell('A17').value, 'Pape');
    assert.equal(sheet.getCell('B17').value, 0);
    assert.equal(sheet.getCell('A23').value, 'Restant du mois dernier');
    assert.equal(sheet.getCell('B23').value, 0);
    assert.equal(sheet.getRow(23).hidden, false);
    assert.equal(sheet.getCell('A24').value, 'Total des ressources disponibles');
    assert.equal(sheet.getCell('B24').value.formula, 'B21');
    assert.equal(Number(sheet.getCell('B24').value.result || 0), 0);
    assert.equal(sheet.getCell('A30').value.trim(), 'Epargne pension père');
    assert.equal(sheet.getCell('B30').value, 0);
    assert.equal(sheet.getCell('B21').value.formula, 'B15+B17+F13');
    assertWrappedTemplateStyle(sheet.getCell('A15'), janTemplate.getCell('A15'));
    assert.deepEqual(sheet.getCell('B15').style, janTemplate.getCell('B15').style);
    assertWrappedTemplateStyle(sheet.getCell('A17'), janTemplate.getCell('A17'));
    assert.deepEqual(sheet.getCell('B17').style, janTemplate.getCell('B17').style);
    assertWrappedTemplateStyle(sheet.getCell('A23'), janTemplate.getCell('A23'));
    assert.deepEqual(sheet.getCell('B23').style, janTemplate.getCell('B23').style);
    assertWrappedTemplateStyle(sheet.getCell('A24'), janTemplate.getCell('A21'));
    assert.deepEqual(sheet.getCell('B24').style, janTemplate.getCell('B21').style);
    assertWrappedTemplateStyle(sheet.getCell('A30'), janTemplate.getCell('A30'));
    assert.deepEqual(sheet.getCell('B30').style, janTemplate.getCell('B30').style);
  }
});

test('reference exports expand tenants, income sources and expense rows without overlap', async () => {
  const tenants = Array.from({ length: 9 }, (_, index) => ({
    floor: index < 3 ? 'Rez-de-chaussée' : '1er étage',
    name: `Locataire ${index + 1}`,
    housingType: index === 0 ? 'Un logement avec une très longue description qui doit être affichée intégralement dans la cellule sans être tronquée' : 'Une chambre',
    peopleCount: 1,
    expectedAmount: 10_000,
    status: 'Payé',
  }));
  const expenses = Array.from({ length: 25 }, (_, index) => ({
    title: index === 0 ? `Dépense exceptionnelle ${'avec une désignation très détaillée '.repeat(5)}` : `Dépense ${index + 1}`,
    amount: 1_000,
  }));
  const output = await buildReferenceWorkbook([monthData({
    tenants,
    totalPeople: 9,
    rentExpected: 90_000,
    otherIncomes: Array.from({ length: 4 }, (_, index) => ({ source: `Source ${index + 1}`, amount: 1_000 })),
    incomeTotal: 4_000,
    expenses,
    totalExpenses: 25_000,
    closingBalance: -21_000,
  })]);
  const generated = await workbookFromBuffer(await output.xlsx.writeBuffer());
  const sheet = generated.getWorksheet('Janvier');

  assertTenantHeaders(sheet);
  assert.equal(sheet.getCell('B12').value, 'Total');
  assert.equal(sheet.getCell('E12').value.result, 9);
  assert.equal(sheet.getCell('F12').value.result, 90_000);
  assert.equal(sheet.getCell('A3').value, 'Rez-de-chaussée');
  assert.equal(sheet.getCell('B3').value, 'Locataire 1');
  assert.equal(sheet.getCell('D3').value, tenants[0].housingType);
  assert.equal(sheet.getCell('E3').value, 1);
  assert.equal(sheet.getCell('F3').value, 10_000);
  assert.equal(sheet.getCell('G3').value, 'Payé');
  assert.equal(sheet.getCell('B11').value, 'Locataire 9');
  assert.equal(sheet.getCell('D22').value, 'Dépenses du mois de Janvier');
  assert.equal(sheet.getCell('D23').value, 'Type de depenses');
  assert.match(sheet.getCell('D24').value, /^Dépense exceptionnelle/);
  assert.equal(sheet.getCell('D48').value, 'Dépense 25');
  assert.equal(sheet.getCell('D49').value, 'Total');
  assert.equal(sheet.getCell('E49').value.formula, 'SUM(E24:E48)');
  assert.equal(sheet.getCell('E49').value.result, 25_000);
  assert.equal(sheet.getCell('D3').alignment.wrapText, true);
  assert.equal(sheet.getCell('D24').alignment.wrapText, true);
  assert.ok(sheet.getRow(3).height > 15);
  assert.ok(sheet.getRow(24).height > 15);
  assert.ok(sheet.getColumn(4).width <= 50);
  assert.equal(sheet.getCell('A20').value, 'Source 2');
  assert.equal(sheet.getCell('A22').value, 'Source 4');
  assert.equal(sheet.getCell('B23').value.formula, 'B16+B18+F14+SUM(B19:B22)');
  assert.ok(sheet.model.merges.includes('D22:E22'));
  assert.ok(sheet.model.merges.includes('B12:D12'));
  assert.ok(sheet.model.merges.includes('A3:A5'));
  assert.ok(sheet.model.merges.includes('B11:C11'));
});
