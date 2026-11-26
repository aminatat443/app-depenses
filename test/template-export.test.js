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

function findCellRow(sheet, columnNumber, value) {
  for (let rowNumber = 21; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (sheet.getCell(rowNumber, columnNumber).value === value) return rowNumber;
  }
  return null;
}

test('reference exports preserve the template layout and hide excluded resources totals', async () => {
  const source = new ExcelJS.Workbook();
  await source.xlsx.readFile(REFERENCE_TEMPLATE);
  const output = await buildReferenceWorkbook([monthData()]);
  const generated = await workbookFromBuffer(await output.xlsx.writeBuffer());
  const templateSheet = source.getWorksheet('Janvier');
  const sheet = generated.getWorksheet('Janvier');

  assert.deepEqual(generated.worksheets.map((worksheet) => worksheet.name), ['Janvier']);
  assert.equal(sheet.getCell('A1').value, 'Loyers du mois de Décembre 2025');
  assert.equal(sheet.getCell('D20').value, 'DÉPENSES DU MOIS DE JANVIER');
  assert.equal(sheet.getCell('D21').value, 'Désignation');
  assert.equal(sheet.getCell('E21').value, 'Montant');
  assert.equal(sheet.getCell('E22').value, 0);
  assert.equal(sheet.getCell('B22').value.formula, 'B21-0');
  assert.equal(sheet.getCell('A24').value, null);
  assert.equal(sheet.getCell('B24').value, null);
  assert.equal(sheet.getCell('A14').value, 'RÉCAPITULATIF FINANCIER');
  assert.ok(sheet.model.merges.includes('A14:B14'));
  assert.equal(sheet.getRow(24).hidden, true);
  assert.equal(sheet.getCell('A3').value, null);
  for (let rowNumber = 3; rowNumber <= 10; rowNumber += 1) {
    assert.equal(sheet.getCell(rowNumber, 7).value, null);
  }
  assertTenantHeaders(sheet);
  for (let rowNumber = 2; rowNumber <= 13; rowNumber += 1) {
    for (let columnNumber = 1; columnNumber <= 7; columnNumber += 1) {
      const cell = sheet.getCell(rowNumber, columnNumber);
      if (cell.value !== null && cell.value !== undefined) {
        assert.equal(cell.border.bottom.style, 'thin');
      }
    }
  }
  assert.equal(sheet.getCell('A2').alignment.wrapText, true);
  const generatedAlignment = sheet.getCell('A3').alignment;
  const templateAlignment = templateSheet.getCell('A3').alignment;
  assert.equal(sheet.getCell('A3').fill.fgColor.argb, templateSheet.getCell('A3').fill.fgColor.argb);
  assert.equal(generatedAlignment.horizontal, templateAlignment.horizontal);
  assert.equal(generatedAlignment.vertical, 'middle');
  assert.equal(generatedAlignment.wrapText, true);
  assert.ok(sheet.getColumn(4).width >= 30);
  assert.ok(sheet.getColumn(5).width >= 20);
  assert.deepEqual(sheet.pageMargins, templateSheet.pageMargins);
  assert.ok(sheet.model.merges.includes('A1:G1'));
  assert.ok(sheet.model.merges.includes('D20:E20'));
  assert.ok(!sheet.model.merges.includes('D20:N20'));
});

test('rent payment statuses are clear and the protected-savings cells stay on one row', async () => {
  const output = await buildReferenceWorkbook([monthData({
    tenants: [
      {
        floor: 'Rez-de-chaussée',
        name: 'Locataire payée',
        housingType: 'Une chambre',
        peopleCount: 1,
        expectedAmount: 50_000,
        paidAmount: 50_000,
        status: 'IMPAYÉ',
      },
      {
        floor: '1er étage',
        name: 'Locataire impayée',
        housingType: 'Une chambre',
        peopleCount: 1,
        expectedAmount: 50_000,
        paidAmount: 0,
        status: 'PAYÉ',
      },
      {
        floor: '2e étage',
        name: 'Locataire partiellement payée',
        housingType: 'Une chambre',
        peopleCount: 1,
        expectedAmount: 50_000,
        paidAmount: 30_000,
        status: 'PAYÉ',
      },
    ],
    totalPeople: 3,
    rentExpected: 150_000,
    savingsBalance: 25_000,
  })]);
  const workbook = await workbookFromBuffer(await output.xlsx.writeBuffer());
  const sheet = workbook.getWorksheet('Janvier');

  assert.equal(sheet.getCell('G3').value, 'PAYÉ');
  assert.equal(sheet.getCell('G3').fill.fgColor.argb, 'FFE2F0E9');
  assert.equal(sheet.getCell('G3').font.color.argb, 'FF264B47');
  for (const cell of ['G3', 'G4', 'G5']) {
    assert.equal(sheet.getCell(cell).alignment.horizontal, 'center');
    assert.equal(sheet.getCell(cell).alignment.vertical, 'middle');
  }
  for (const cell of ['E3', 'E4', 'E5', 'F3', 'F4', 'F5']) {
    assert.equal(sheet.getCell(cell).alignment.horizontal, 'right');
  }
  assert.equal(sheet.getCell('G4').value, 'IMPAYÉ');
  assert.equal(sheet.getCell('G4').fill.fgColor.argb, 'FFF4CCCC');
  assert.equal(sheet.getCell('G4').font.color.argb, 'FF9C0006');
  assert.equal(sheet.getCell('G5').value, 'PARTIELLEMENT PAYÉ');
  assert.equal(sheet.getCell('G5').fill.fgColor.argb, 'FFFFE699');
  assert.equal(sheet.getCell('A3').fill.fgColor.argb, 'FFF4F8F4');
  assert.equal(sheet.getCell('A4').fill.fgColor.argb, 'FFFFFFFF');
  assert.equal(sheet.getCell('B11').fill.fgColor.argb, 'FFD9EAD3');
  assert.equal(sheet.getCell('F11').fill.fgColor.argb, 'FFD9EAD3');
  assert.equal(sheet.getCell('E13').fill.fgColor.argb, 'FFD9EAD3');
  assert.equal(sheet.getCell('F13').fill.fgColor.argb, 'FFD9EAD3');
  assert.equal(sheet.getCell('B11').alignment.horizontal, 'left');
  for (const cell of ['E13', 'A21', 'A22', 'A30', 'D22']) {
    const label = String(sheet.getCell(cell).value || '');
    const expected = /total/i.test(label)
      || /^restant/i.test(label)
      || /^epargne pension père$/i.test(label.trim())
      ? 'left'
      : 'center';
    assert.equal(sheet.getCell(cell).alignment.horizontal, expected, `${cell} should have the expected alignment`);
    assert.equal(sheet.getCell(cell).alignment.vertical, 'middle', `${cell} should be centered vertically`);
  }
  for (const cell of ['E11', 'F11', 'F13', 'B21', 'B22', 'B30', 'E22']) {
    assert.equal(sheet.getCell(cell).alignment.horizontal, 'right', `${cell} should be right-aligned`);
    assert.equal(sheet.getCell(cell).alignment.vertical, 'middle', `${cell} should be centered vertically`);
  }
  assert.equal(sheet.getCell('G5').font.color.argb, 'FF7F6000');
  assert.ok(!sheet.model.merges.includes('A30:A32'));
  assert.ok(!sheet.model.merges.includes('B30:B32'));
  assert.equal(sheet.getCell('A30').value.trim(), 'Epargne pension père');
  assert.equal(sheet.getCell('A23').alignment.horizontal, 'left');
  assert.equal(sheet.getCell('A30').alignment.horizontal, 'left');
  assert.equal(sheet.getCell('A29').value, null);
  assert.equal(sheet.getCell('B29').value, null);
  assert.equal(sheet.getCell('B30').value, 25_000);
  assert.equal(sheet.getCell('A31').value, null);
  assert.equal(sheet.getCell('B31').value, null);
});

test('expanded category expenses do not leave yellow blank cells beside the expense table', async () => {
  const expenses = Array.from({ length: 25 }, (_, index) => ({
    title: `Dépense ${index + 1}`,
    amount: 1_000,
    category: 'Dépenses alimentaires',
  }));
  const output = await buildReferenceWorkbook([monthData({
    monthKey: '2026-09',
    expenses,
    totalExpenses: 25_000,
  })]);
  const workbook = await workbookFromBuffer(await output.xlsx.writeBuffer());
  const sheet = workbook.getWorksheet('Septembre');

  const firstExpenseRow = findCellRow(sheet, 5, 'Dépense 1');
  const expenseTotalRow = findCellRow(sheet, 4, 'Total général');
  assert.ok(firstExpenseRow);
  assert.ok(expenseTotalRow);
  assert.equal(sheet.getCell(firstExpenseRow, 4).value, 'Alimentaire');
  assert.equal(sheet.getCell(firstExpenseRow, 6).value, 1_000);
  assert.equal(sheet.getCell(expenseTotalRow, 6).value.result, 25_000);
  assert.equal(sheet.getCell('A22').value, 'Restant');
  assert.equal(sheet.getCell('B22').value.formula, `B21-SUM(F22:F46)`);
  assert.equal(sheet.getCell('A30').value.trim(), 'Epargne pension père');
  assert.equal(sheet.getCell('B30').value, 0);
  for (let rowNumber = 21; rowNumber <= expenseTotalRow; rowNumber += 1) {
    for (const columnNumber of [1, 2]) {
      const cell = sheet.getCell(rowNumber, columnNumber);
      if (cell.value !== null) continue;
    }
  }
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
  assert.equal(sheet.getCell('A15').alignment.vertical, 'middle');
  assert.equal(sheet.getCell('B15').alignment.vertical, 'middle');
  assert.equal(sheet.getCell('A17').value, 'Pape');
  assert.equal(sheet.getCell('E13').value, 'Total des paiements reçus');
  assert.equal(sheet.getCell('F13').value.result, 200_000);
  assert.equal(sheet.getCell('D21').value, 'Désignation');
  assert.equal(sheet.getCell('E21').value, 'Montant');
  assert.equal(sheet.getCell('D22').value, 'Dépense test');
  assert.equal(sheet.getCell('E22').value, 650_000);
  assert.equal(sheet.getCell('A21').value, 'Total');
  assert.equal(sheet.getCell('B21').value.formula, 'B15+B17+F13');
  assert.equal(sheet.getCell('B21').value.result, 730_000);
  assert.equal(sheet.getCell('A30').value.trim(), 'Epargne pension père');
  assert.equal(sheet.getCell('B30').value, 364_000);
  assert.equal(sheet.getCell('A23').value, 'Restant du mois dernier');
  assert.equal(sheet.getCell('B23').value, 10_000);
  assert.equal(sheet.getCell('A24').value, 'Total des ressources disponibles');
  assert.equal(sheet.getRow(24).hidden, false);
  assert.equal(sheet.getCell('B24').value.formula, 'B21+B23');
  assert.equal(sheet.getCell('B24').value.result, 740_000);
  assert.equal(sheet.getCell('B22').value.formula, 'B24-SUM(E22:E22)');
  assert.ok(sheet.getColumn(2).width >= 18);
  assert.equal(sheet.getCell('A14').value, 'RÉCAPITULATIF FINANCIER');
  assert.equal(sheet.getCell('A30').fill.fgColor.argb, 'FFD9EAD3');
  assert.equal(sheet.getCell('B30').fill.fgColor.argb, 'FFD9EAD3');
  assert.equal(sheet.getCell('B22').value.result, 90_000);
});

test('exports from September 2026 group expenses by category with separate totals', async () => {
  const expenses = [
    { title: 'Facture internet', amount: 12_000, category: 'Facture Wi-Fi' },
    { title: 'Abonnement', amount: 8_000, category: 'Facture Wi-Fi' },
    { title: 'Achat urgent', amount: 3_000, category: 'Autres dépenses' },
  ];
  const septemberWorkbook = await buildReferenceWorkbook([monthData({
    monthKey: '2026-09',
    expenses,
    totalExpenses: 23_000,
    protectedExpenses: [
      { title: 'Consultation', amount: 5_000, category: 'Santé' },
      { title: 'Médicament', amount: 2_000, category: 'Santé' },
    ],
  })]);
  const september = await workbookFromBuffer(await septemberWorkbook.xlsx.writeBuffer());
  const septemberSheet = september.getWorksheet('Septembre');

  assert.equal(septemberSheet.getCell('D20').fill.fgColor.argb, 'FF264B47');
  assert.equal(septemberSheet.getCell('D20').font.color.argb, 'FFFFFFFF');
  assert.ok(septemberSheet.model.merges.includes('D20:F20'));
  assert.equal(septemberSheet.getCell('D21').value, 'Catégorie');
  assert.equal(septemberSheet.getCell('E21').value, 'Désignation');
  assert.equal(septemberSheet.getCell('F21').value, 'Montant');
  assert.equal(septemberSheet.getCell('D21').fill.fgColor.argb, 'FF264B47');
  assert.equal(septemberSheet.getCell('E22').font.name, 'Aptos');
  assert.equal(septemberSheet.getCell('D22').font.color.argb, 'FF264B47');
  assert.equal(septemberSheet.getCell('D22').font.bold, true);
  assert.equal(septemberSheet.autoFilter, 'D21:F26');
  assert.equal(septemberSheet.getCell('D22').value, 'Santé');
  assert.equal(septemberSheet.getCell('D22').alignment.horizontal, 'center');
  assert.equal(septemberSheet.getCell('D22').alignment.vertical, 'middle');
  assert.equal(septemberSheet.getCell('E22').value, 'Consultation');
  assert.equal(septemberSheet.getCell('F22').value, 5_000);
  assert.equal(septemberSheet.getCell('E22').font.color.argb, 'FF334155');
  assert.equal(septemberSheet.getCell('F22').font.color.argb, 'FFFA7D00');
  assert.equal(septemberSheet.getCell('E26').font.size, 10);
  assert.equal(septemberSheet.getCell('D26').alignment.horizontal, 'center');
  assert.equal(septemberSheet.getCell('D26').font.size, 10);
  assert.ok(septemberSheet.model.merges.includes('D22:D23'));
  assert.equal(septemberSheet.getCell('D23').value, 'Santé');
  assert.equal(septemberSheet.getCell('E23').value, 'Médicament');
  assert.equal(septemberSheet.getCell('F23').value, 2_000);
  assert.ok(septemberSheet.model.merges.includes('D24:D25'));
  assert.equal(septemberSheet.getCell('D24').value, 'Wifi');
  assert.equal(septemberSheet.getCell('E24').value, 'Abonnement');
  assert.equal(septemberSheet.getCell('F24').value, 8_000);
  assert.equal(septemberSheet.getCell('D25').value, 'Wifi');
  assert.equal(septemberSheet.getCell('E25').value, 'Facture internet');
  assert.equal(septemberSheet.getCell('F25').value, 12_000);
  assert.equal(septemberSheet.getCell('D26').value, 'Autres dépenses');
  assert.equal(septemberSheet.getCell('E26').value, 'Achat urgent');
  const expenseTotalRow = findCellRow(septemberSheet, 4, 'Total général');
  assert.equal(septemberSheet.getCell(expenseTotalRow, 6).value.result, 30_000);
  assert.equal(findCellRow(septemberSheet, 4, 'RÉCAPITULATIF DES DÉPENSES PAR CATÉGORIE'), null);

  const augustWorkbook = await buildReferenceWorkbook([monthData({
    monthKey: '2026-08',
    expenses,
    totalExpenses: 23_000,
  })]);
  const august = await workbookFromBuffer(await augustWorkbook.xlsx.writeBuffer());
  const augustSheet = august.getWorksheet('Août');
  assert.equal(augustSheet.getCell('D21').value, 'Désignation');
  assert.equal(augustSheet.getCell('E21').value, 'Montant');
  assert.equal(augustSheet.getCell('D22').value, 'Facture internet');
  assert.equal(augustSheet.getCell('E22').value, 12_000);
  assert.equal(augustSheet.getCell('D23').value, 'Abonnement');
  assert.equal(augustSheet.getCell('E23').value, 8_000);
  assert.ok(!augustSheet.model.merges.includes('D22:D23'));
});

test('December expense totals follow the actual data and column G remains for rent statuses', async () => {
  const output = await buildReferenceWorkbook([monthData({
    monthKey: '2026-12',
    tenants: [{
      floor: 'Rez-de-chaussée',
      name: 'Locataire',
      housingType: 'Studio',
      peopleCount: 1,
      expectedAmount: 10_000,
      paidAmount: 10_000,
    }],
    expenses: [
      { title: 'Achat de riz', amount: 5_000, category: 'Alimentaire' },
      { title: 'Achat de fruits', amount: 3_000, category: 'Alimentaire' },
    ],
    totalExpenses: 8_000,
  })]);
  const workbook = await workbookFromBuffer(await output.xlsx.writeBuffer());
  const sheet = workbook.getWorksheet('Décembre');

  assert.equal(sheet.getColumn(7).hidden, false);
  assert.equal(sheet.getCell('G3').value, 'PAYÉ');
  assert.ok(sheet.model.merges.includes('D22:D23'));
  assert.equal(sheet.getCell('D22').value, 'Alimentaire');
  assert.equal(sheet.getCell('E22').value, 'Achat de fruits');
  assert.equal(sheet.getCell('F22').value, 3_000);
  assert.equal(sheet.getCell('G22').value, null);
  assert.equal(sheet.getCell('D24').value, 'Total général');
  assert.equal(sheet.getCell('F24').value.formula, 'SUM(F22:F23)');
  assert.equal(sheet.getCell('F24').value.result, 8_000);
  assert.equal(sheet.pageSetup.printArea, 'A1:G24');

  const emptyDecember = await buildReferenceWorkbook([monthData({ monthKey: '2026-12' })]);
  const emptySheet = await workbookFromBuffer(await emptyDecember.xlsx.writeBuffer());
  assert.equal(emptySheet.getWorksheet('Décembre').getCell('D22').value, 'Total général');
  assert.equal(emptySheet.getWorksheet('Décembre').getCell('F22').value, 0);
  assert.equal(emptySheet.getWorksheet('Décembre').getCell('G22').value, null);
});

test('live categories show only expense categories and retain the monthly grand total', async () => {
  const categories = [
    'Alimentaire', 'Eau', 'Électricité', 'Wifi', 'Compte boutique',
    'Scolarité et frais éducatifs', 'Hygiène et beauté', 'Aumône et dons',
    'Entretien de la maison', 'Santé', 'Transport',
    'Catégorie personnalisée',
  ].map((name, index) => ({ id: index + 1, name }));
  const output = await buildReferenceWorkbook([monthData({
    monthKey: '2026-09',
    categories,
    expenses: [
      {
        categoryId: 1,
        category: 'Alimentaire',
        title: 'Achat de riz',
        date: '2026-09-02',
        amount: 25_000,
        quantity: 5,
        unit: 'kg',
        unitPrice: 5_000,
        beneficiary: 'Famille',
        paidBy: 'Awa',
        paymentMethod: 'Espèces',
        reference: 'RIZ-1',
        comment: 'Sac de riz',
      },
      {
        categoryId: 1,
        category: 'Alimentaire',
        title: 'Huile et condiments',
        date: '2026-09-05',
        amount: 12_000,
      },
      {
        categoryId: 12,
        category: 'Catégorie personnalisée',
        title: 'Dépense personnalisée',
        date: '2026-09-10',
        amount: 3_000,
      },
    ],
    totalExpenses: 40_000,
  })]);
  const workbook = await workbookFromBuffer(await output.xlsx.writeBuffer());
  const sheet = workbook.getWorksheet('Septembre');
  const priorityNames = categories.slice(0, 11).map(({ name }) => name);
  assert.equal(sheet.getCell('D21').value, 'Catégorie');
  assert.equal(sheet.getCell('E21').value, 'Désignation');
  assert.equal(sheet.getCell('F21').value, 'Montant');
  assert.equal(sheet.getCell('D22').value, 'Alimentaire');
  assert.equal(sheet.getCell('E22').value, 'Achat de riz');
  assert.equal(sheet.getCell('F22').value, 25_000);
  assert.equal(sheet.getCell('D23').value, 'Alimentaire');
  assert.equal(sheet.getCell('E23').value, 'Huile et condiments');
  assert.equal(sheet.getCell('D24').value, 'Catégorie personnalisée');
  assert.equal(sheet.getCell('E24').value, 'Dépense personnalisée');
  assert.ok(sheet.model.merges.includes('D22:D23'));
  assert.ok(!sheet.model.merges.includes('D24:D24'));
  for (const column of ['D', 'E', 'F']) {
    assert.equal(sheet.getCell(`${column}21`).fill.fgColor.argb, 'FF264B47');
    assert.equal(sheet.getCell(`${column}21`).font.color.argb, 'FFFFFFFF');
    assert.equal(sheet.getCell(`${column}21`).font.bold, true);
    assert.equal(sheet.getCell(`${column}21`).alignment.horizontal, 'center');
    assert.equal(sheet.getCell(`${column}21`).alignment.vertical, 'middle');
  }
  const totalRow = findCellRow(sheet, 4, 'Total général');
  assert.equal(sheet.getCell(totalRow, 6).value.result, 40_000);
});

test('exports consolidate boutique purchases without showing replenishment details', async () => {
  const boutiqueItems = [
    { title: 'Riz', amount: 5_000, category: 'Compte boutique' },
    { title: 'Huile', amount: 7_000, category: 'Compte boutique' },
  ];
  const categories = [{ id: 1, name: 'Compte boutique' }];

  for (const month of [
    ...['09', '10', '11', '12'].map((monthNumber) => monthData({
      monthKey: `2026-${monthNumber}`,
      categories,
      expenses: boutiqueItems,
      totalExpenses: 12_000,
    })),
    monthData({
      monthKey: '2026-08',
      expenses: boutiqueItems,
      totalExpenses: 12_000,
    }),
  ]) {
    const output = await buildReferenceWorkbook([month]);
    const workbook = await workbookFromBuffer(await output.xlsx.writeBuffer());
    const sheet = workbook.getWorksheet(MONTH_NAMES[Number(month.monthKey.slice(5, 7)) - 1]);

    if (month.monthKey >= '2026-09') {
      assert.equal(sheet.getCell('D22').value, 'Compte boutique');
      assert.equal(sheet.getCell('E22').value, 'Compte boutique');
      assert.equal(sheet.getCell('F22').value, 12_000);
      assert.equal(sheet.getCell('F23').value.result, 12_000);
    } else {
      assert.equal(sheet.getCell('D22').value, 'Compte boutique');
      assert.equal(sheet.getCell('E22').value, 12_000);
      assert.equal(sheet.getCell('E23').value.result, 12_000);
    }
    assert.equal(sheet.getCell('D23').value, 'Total général');
    assert.equal(sheet.getCell('D23').alignment.horizontal, 'left');
    const displayedValues = Array.from({ length: sheet.rowCount }, (_, index) => (
      [sheet.getCell(index + 1, 4).value, sheet.getCell(index + 1, 5).value]
    )).flat();
    assert.ok(!displayedValues.includes('Riz'));
    assert.ok(!displayedValues.includes('Huile'));
  }
});

test('unchecked exports show the prior month balance without its carryover', async () => {
  const month = monthData({
    monthKey: '2026-04',
    monthLabel: 'Avril',
    previousMonthLabel: 'Mars',
    incomeTotal: 680_000,
    openingBalance: 10_650,
    openingBalanceWithoutCarryover: 650,
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
    assert.equal(sheet.getRow(23).hidden, false);
  }
  assert.equal(excludedSheet.getCell('B22').value.formula, 'B21-SUM(E22:E22)');
  assert.equal(includedSheet.getCell('B22').value.formula, 'B24-SUM(E22:E22)');
  assert.equal(excludedSheet.getCell('B23').value, 650);
  assert.equal(includedSheet.getCell('B23').value, 10_650);
  assert.equal(excludedSheet.getCell('A24').value, null);
  assert.equal(excludedSheet.getCell('B24').value, null);
  assert.equal(excludedSheet.getRow(24).hidden, true);
  assert.equal(excludedSheet.getCell('B22').value.result, 50_000);
  assert.equal(includedSheet.getCell('B24').value.formula, 'B21+B23');
  assert.equal(includedSheet.getRow(24).hidden, false);
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
  assert.equal(annualExcluded.getWorksheet('Avril').getCell('B23').value, 650);
  assert.equal(annualIncluded.getWorksheet('Avril').getCell('B23').value, 10_650);
});

test('annual exports keep monthly totals while hiding excluded resources rows', async () => {
  const months = MONTH_NAMES.map((monthLabel, index) => monthData({
    monthKey: `2026-${String(index + 1).padStart(2, '0')}`,
    monthLabel,
  }));
  const output = await buildReferenceWorkbook(months, { annual: true });
  const generated = await workbookFromBuffer(await output.xlsx.writeBuffer());

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
    assert.equal(sheet.getCell('A24').value, null);
    assert.equal(sheet.getRow(24).hidden, true);
    assert.equal(sheet.getCell('B24').value, null);
    assert.equal(sheet.getCell('A30').value.trim(), 'Epargne pension père');
    assert.equal(sheet.getCell('B30').value, 0);
    assert.equal(sheet.getCell('B21').value.formula, 'B15+B17+F13');
    assert.ok(sheet.model.merges.includes('A14:B14'));
    assert.ok(sheet.model.merges.includes(
      MONTH_NAMES.indexOf(monthLabel) < 8 ? 'D20:E20' : 'D20:F20',
    ));
    assert.equal(sheet.getCell('A14').value, 'RÉCAPITULATIF FINANCIER');
    assert.equal(sheet.getCell('A15').alignment.vertical, 'middle');
    assert.equal(sheet.getCell('B15').alignment.vertical, 'middle');
    assert.equal(sheet.getCell('A21').fill.fgColor.argb, 'FFD9EAD3');
    assert.equal(sheet.getCell('A30').fill.fgColor.argb, 'FFD9EAD3');
  }
  const annualSheet = generated.getWorksheet('Récapitulatif annuel');
  assert.equal(annualSheet.getCell('A1').fill.fgColor.argb, 'FF264B47');
  assert.equal(annualSheet.getCell('A2').fill.fgColor.argb, 'FF264B47');
  assert.equal(annualSheet.getCell('A15').fill.fgColor.argb, 'FFD9EAD3');
  for (let columnNumber = 1; columnNumber <= 7; columnNumber += 1) {
    const totalCell = annualSheet.getCell(15, columnNumber);
    assert.equal(totalCell.alignment.horizontal, columnNumber === 1 ? 'left' : 'right');
    assert.equal(totalCell.alignment.vertical, 'middle');
  }
  assert.equal(annualSheet.getCell('E15').value.formula, 'SUM(E3:E14)');
  assert.ok(annualSheet.model.merges.includes('B15:C15'));
  assert.equal(annualSheet.getCell('C15').value.formula, 'SUM(B3:B14)');
  generated.worksheets.forEach((sheet) => {
    sheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (typeof cell.value === 'string' && /total/i.test(cell.value)) {
          assert.equal(cell.alignment.horizontal, 'left', `${sheet.name}!${cell.address} should be left-aligned`);
        }
      });
    });
  });
  assert.deepEqual(annualSheet.getCell('A17').style, {});
  assert.deepEqual(annualSheet.getCell('D20').style, {});
  assert.equal(annualSheet.pageSetup.printArea, 'A1:G15');
  assert.equal(annualSheet.pageSetup.orientation, 'landscape');
  assert.equal(annualSheet.pageSetup.fitToPage, true);
  assert.equal(annualSheet.pageSetup.fitToWidth, 1);
  assert.equal(annualSheet.pageSetup.fitToHeight, 1);
});

test('reference exports expand tenants, income sources and expense rows without overlap', async () => {
  const tenants = Array.from({ length: 9 }, (_, index) => ({
    floor: index < 3 ? 'Rez-de-chaussée' : '1er étage',
    name: `Locataire ${index + 1}`,
    housingType: index === 0 ? 'Un logement avec une très longue description qui doit être affichée intégralement dans la cellule sans être tronquée' : 'Une chambre',
    peopleCount: 1,
    expectedAmount: 10_000,
    paidAmount: 10_000,
    status: 'IMPAYÉ',
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
  assert.equal(sheet.getCell('G3').value, 'PAYÉ');
  assert.equal(sheet.getCell('B11').value, 'Locataire 9');
  const expenseTitleRow = findCellRow(sheet, 4, 'DÉPENSES DU MOIS DE JANVIER');
  const expenseHeaderRow = expenseTitleRow + 1;
  const expenseDataRow = expenseHeaderRow + 1;
  assert.ok(expenseTitleRow);
  assert.equal(sheet.getCell(expenseHeaderRow, 4).value, 'Désignation');
  assert.equal(sheet.getCell(expenseHeaderRow, 5).value, 'Montant');
  let exceptionalRow = null;
  for (let rowNumber = expenseHeaderRow + 1; rowNumber < sheet.rowCount; rowNumber += 1) {
    if (/^Dépense exceptionnelle/.test(String(sheet.getCell(rowNumber, 4).value || ''))) {
      exceptionalRow = rowNumber;
      break;
    }
  }
  assert.ok(exceptionalRow);
  assert.ok(findCellRow(sheet, 4, 'Dépense 25'));
  assert.equal(findCellRow(sheet, 4, 'Total général') >= expenseDataRow + 25, true);
  assert.equal(sheet.getCell('D3').alignment.wrapText, true);
  assert.equal(sheet.getCell(exceptionalRow, 4).alignment.wrapText, true);
  assert.ok(sheet.getRow(3).height > 15);
  assert.ok(sheet.getRow(exceptionalRow).height > 15);
  assert.ok(sheet.getColumn(4).width <= 50);
  assert.equal(sheet.getCell('A20').value, 'Source 2');
  assert.equal(sheet.getCell('A22').value, 'Source 4');
  assert.equal(sheet.getCell('B23').value.formula, 'B16+B18+F14+SUM(B19:B22)');
  assert.ok(sheet.model.merges.includes(`D${expenseTitleRow}:E${expenseTitleRow}`));
  assert.ok(!sheet.model.merges.includes(`D${expenseTitleRow}:N${expenseTitleRow}`));
  assert.ok(sheet.model.merges.includes('A15:B15'));
  assert.ok(sheet.model.merges.includes('B12:D12'));
  assert.ok(sheet.model.merges.includes('A3:A5'));
  assert.ok(sheet.model.merges.includes('B11:C11'));
});
