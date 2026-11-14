const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const XLSX = require('xlsx-js-style');
const ExcelJS = require('exceljs');

const testDataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'app-depenses-app-test-'));
delete process.env.DATABASE_URL;
process.env.PGHOST = '/var/run/postgresql';
process.env.PGUSER = 'root';
process.env.PGDATABASE = 'app_depenses_test';
process.env.APP_DATA_DIR = path.join(testDataDir, 'data');
const { createSession, hashPassword, SESSION_COOKIE } = require('../src/auth');
const { app, db, initializeDatabase, initialization } = require('../src/app');
let sessionToken;

test.before(async () => {
  await initialization;
  await db.query(`
    TRUNCATE expenses, grocery_items, import_batches, backup_runs, settings, categories, users,
      units, savings_accounts
    RESTART IDENTITY CASCADE
  `);
  await initializeDatabase();
  const user = await db.prepare('INSERT INTO users (email, password_hash) VALUES (?, ?)').run(
    'test@example.com',
    await hashPassword('test-password'),
  );
  sessionToken = await createSession(user.lastInsertRowid);
});

test.after(async () => {
  await db.close();
  fs.rmSync(testDataDir, { recursive: true, force: true });
});

const startServer = async () => {
  const server = app.listen(0);
  await new Promise((resolve) => server.once('listening', resolve));
  return server;
};

const authedFetch = (url, options = {}) => fetch(url, {
  ...options,
  headers: {
    ...(options.headers || {}),
    Cookie: `${SESSION_COOKIE}=${sessionToken}`,
  },
});

async function loadExcelJsWorkbook(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
}

function findExcelRow(sheet, label) {
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (sheet.getCell(rowNumber, 1).value === label) return sheet.getRow(rowNumber);
  }
  return null;
}

test('GET / returns dashboard HTML', async () => {
  const server = await startServer();
  const { port } = server.address();

  try {
    const response = await authedFetch(`http://127.0.0.1:${port}/`);
    assert.equal(response.status, 200);
    const text = await response.text();
    assert.match(text, /Gestion des dépenses familiales/i);
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});

test('GET /login renders without an error message', async () => {
  const server = await startServer();
  const { port } = server.address();

  try {
    const response = await fetch(`http://127.0.0.1:${port}/login?next=%2F`);
    assert.equal(response.status, 200);
    const html = await response.text();
    assert.match(html, /Connectez-vous pour accéder aux finances de votre foyer/);
    assert.match(html, /name="identifier"/);
    assert.match(html, /data-toggle-password/);
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});

test('login creates a persistent session and logout invalidates it and prevents private page access', async () => {
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const protectedResponse = await fetch(`${baseUrl}/income`, { redirect: 'manual' });
    assert.equal(protectedResponse.status, 302);
    assert.match(protectedResponse.headers.get('location'), /^\/login\?next=/);

    const login = await fetch(`${baseUrl}/login`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        identifier: 'test@example.com',
        password: 'test-password',
        next: '/income?month=2026-11',
      }),
      redirect: 'manual',
    });
    assert.equal(login.status, 302);
    assert.equal(login.headers.get('location'), '/income?month=2026-11');
    assert.match(login.headers.get('set-cookie'), /household_session=.*HttpOnly/);
    assert.match(login.headers.get('cache-control'), /no-store/);
    const sessionCookie = login.headers.get('set-cookie').split(';')[0];

    for (let index = 0; index < 2; index += 1) {
      const privatePage = await fetch(`${baseUrl}/income?month=2026-11`, {
        headers: { Cookie: sessionCookie },
      });
      assert.equal(privatePage.status, 200);
      assert.match(privatePage.headers.get('cache-control'), /no-store/);
    }

    const logout = await fetch(`${baseUrl}/logout`, {
      method: 'POST',
      headers: { Cookie: sessionCookie },
      redirect: 'manual',
    });
    assert.equal(logout.status, 302);
    assert.equal(logout.headers.get('location'), '/login?loggedOut=1');
    assert.match(logout.headers.get('set-cookie'), /Max-Age=0/);

    const afterLogout = await fetch(`${baseUrl}/income`, {
      headers: { Cookie: sessionCookie },
      redirect: 'manual',
    });
    assert.equal(afterLogout.status, 302);
    assert.match(afterLogout.headers.get('location'), /^\/login\?next=/);
    const loginPage = await fetch(`${baseUrl}/login?loggedOut=1`);
    assert.match(await loginPage.text(), /Vous êtes déconnecté\(e\)\./);
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});

test('GET / renders the selected dashboard month and trend period', async () => {
  const server = await startServer();
  const { port } = server.address();

  try {
    const response = await authedFetch(`http://127.0.0.1:${port}/?month=2026-11&period=3`);
    assert.equal(response.status, 200);
    const text = await response.text();
    assert.match(text, /name="month" value="2026-11"/);
    assert.match(text, /Comparaison mensuelle des revenus et dépenses sur 3 mois/);
    assert.match(text, /Revenus encaissés/);
    assert.match(text, /L’épargne pension père reste exclue des revenus disponibles/);
    assert.match(text, /Évolution du solde/);
    assert.match(text, /Revenus par source/);
    assert.match(text, /Dernières opérations/);
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});

test('sidebar links open pages and page actions expose modal triggers', async () => {
  const server = await startServer();
  const { port } = server.address();

  try {
    for (const route of [
      '/expenses', '/groceries', '/history', '/monthly', '/import', '/export', '/backup',
      '/categories', '/tenants', '/rents', '/income', '/savings', '/account',
    ]) {
      const response = await authedFetch(`http://127.0.0.1:${port}${route}`);
      assert.equal(response.status, 200, `${route} should render as a page`);
      const text = await response.text();
      assert.match(text, /class="sidebar"/);
      assert.match(text, /href="\/history"/);
      assert.match(text, /src="\/app.js"/);
    }

    const expenses = await authedFetch(`http://127.0.0.1:${port}/expenses`);
    const expensesHtml = await expenses.text();
    assert.match(expensesHtml, /data-modal-template="expense-create"/);
    assert.match(expensesHtml, /class="close-modal"[^>]*data-close-modal/);
    assert.match(expensesHtml, /class="sidebar-logout"/);
    assert.doesNotMatch(expensesHtml, /href="\/groceries" data-modal-link/);
    const dashboard = await authedFetch(`http://127.0.0.1:${port}/`);
    assert.match(await dashboard.text(), /class="close-modal"[^>]*data-close-modal/);
    const incomePage = await authedFetch(`http://127.0.0.1:${port}/income?month=2026-11`);
    const incomeHtml = await incomePage.text();
    assert.match(incomeHtml, /Revenus mensuels/);
    assert.match(incomeHtml, /class="income-source-card"/);
    assert.match(incomeHtml, /data-modal-template="income-create"/);
    assert.match(incomeHtml, /data-default-pension="180000"/);
    assert.doesNotMatch(incomeHtml, /Sources mensuelles principales/);
    const savingsPage = await authedFetch(`http://127.0.0.1:${port}/savings`);
    assert.match(await savingsPage.text(), /Épargne pension père/);
  } finally {
    server.close();
  }
});

test('administrator can change email and password after reauthentication', async () => {
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;
  const account = await db.prepare('SELECT id FROM users LIMIT 1').get();
  const previousSession = await createSession(account.id);
  const oldCookie = `${SESSION_COOKIE}=${previousSession}`;
  const originalPassword = 'test-password';

  try {
    const page = await authedFetch(`${baseUrl}/account`);
    assert.equal(page.status, 200);
    const html = await page.text();
    assert.match(html, /Compte administrateur/);
    assert.match(html, /name="new_password"/);
    assert.match(html, /name="email"/);

    const invalidEmail = await authedFetch(`${baseUrl}/account/email`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ email: 'adresse-invalide', current_password: originalPassword }),
    });
    assert.equal(invalidEmail.status, 400);
    assert.match(await invalidEmail.text(), /adresse e-mail valide/i);

    const emailUpdate = await authedFetch(`${baseUrl}/account/email`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        email: 'ADMIN-UPDATED@example.com',
        current_password: originalPassword,
      }),
      redirect: 'manual',
    });
    assert.equal(emailUpdate.status, 302);
    assert.equal(emailUpdate.headers.get('location'), '/account?updated=email');
    assert.equal((await db.prepare('SELECT email FROM users WHERE id = ?').get(account.id)).email, 'admin-updated@example.com');

    const wrongPassword = await authedFetch(`${baseUrl}/account/password`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        current_password: 'incorrect-password',
        new_password: 'nouveau-mot-de-passe-solide',
        confirm_password: 'nouveau-mot-de-passe-solide',
      }),
    });
    assert.equal(wrongPassword.status, 401);

    const mismatch = await authedFetch(`${baseUrl}/account/password`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        current_password: originalPassword,
        new_password: 'nouveau-mot-de-passe-solide',
        confirm_password: 'un-autre-mot-de-passe-solide',
      }),
    });
    assert.equal(mismatch.status, 400);
    assert.match(await mismatch.text(), /confirmation.*ne correspond pas/i);

    const passwordUpdate = await authedFetch(`${baseUrl}/account/password`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        current_password: originalPassword,
        new_password: 'nouveau-mot-de-passe-solide',
        confirm_password: 'nouveau-mot-de-passe-solide',
      }),
      redirect: 'manual',
    });
    assert.equal(passwordUpdate.status, 302);
    assert.equal(passwordUpdate.headers.get('location'), '/account?updated=password');
    const newCookie = passwordUpdate.headers.get('set-cookie').split(';')[0];
    assert.notEqual(newCookie, `${SESSION_COOKIE}=${sessionToken}`);

    const oldSessionResponse = await fetch(`${baseUrl}/account`, {
      headers: { Cookie: oldCookie },
      redirect: 'manual',
    });
    assert.equal(oldSessionResponse.status, 302);
    const newSessionPage = await fetch(`${baseUrl}/account`, { headers: { Cookie: newCookie } });
    assert.equal(newSessionPage.status, 200);

    const oldPasswordLogin = await fetch(`${baseUrl}/login`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        identifier: 'admin-updated@example.com',
        password: originalPassword,
      }),
    });
    assert.equal(oldPasswordLogin.status, 401);

    const newPasswordLogin = await fetch(`${baseUrl}/login`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        identifier: 'admin-updated@example.com',
        password: 'nouveau-mot-de-passe-solide',
      }),
      redirect: 'manual',
    });
    assert.equal(newPasswordLogin.status, 302);
    assert.match(newPasswordLogin.headers.get('set-cookie'), /household_session=.*HttpOnly/);
  } finally {
    await db.prepare('DELETE FROM user_sessions WHERE user_id = ?').run(account.id);
    await db.prepare(
      'UPDATE users SET email = ?, password_hash = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
    ).run('test@example.com', await hashPassword(originalPassword), account.id);
    sessionToken = await createSession(account.id);
    await new Promise((resolve) => server.close(resolve));
  }
});

test('monthly revenues, exceptional income, rent and protected savings stay separate', async () => {
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;
  const postForm = (route, values) => authedFetch(`${baseUrl}${route}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(values),
  });

  try {
    await postForm('/units', { floor: 'Rez-de-chaussée', label: 'TEST-1', housing_type: 'Studio' });
    const unit = await db.prepare("SELECT id FROM units WHERE label = 'TEST-1'").get();
    assert.ok(unit);
    await postForm('/tenants', {
      unit_id: String(unit.id),
      first_name: 'Test',
      last_name: 'Locataire',
      people_count: '2',
      monthly_rent: '100000',
      entry_date: '2026-01-01',
      phone: '',
    });
    await postForm('/rents/generate', { month_key: '2026-11' });
    await postForm('/rents/generate', { month_key: '2026-11' });
    const obligation = await db.prepare(`
      SELECT o.id FROM rent_obligations o JOIN tenants t ON t.id = o.tenant_id
      WHERE t.first_name = 'Test' AND o.month_key = '2026-11'
    `).get();
    assert.ok(obligation);
    for (const amount of ['40000', '60000']) {
      const payment = await postForm(`/rents/${obligation.id}/payments`, {
        amount,
        payment_date: '2026-11-15',
        payment_method: 'Virement',
        comment: '',
      });
      assert.equal(payment.status, 200);
    }

    await postForm('/income/monthly', {
      month_key: '2026-11',
      pape_amount: '300000',
      pape_received_date: '2026-11-03',
      pension_amount: '180000',
      pension_received_date: '2026-11-05',
      comment: 'Revenus initiaux du test',
    });
    await postForm('/income/monthly', {
      month_key: '2026-11',
      pape_amount: '350000',
      pape_received_date: '2026-11-20',
      pension_amount: '180000',
      pension_received_date: '2026-11-05',
      comment: 'Montant corrigé',
    });

    const incomeCategory = await db.prepare("SELECT id FROM income_categories WHERE name = 'Autres recettes'").get();
    await postForm('/income', {
      income_date: '2026-11-22',
      month_key: '2026-11',
      title: 'Prime exceptionnelle',
      category_id: String(incomeCategory.id),
      amount: '25000',
      source: 'Prime exceptionnelle',
      payment_method: 'Virement',
      comment: '',
    });
    const exceptionalIncome = await db.prepare(`
      SELECT id FROM incomes WHERE month_key = '2026-11' AND source = 'Prime exceptionnelle'
    `).get();
    await postForm(`/income/${exceptionalIncome.id}/update`, {
      income_date: '2026-11-22',
      month_key: '2026-11',
      title: 'Prime exceptionnelle ajustée',
      category_id: String(incomeCategory.id),
      amount: '30000',
      source: 'Prime exceptionnelle',
      payment_method: 'Virement',
      comment: 'Montant corrigé',
    });

    await postForm('/savings/accounts', { name: 'Épargne test' });
    const account = await db.prepare("SELECT id FROM savings_accounts WHERE name = 'Épargne test'").get();
    await postForm('/savings/transactions', {
      account_id: String(account.id),
      transaction_date: '2026-11-10',
      transaction_type: 'deposit',
      amount: '10000',
      comment: 'Versement de test',
    });
    const protectedAccount = await db.prepare(`
      SELECT id FROM savings_accounts WHERE is_protected = TRUE
    `).get();
    assert.ok(protectedAccount);
    await db.prepare(`
      INSERT INTO savings_balance_snapshots (account_id, month_key, amount)
      VALUES (?, '2026-10', 500000)
    `).run(protectedAccount.id);
    const unconfirmedWithdrawal = await postForm('/savings/transactions', {
      account_id: String(protectedAccount.id),
      transaction_date: '2026-11-12',
      transaction_type: 'withdrawal',
      amount: '20000',
      comment: 'Retrait volontaire',
    });
    assert.equal(unconfirmedWithdrawal.status, 400);
    await postForm('/savings/transactions', {
      account_id: String(protectedAccount.id),
      transaction_date: '2026-11-12',
      transaction_type: 'withdrawal',
      amount: '20000',
      comment: 'Retrait volontaire confirmé',
      protected_confirmation: '1',
    });

    const financeResponse = await authedFetch(`${baseUrl}/finance/monthly-data?month=2026-11`);
    const finance = await financeResponse.json();
    assert.equal(finance.rentExpected, 100000);
    assert.equal(finance.rentReceived, 100000);
    assert.equal(finance.rentOutstanding, 0);
    assert.equal(finance.ordinaryIncome, 560000);
    assert.equal(finance.papeAmount, 350000);
    assert.equal(finance.pensionAmount, 180000);
    assert.equal(finance.resourcesAvailable, 660000);
    assert.equal(finance.savingsBalance, 490000);
    assert.equal(finance.protectedSavingsBalance, 480000);
    assert.equal(finance.closingBalance, 660000);
    const incomePage = await authedFetch(`${baseUrl}/income?month=2026-11`);
    const incomeHtml = await incomePage.text();
    assert.match(incomeHtml, /TOTAL DES REVENUS ENCAISSÉS/);
    assert.match(incomeHtml, /Report du mois précédent/);
    assert.match(incomeHtml, /Total des ressources disponibles/);
    assert.match(incomeHtml, /350[ \u00a0\u202f]000/);
    assert.match(incomeHtml, /Test Locataire/);
    assert.match(incomeHtml, /Détail des paiements reçus/);
    assert.match(incomeHtml, /Prime exceptionnelle ajustée/);
    assert.match(incomeHtml, /data-income-edit-source="Pape"/);
    const expensesHtml = await (await authedFetch(`${baseUrl}/expenses?month=2026-11`)).text();
    assert.match(expensesHtml, /Report du mois précédent/);
    assert.match(expensesHtml, /Total des ressources disponibles/);
    const monthlyHtml = await (await authedFetch(`${baseUrl}/monthly?year=2026&month=11`)).text();
    assert.match(monthlyHtml, /Report du mois précédent/);
    assert.match(monthlyHtml, /Total des ressources disponibles/);
    const dashboardHtml = await (await authedFetch(`${baseUrl}/?month=2026-11`)).text();
    assert.match(dashboardHtml, /Report précédent :/);
    assert.equal(Number((await db.prepare(`
      SELECT COUNT(*) AS count FROM incomes WHERE month_key = '2026-11' AND source = 'Prime exceptionnelle'
    `).get()).count), 1);
    const decemberIncomePage = await authedFetch(`${baseUrl}/income?month=2026-12`);
    assert.match(await decemberIncomePage.text(), /name="month" value="2026-12"/);
    const monthlyExport = await authedFetch(`${baseUrl}/export/month/2026/11`);
    const exportedMonth = XLSX.read(Buffer.from(await monthlyExport.arrayBuffer()), { type: 'buffer' });
    assert.equal(exportedMonth.Sheets.Novembre.B15.v, 180000);
    assert.equal(exportedMonth.Sheets.Novembre.B17.v, 350000);
    assert.equal(exportedMonth.Sheets.Novembre.B21.f, 'B15+B17+F13+SUM(B18:B18)');
    assert.equal(exportedMonth.Sheets.Novembre.B21.v, 660000);
    assert.equal(exportedMonth.Sheets.Novembre.B24.f, 'B21');
    assert.equal(exportedMonth.Sheets.Novembre.B22.v, finance.incomeTotal - finance.totalExpenses);
    assert.equal(exportedMonth.Sheets.Novembre.A23.v, 'Restant du mois dernier');
    assert.equal(exportedMonth.Sheets.Novembre.B23.v, finance.openingBalance);
    assert.equal(exportedMonth.Sheets.Novembre.A24.v, 'Total des ressources disponibles');
    assert.equal(exportedMonth.Sheets.Novembre.B24.v, finance.incomeTotal);
    assert.equal(exportedMonth.Sheets.Novembre.B22.f, 'B24-E42');
    const decemberFinanceResponse = await authedFetch(`${baseUrl}/finance/monthly-data?month=2026-12`);
    const decemberFinance = await decemberFinanceResponse.json();
    assert.equal(decemberFinance.openingBalance, finance.closingBalance);
    const decemberExport = await authedFetch(
      `${baseUrl}/export/month/2026/12?includeCarryover=0&includeCarryover=1`,
    );
    const decemberBuffer = Buffer.from(await decemberExport.arrayBuffer());
    const exportedDecember = XLSX.read(decemberBuffer, { type: 'buffer' });
    assert.equal(exportedDecember.Sheets.Décembre.B23.v, finance.closingBalance);
    assert.equal(exportedDecember.Sheets.Décembre.B24.v, decemberFinance.resourcesAvailable);
    assert.equal(exportedDecember.Sheets.Décembre.B22.v, decemberFinance.closingBalance);
    const visibleDecember = await loadExcelJsWorkbook(decemberBuffer);
    const visibleReportRow = findExcelRow(
      visibleDecember.getWorksheet('Décembre'),
      'Restant du mois dernier',
    );
    assert.ok(visibleReportRow);
    assert.equal(visibleReportRow.hidden, false);
    assert.equal(visibleDecember.getWorksheet('Décembre').getCell('B24').value.result, decemberFinance.resourcesAvailable);
    assert.equal(visibleDecember.getWorksheet('Décembre').getCell('B22').value.result, decemberFinance.closingBalance);

    const decemberOverrideBefore = await db.prepare(
      'SELECT opening_override FROM monthly_balances WHERE month_key = ?',
    ).get('2026-12');
    const uncheckedDecemberResponse = await authedFetch(
      `${baseUrl}/export/month/2026/12?includeCarryover=0`,
    );
    const uncheckedDecember = await loadExcelJsWorkbook(
      Buffer.from(await uncheckedDecemberResponse.arrayBuffer()),
    );
    const uncheckedDecemberSheet = uncheckedDecember.getWorksheet('Décembre');
    const uncheckedReportRow = findExcelRow(uncheckedDecemberSheet, 'Restant du mois dernier');
    assert.ok(uncheckedReportRow);
    assert.equal(uncheckedReportRow.hidden, false);
    assert.equal(uncheckedReportRow.getCell(2).value, decemberFinance.openingBalance);
    assert.equal(
      Number(uncheckedDecemberSheet.getCell('B24').value.result || 0),
      decemberFinance.incomeTotal,
    );
    assert.equal(
      Number(uncheckedDecemberSheet.getCell('B22').value.result || 0),
      decemberFinance.incomeTotal - decemberFinance.totalExpenses,
    );
    const decemberFinanceAfterExport = await (
      await authedFetch(`${baseUrl}/finance/monthly-data?month=2026-12`)
    ).json();
    const decemberOverrideAfter = await db.prepare(
      'SELECT opening_override FROM monthly_balances WHERE month_key = ?',
    ).get('2026-12');
    assert.deepEqual(decemberFinanceAfterExport, decemberFinance);
    assert.deepEqual(decemberOverrideAfter, decemberOverrideBefore);
    const balanceCorrection = await postForm('/balances/adjustment', {
      month_key: '2026-11',
      opening_balance: String(finance.openingBalance + 10000),
      reason: 'Correction test avant export',
    });
    assert.equal(balanceCorrection.status, 200);
    const correctedNovember = await (await authedFetch(`${baseUrl}/finance/monthly-data?month=2026-11`)).json();
    const correctedDecember = await (await authedFetch(`${baseUrl}/finance/monthly-data?month=2026-12`)).json();
    assert.equal(correctedDecember.openingBalance, correctedNovember.closingBalance);
    const januaryOpeningCorrection = await postForm('/balances/adjustment', {
      month_key: '2027-01',
      opening_balance: '80000',
      reason: 'Solde initial reporté',
    });
    assert.equal(januaryOpeningCorrection.status, 200);
    const januaryFinance = await (await authedFetch(`${baseUrl}/finance/monthly-data?month=2027-01`)).json();
    assert.equal(januaryFinance.openingBalance, 80000);
    const incomeCsv = await authedFetch(`${baseUrl}/export/csv/2026/incomes`);
    const incomeCsvText = await incomeCsv.text();
    assert.match(incomeCsvText, /350000/);
    assert.match(incomeCsvText, /180000/);
    assert.match(incomeCsvText, /30000/);
    assert.doesNotMatch(incomeCsvText, /300000/);
    const revenueAudit = await db.prepare(`
      SELECT COUNT(*) AS count FROM audit_log
      WHERE entity_type = 'monthly_revenue' AND entity_id = '2026-11'
    `).get();
    assert.equal(Number(revenueAudit.count), 2);

    const historyResponse = await authedFetch(`${baseUrl}/history?month=2026-11`);
    const historyHtml = await historyResponse.text();
    assert.equal(historyResponse.status, 200);
    assert.match(historyHtml, /Loyer encaissé/);
    assert.match(historyHtml, /Recette/);
    assert.match(historyHtml, /Épargne/);
  } finally {
    server.close();
  }
});

test('history filters financial operations and paginates large result sets', async () => {
  const category = await db.prepare('SELECT id FROM categories ORDER BY id LIMIT 1').get();
  await db.prepare(`
    INSERT INTO expenses (title, category_id, amount, expense_date, month_key)
    SELECT 'Historique pagination ' || number, ?, 100, '2026-11-13', '2026-11'
    FROM generate_series(1, 51) AS number
  `).run(category.id);
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const firstPage = await authedFetch(`${baseUrl}/history?type=expense&month=2026-11&search=pagination`);
    const firstHtml = await firstPage.text();
    assert.equal(firstPage.status, 200);
    assert.match(firstHtml, /51 opérations retrouvées/);
    assert.match(firstHtml, /page 1 sur 2/);
    const nextHref = firstHtml.match(/href="([^"]*page=2[^"]*)"/)?.[1]?.replaceAll('&amp;', '&');
    assert.ok(nextHref);
    assert.match(nextHref, /type=expense/);
    assert.match(nextHref, /month=2026-11/);
    assert.match(nextHref, /search=pagination/);

    const secondPage = await authedFetch(new URL(nextHref, baseUrl));
    assert.match(await secondPage.text(), /page 2 sur 2/);
  } finally {
    await db.prepare("DELETE FROM expenses WHERE title LIKE 'Historique pagination %'").run();
    server.close();
  }
});

test('POST /expenses stores an ordinary expense', async () => {
  const server = await startServer();
  const { port } = server.address();

  try {
    const form = new FormData();
    Object.entries({
      title: 'Electricité',
      category_id: '1',
      amount: '25000',
      expense_date: '2026-11-13',
      month_key: '2026-11',
      beneficiary: 'Fournisseur test',
      paid_by: 'Administrateur',
      payment_method: 'Virement',
      description: 'Test',
      reference: 'facture-1',
    }).forEach(([key, value]) => form.append(key, value));
    const receiptBytes = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/lvsAAAAASUVORK5CYII=', 'base64');
    form.append('receipt', new Blob([receiptBytes], { type: 'image/png' }), 'preuve.png');
    const response = await authedFetch(`http://127.0.0.1:${port}/expenses`, {
      method: 'POST',
      body: form,
    });

    assert.equal(response.status, 200);
    const expense = await db.prepare("SELECT * FROM expenses WHERE title = 'Electricité'").get();
    assert.equal(Number(expense.amount), 25000);
    assert.equal(expense.beneficiary, 'Fournisseur test');
    assert.equal(expense.month_key, '2026-11');
    assert.ok(expense.created_by);
    const attachment = await db.prepare('SELECT * FROM expense_attachments WHERE expense_id = ?').get(expense.id);
    assert.equal(attachment.file_name, 'preuve.png');
    const receiptResponse = await authedFetch(`http://127.0.0.1:${port}/expenses/${expense.id}/receipt`);
    assert.equal(receiptResponse.status, 200);
    assert.equal(receiptResponse.headers.get('content-type'), 'image/png');
    assert.deepEqual(Buffer.from(await receiptResponse.arrayBuffer()), receiptBytes);
  } finally {
    server.close();
  }
});

test('expense can be edited, filtered and deleted with an audit trail', async () => {
  const expense = await db.prepare("SELECT * FROM expenses WHERE title = 'Electricité'").get();
  const attachment = await db.prepare('SELECT storage_path FROM expense_attachments WHERE expense_id = ?').get(expense.id);
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;
  const fields = {
    title: 'Facture électrique corrigée',
    category_id: String(expense.category_id),
    amount: '27500',
    expense_date: '2026-11-14',
    month_key: '2026-11',
    beneficiary: 'Fournisseur corrigé',
    paid_by: 'Administrateur',
    payment_method: 'Virement',
    description: 'Montant corrigé',
    reference: 'facture-1-corrigée',
  };

  try {
    const updateResponse = await authedFetch(`${baseUrl}/expenses/${expense.id}/update`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams(fields),
    });
    assert.equal(updateResponse.status, 200);
    const filtered = await authedFetch(`${baseUrl}/expenses?month=2026-11&category=${expense.category_id}&search=corrigée`);
    assert.match(await filtered.text(), /Facture électrique corrigée/);
    assert.equal(Number((await db.prepare('SELECT amount FROM expenses WHERE id = ?').get(expense.id)).amount), 27500);

    const deleteResponse = await authedFetch(`${baseUrl}/expenses/${expense.id}/delete`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ return_to: '/expenses' }),
    });
    assert.equal(deleteResponse.status, 200);
    assert.equal(await db.prepare('SELECT id FROM expenses WHERE id = ?').get(expense.id), undefined);
    assert.equal(fs.existsSync(path.join(testDataDir, 'data', 'receipts', attachment.storage_path)), false);
    assert.equal(Number((await db.prepare(`
      SELECT COUNT(*) AS count FROM audit_log WHERE entity_type = 'expense' AND entity_id = ?
    `).get(String(expense.id))).count), 3);
  } finally {
    server.close();
  }
});

test('GET /export/2026 returns the twelve reference sheets and annual summary', async () => {
  const server = await startServer();
  const { port } = server.address();

  try {
    const response = await authedFetch(
      `http://127.0.0.1:${port}/export/2026?includeCarryover=0&includeCarryover=1`,
    );
    assert.equal(response.status, 200);
    assert.match(response.headers.get('content-disposition'), /DEPENSES_MAISON_2026\.xlsx/);
    const workbook = XLSX.read(Buffer.from(await response.arrayBuffer()), { type: 'buffer' });
    const exportPage = await authedFetch(`http://127.0.0.1:${port}/export`);
    const exportHtml = await exportPage.text();
    assert.match(exportHtml, /name="includeCarryover" value="1"/);
    assert.doesNotMatch(exportHtml, /name="includeCarryover" value="1" checked/);
    assert.match(exportHtml, /Inclure le report du mois dernier dans le calcul du restant final/);
    assert.deepEqual(workbook.SheetNames, [
      'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin',
      'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre',
      'Récapitulatif annuel',
    ]);
    const exportedFinances = new Map();
    for (const month of workbook.SheetNames.slice(0, 12)) {
      assert.deepEqual(
        ['A2', 'B2', 'D2', 'E2', 'F2', 'G2'].map((address) => workbook.Sheets[month][address]?.v),
        ['NIVEAU', 'NOM ET PRÉNOM', 'TYPE DE LOGEMENT', 'NOMBRE DE PERSONNES', 'MONTANT À PAYER', 'STATUT DU PAIEMENT'],
      );
      assert.equal(workbook.Sheets[month].A15.v, 'Pension père');
      assert.equal(typeof workbook.Sheets[month].B15.v, 'number');
      assert.equal(workbook.Sheets[month].A17.v, 'Pape');
      assert.equal(typeof workbook.Sheets[month].B17.v, 'number');
      assert.equal(workbook.Sheets[month].A23.v, 'Restant du mois dernier');
      assert.equal(typeof workbook.Sheets[month].B23.v, 'number');
      assert.equal(workbook.Sheets[month].A30.v.trim(), 'Epargne pension père');
      assert.equal(typeof workbook.Sheets[month].B30.v, 'number');
      assert.doesNotMatch(workbook.Sheets[month].B21.f, /B30/);
    }
    let previousClosingBalance = null;
    for (let index = 0; index < 12; index += 1) {
      const monthKey = `2026-${String(index + 1).padStart(2, '0')}`;
      const monthName = workbook.SheetNames[index];
      const finance = await (await authedFetch(
        `http://127.0.0.1:${port}/finance/monthly-data?month=${monthKey}`,
      )).json();
      exportedFinances.set(monthKey, finance);
      assert.equal(workbook.Sheets[monthName].B15.v, finance.pensionAmount);
      assert.equal(workbook.Sheets[monthName].B17.v, finance.papeAmount);
      assert.equal(workbook.Sheets[monthName].B23.v, finance.openingBalance);
      assert.equal(Number(workbook.Sheets[monthName].B24.v || 0), finance.resourcesAvailable);
      assert.equal(Number(workbook.Sheets[monthName].B22.v || 0), finance.closingBalance);
      assert.equal(workbook.Sheets[monthName].B30.v, finance.protectedSavingsBalance);
      const monthBalance = await db.prepare(
        'SELECT opening_override FROM monthly_balances WHERE month_key = ?',
      ).get(monthKey);
      if (previousClosingBalance !== null && monthBalance?.opening_override == null) {
        assert.equal(finance.openingBalance, previousClosingBalance);
      }
      previousClosingBalance = finance.closingBalance;
    }
    assert.equal(workbook.Sheets.Décembre.B23.v, workbook.Sheets.Novembre.B22.v);
    assert.match(workbook.Sheets.Janvier.A1.v, /Loyers du mois de Décembre 2025/);
    assert.match(workbook.Sheets.Janvier.D20.v, /Dépenses du mois de Janvier/);
    assert.equal(workbook.Sheets.Janvier.E42.f, 'SUM(E22:E41)');
    assert.equal(workbook.Sheets['Récapitulatif annuel'].A1.v, 'Récapitulatif annuel 2026');
    assert.ok(workbook.Sheets.Janvier['!merges'].some((range) => range.s.r === 0 && range.s.c === 0 && range.e.c === 6));
    assert.ok(workbook.Sheets.Janvier['!merges'].some((range) => range.s.r === 19 && range.s.c === 3 && range.e.c === 4));
    const uncheckedAnnualResponse = await authedFetch(
      `http://127.0.0.1:${port}/export/2026?includeCarryover=0`,
    );
    const uncheckedAnnual = await loadExcelJsWorkbook(
      Buffer.from(await uncheckedAnnualResponse.arrayBuffer()),
    );
    for (const monthName of workbook.SheetNames.slice(0, 12)) {
      const sheet = uncheckedAnnual.getWorksheet(monthName);
      const reportRow = findExcelRow(sheet, 'Restant du mois dernier');
      const resourcesRow = findExcelRow(sheet, 'Total des ressources disponibles');
      assert.ok(reportRow);
      assert.ok(resourcesRow);
      assert.equal(reportRow.hidden, false);
      assert.equal(resourcesRow.hidden, true);
      const monthKey = `2026-${String(workbook.SheetNames.indexOf(monthName) + 1).padStart(2, '0')}`;
      const finance = exportedFinances.get(monthKey);
      assert.equal(reportRow.getCell(2).value, finance.openingBalance);
      assert.equal(
        Number(resourcesRow.getCell(2).value.result || 0),
        finance.incomeTotal,
      );
      const closingRow = sheet.getRow(resourcesRow.number - 2);
      assert.equal(
        Number(closingRow.getCell(2).value.result || 0),
        finance.incomeTotal - finance.totalExpenses,
      );
    }
  } finally {
    server.close();
  }
});

test('monthly exports provide XLSX, CSV and PDF formats', async () => {
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const monthlyPage = await authedFetch(`${baseUrl}/monthly?year=2026&month=11`);
    assert.equal(monthlyPage.status, 200);
    const monthlyHtml = await monthlyPage.text();
    assert.match(monthlyHtml, /\/export\/month\/2026\/11/);
    assert.match(monthlyHtml, /name="includeCarryover" value="1"/);
    assert.doesNotMatch(monthlyHtml, /name="includeCarryover" value="1" checked/);
    const exportHtml = await (await authedFetch(`${baseUrl}/export`)).text();
    assert.match(exportHtml, /\/export\/csv\/\d+\/expenses/);
    assert.match(exportHtml, /name="includeCarryover" value="1"/);
    assert.doesNotMatch(exportHtml, /name="includeCarryover" value="1" checked/);

    const workbookResponse = await authedFetch(`${baseUrl}/export/month/2026/11`);
    assert.equal(workbookResponse.status, 200);
    const workbook = XLSX.read(Buffer.from(await workbookResponse.arrayBuffer()), { type: 'buffer' });
    assert.deepEqual(workbook.SheetNames, ['Novembre']);
    assert.match(workbook.Sheets.Novembre.A1.v, /Loyers du mois d'Octobre/);
    assert.equal(workbook.Sheets.Novembre.E42.f, 'SUM(E22:E41)');
    assert.equal(workbook.Sheets.Novembre.B24.f, 'B21');

    const multiMonthResponse = await authedFetch(`${baseUrl}/export/months?year=2026&months=1&months=2`);
    assert.equal(multiMonthResponse.status, 200);
    const multiMonthWorkbook = XLSX.read(Buffer.from(await multiMonthResponse.arrayBuffer()), { type: 'buffer' });
    assert.deepEqual(multiMonthWorkbook.SheetNames, ['Janvier', 'Février']);
    const uncheckedMultiMonthResponse = await authedFetch(
      `${baseUrl}/export/months?year=2026&months=1&months=2&includeCarryover=0`,
    );
    const uncheckedMultiMonth = await loadExcelJsWorkbook(
      Buffer.from(await uncheckedMultiMonthResponse.arrayBuffer()),
    );
    const checkedMultiMonthResponse = await authedFetch(
      `${baseUrl}/export/months?year=2026&months=1&months=2&includeCarryover=0&includeCarryover=1`,
    );
    const checkedMultiMonth = await loadExcelJsWorkbook(
      Buffer.from(await checkedMultiMonthResponse.arrayBuffer()),
    );
    for (const monthName of ['Janvier', 'Février']) {
      const sheet = uncheckedMultiMonth.getWorksheet(monthName);
      const reportRow = findExcelRow(sheet, 'Restant du mois dernier');
      assert.ok(reportRow);
      assert.equal(reportRow.hidden, false);
      assert.equal(findExcelRow(sheet, 'Total des ressources disponibles').hidden, true);
      assert.equal(typeof reportRow.getCell(2).value, 'number');
      assert.equal(sheet.getCell('B24').value.formula, 'B21');
      const checkedSheet = checkedMultiMonth.getWorksheet(monthName);
      assert.equal(findExcelRow(checkedSheet, 'Restant du mois dernier').hidden, false);
      assert.equal(findExcelRow(checkedSheet, 'Total des ressources disponibles').hidden, false);
      assert.equal(checkedSheet.getCell('B24').value.formula, 'B21+B23');
    }

    const csvResponse = await authedFetch(`${baseUrl}/export/csv/2026/expenses`);
    assert.equal(csvResponse.status, 200);
    assert.match(await csvResponse.text(), /^"Date";"Mois";"Désignation";"Catégorie";"Montant"/);

    const pdfResponse = await authedFetch(`${baseUrl}/export/pdf/month/2026/11`);
    assert.equal(pdfResponse.status, 200);
    assert.match(pdfResponse.headers.get('content-type'), /application\/pdf/);
    assert.equal((await pdfResponse.arrayBuffer()).byteLength > 100, true);
  } finally {
    server.close();
  }
});

test('historical import backs up before one-time cleanup, filters dates and flags reimport duplicates', async () => {
  const categoryId = (await db.prepare('SELECT id FROM categories ORDER BY id LIMIT 1').get()).id;
  await db.prepare(`
    INSERT INTO expenses (title, category_id, amount, expense_date)
    VALUES ('Ancienne dépense de test', ?, 100, '2026-01-02')
  `).run(categoryId);
  await db.prepare(`
    INSERT INTO grocery_items (item_name, purchase_date, quantity, unit, unit_price, total_amount)
    VALUES ('Ancien produit de test', '2026-01-03', 1, 'pièce', 200, 200)
  `).run();
  await db.prepare(`
    INSERT INTO import_batches (file_name, rows_json, status)
    VALUES ('old-test.xlsx', '[]', 'imported')
  `).run();
  await db.prepare(`
    INSERT INTO settings (key, value) VALUES (?, ?)
    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value
  `).run('setting_to_preserve', 'yes');

  const source = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(source, XLSX.utils.aoa_to_sheet([
    ['Date', 'Intitulé', 'Catégorie', 'Montant'],
    ['2026-01-06', 'Dépense catégorie nouvelle', 'Catégorie inconnue', 2000],
    ['2025-09-06', 'Dépense antérieure interdite', 'Eau', 700],
  ]), 'Janvier');
  XLSX.utils.book_append_sheet(source, XLSX.utils.aoa_to_sheet([
    ['Date', 'Marchandise', 'Quantité', 'Unité', 'Prix unitaire', 'Montant'],
    ['2026-01-07', 'Riz', 2, 'sac', 2500, 5000],
    ['2026-01-07', 'COMPTE BOUTIQUE', null, null, null, 5000],
  ]), 'Compte boutique');
  const fileBuffer = XLSX.write(source, { type: 'buffer', bookType: 'xlsx' });
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const uploadForm = new FormData();
    uploadForm.append('file', new Blob([fileBuffer]), 'history.xlsx');
    const uploadResponse = await authedFetch(`${baseUrl}/import`, { method: 'POST', body: uploadForm, redirect: 'manual' });
    assert.equal(uploadResponse.status, 302);
    const previewPath = uploadResponse.headers.get('location');
    const previewResponse = await authedFetch(`${baseUrl}${previewPath}`);
    const previewHtml = await previewResponse.text();
    assert.equal(previewResponse.status, 200);
    assert.match(previewHtml, /Catégories classées dans Autres/);
    assert.match(previewHtml, /Dépense antérieure interdite/);
    assert.match(previewHtml, /Sauvegarder et remplacer l’historique/);

    const importResponse = await authedFetch(`${baseUrl}/import/confirm/${previewPath.split('/').at(-1)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ reinitialize: '1' }),
      redirect: 'manual',
    });
    assert.equal(importResponse.status, 302);
    const reportResponse = await authedFetch(`${baseUrl}/import`);
    assert.match(await reportResponse.text(), /Rapport du dernier import/);
    assert.equal(Number((await db.prepare('SELECT COUNT(*) AS count FROM expenses').get()).count), 1);
    assert.equal(Number((await db.prepare('SELECT COUNT(*) AS count FROM grocery_items').get()).count), 1);
    assert.equal(Number((await db.prepare('SELECT COUNT(*) AS count FROM import_batches').get()).count), 1);
    assert.equal((await db.prepare("SELECT value FROM settings WHERE key = 'setting_to_preserve'").get()).value, 'yes');
    assert.equal(Number((await db.prepare("SELECT COUNT(*) AS count FROM expenses WHERE title = 'Dépense antérieure interdite'").get()).count), 0);
    assert.equal((await db.prepare("SELECT name FROM categories WHERE id = (SELECT category_id FROM expenses LIMIT 1)").get()).name, 'Autres dépenses');
    assert.equal((await db.prepare("SELECT value FROM settings WHERE key = 'historical_2026_reinitialized'").get()).value.length > 0, true);

    const backupRecord = await db.prepare("SELECT file_name, metadata FROM backup_runs WHERE source = 'pre-import-cleanup'").get();
    assert.ok(backupRecord);
    const backupMetadata = JSON.parse(backupRecord.metadata);
    assert.equal(backupMetadata.archiveVerified, true);
    assert.equal(backupMetadata.rowCounts.expenses, 1);
    const backupPath = path.join(testDataDir, 'data', 'backups', backupRecord.file_name);
    const archiveListing = execFileSync('pg_restore', ['--list', backupPath], { encoding: 'utf8' });
    assert.match(archiveListing, /TABLE DATA/);

    const household = XLSX.utils.book_new();
    const householdMonth = (heading) => [
      ['Gestion financière du foyer'],
      ['Niveau', 'Prénom et nom', 'Catégorie de logement', 'Nombre de personnes', 'Montant à payer', 'Statut de paiement'],
      ['Rez-de-chaussée', 'Irene Bassene', 'Un magasin', 1, 15_000, 'Payé'],
      ['', 'Total', '', '', 15_000],
      ['Pension père', 180_000],
      ['Pape', 300_000],
      ['', '', '', heading],
      ['', '', '', 'Type de dépenses', 'Montant'],
      ['', '', '', 'Facture WiFi', 15_000],
      ['Épargne pension père', 364_000, '', 'Pain', 2_500],
      ['', '', '', 'Compte boutique', 20_000],
      ['', '', '', 'Total', 37_500],
    ];
    XLSX.utils.book_append_sheet(household, XLSX.utils.aoa_to_sheet(householdMonth(
      'Dépenses du mois de Janvier',
    )), 'Janvier');
    XLSX.utils.book_append_sheet(household, XLSX.utils.aoa_to_sheet(householdMonth(
      'Dépenses du mois de Février',
    )), 'Février');
    const householdForm = new FormData();
    householdForm.append('file', new Blob([
      XLSX.write(household, { type: 'buffer', bookType: 'xlsx' }),
    ]), 'household.xlsx');
    const householdUpload = await authedFetch(`${baseUrl}/import`, {
      method: 'POST',
      body: householdForm,
      redirect: 'manual',
    });
    assert.equal(householdUpload.status, 302);
    const householdPreviewPath = householdUpload.headers.get('location');
    const householdPreview = await authedFetch(`${baseUrl}${householdPreviewPath}`);
    const householdPreviewHtml = await householdPreview.text();
    assert.match(householdPreviewHtml, /Locataires détectés/);
    assert.match(householdPreviewHtml, /Statuts de loyers détectés/);
    assert.match(householdPreviewHtml, /Recettes détectées/);
    assert.match(householdPreviewHtml, /Soldes d’épargne observés/);

    const householdImport = await authedFetch(
      `${baseUrl}/import/confirm/${householdPreviewPath.split('/').at(-1)}`,
      { method: 'POST', redirect: 'manual' },
    );
    assert.equal(householdImport.status, 302);
    assert.equal(Number((await db.prepare(`
      SELECT COUNT(*) AS count FROM tenants WHERE first_name = 'Irene' AND last_name = 'Bassene'
    `).get()).count), 1);
    assert.equal(Number((await db.prepare(`
      SELECT COUNT(*) AS count FROM rent_obligations o JOIN tenants t ON t.id = o.tenant_id
      WHERE t.first_name = 'Irene' AND t.last_name = 'Bassene'
    `).get()).count), 2);
    assert.equal(Number((await db.prepare(`
      SELECT COUNT(*) AS count FROM rent_payments p
      JOIN rent_obligations o ON o.id = p.obligation_id
      JOIN tenants t ON t.id = o.tenant_id
      WHERE p.payment_date IS NULL AND t.first_name = 'Irene' AND t.last_name = 'Bassene'
    `).get()).count), 2);
    assert.equal(Number((await db.prepare(`
      SELECT COUNT(*) AS count FROM incomes WHERE income_date IS NULL AND comment LIKE 'Importée de%'
    `).get()).count), 4);
    assert.equal(Number((await db.prepare(`
      SELECT COUNT(*) AS count FROM expenses
      WHERE expense_date IS NULL AND title IN ('Facture WiFi', 'Pain')
    `).get()).count), 4);
    assert.equal(Number((await db.prepare(`
      SELECT COUNT(*) AS count FROM grocery_items
      WHERE purchase_date IS NULL AND item_name = 'Compte boutique'
    `).get()).count), 2);
    assert.equal(Number((await db.prepare(`
      SELECT COUNT(*) AS count FROM savings_balance_snapshots s
      JOIN savings_accounts a ON a.id = s.account_id
      WHERE a.name = 'Épargne pension père' AND s.amount = 364000
    `).get()).count), 2);
    const savingsPage = await authedFetch(`${baseUrl}/savings`);
    const savingsHtml = await savingsPage.text();
    assert.match(savingsHtml, /Soldes mensuels observés/);
    assert.match(savingsHtml, /364/);
    const monthlyFinanceResponse = await authedFetch(`${baseUrl}/finance/monthly-data?month=2026-01`);
    assert.equal(Number((await monthlyFinanceResponse.json()).savingsBalance), 364_000);
    const replay = await authedFetch(
      `${baseUrl}/import/confirm/${householdPreviewPath.split('/').at(-1)}`,
      { method: 'POST', redirect: 'manual' },
    );
    assert.equal(replay.status, 409);

    const unknownSheet = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(unknownSheet, XLSX.utils.aoa_to_sheet([
      ['Operation historique à classer'],
    ]), 'Archive');
    const unknownForm = new FormData();
    unknownForm.append('file', new Blob([
      XLSX.write(unknownSheet, { type: 'buffer', bookType: 'xlsx' }),
    ]), 'unknown.xlsx');
    const unknownUpload = await authedFetch(`${baseUrl}/import`, {
      method: 'POST',
      body: unknownForm,
      redirect: 'manual',
    });
    const unknownPreviewPath = unknownUpload.headers.get('location');
    const unknownPreview = await authedFetch(`${baseUrl}${unknownPreviewPath}`);
    assert.match(await unknownPreview.text(), /Lignes non interprétées/);
    const correction = await authedFetch(
      `${baseUrl}${unknownPreviewPath}/correct`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          sheetName: 'Archive',
          rowNumber: '1',
          kind: 'expense',
          title: 'Operation historique à classer',
          categoryName: 'Autres dépenses',
          monthKey: '2026-03',
          amount: '123',
        }),
        redirect: 'manual',
      },
    );
    assert.equal(correction.status, 302);
    const correctedPreview = await authedFetch(`${baseUrl}${unknownPreviewPath}`);
    assert.doesNotMatch(await correctedPreview.text(), /Lignes non interprétées/);
    const correctedImport = await authedFetch(
      `${baseUrl}/import/confirm/${unknownPreviewPath.split('/').at(-1)}`,
      { method: 'POST', redirect: 'manual' },
    );
    assert.equal(correctedImport.status, 302);
    assert.equal(Number((await db.prepare(`
      SELECT COUNT(*) AS count FROM expenses
      WHERE title = 'Operation historique à classer' AND expense_date IS NULL AND month_key = '2026-03'
    `).get()).count), 1);

    const duplicateForm = new FormData();
    duplicateForm.append('file', new Blob([fileBuffer]), 'history-again.xlsx');
    const duplicateUpload = await authedFetch(`${baseUrl}/import`, { method: 'POST', body: duplicateForm, redirect: 'manual' });
    const duplicatePreview = await authedFetch(`${baseUrl}${duplicateUpload.headers.get('location')}`);
    assert.match(await duplicatePreview.text(), /Doublons potentiels/);
    const duplicateConfirm = await authedFetch(`${baseUrl}${duplicateUpload.headers.get('location').replace('/preview/', '/confirm/')}`, {
      method: 'POST',
      redirect: 'manual',
    });
    assert.equal(duplicateConfirm.status, 400);
    assert.equal(Number((await db.prepare('SELECT COUNT(*) AS count FROM expenses').get()).count), 6);

    await authedFetch(`${baseUrl}/expenses`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        title: 'Ajout manuel après import',
        category_id: String(categoryId),
        amount: '3000',
        expense_date: '2026-11-02',
      }),
      redirect: 'manual',
    });
    const exportResponse = await authedFetch(`${baseUrl}/export/2026`);
    assert.equal(exportResponse.status, 200);
    const exported = XLSX.read(Buffer.from(await exportResponse.arrayBuffer()), { type: 'buffer' });
    const november = XLSX.utils.sheet_to_json(exported.Sheets.Novembre, { header: 1, raw: true });
    assert.ok(november.some((row) => row[3] === 'Ajout manuel après import'));
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});
