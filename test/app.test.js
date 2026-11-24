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

function findExcelColumnRow(sheet, columnNumber, label) {
  for (let rowNumber = 21; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (sheet.getCell(rowNumber, columnNumber).value === label) return rowNumber;
  }
  return null;
}

function findXlsxColumnRow(sheet, column, label) {
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const columnNumber = XLSX.utils.decode_col(column);
  for (let rowNumber = Math.max(range.s.r, 20); rowNumber <= range.e.r; rowNumber += 1) {
    const cell = sheet[XLSX.utils.encode_cell({ r: rowNumber, c: columnNumber })];
    if (cell?.v === label) return rowNumber + 1;
  }
  return null;
}

function worksheetContentAndStyles(sheet) {
  return Array.from({ length: sheet.rowCount }, (_, rowIndex) => (
    Array.from({ length: sheet.columnCount }, (_, columnIndex) => {
      const cell = sheet.getCell(rowIndex + 1, columnIndex + 1);
      return { value: cell.value, style: cell.style, numFmt: cell.numFmt };
    })
  ));
}

test('GET / returns dashboard HTML', async () => {
  const server = await startServer();
  const { port } = server.address();

  try {
    const response = await authedFetch(`http://127.0.0.1:${port}/`);
    assert.equal(response.status, 200);
    const text = await response.text();
    assert.match(text, /Gestion des dépenses familiales/i);
    assert.match(text, /class="metric-card metric-income dashboard-metric-link" href="\/income\?month=/);
    assert.match(text, /class="metric-card metric-expense dashboard-metric-link" href="\/expenses\?month=/);
    const arrearsCard = text.match(/<a class="metric-card metric-arrears dashboard-metric-link"[\s\S]*?<\/a>/);
    assert.ok(arrearsCard);
    assert.match(arrearsCard[0], /href="\/rents\?month=[^"]+"/);
    assert.doesNotMatch(arrearsCard[0], /data-modal-template/);
    assert.match(text, /class="metric-card metric-savings dashboard-metric-link" href="\/savings"/);
    assert.match(text, /src="\/theme\.js"/);
    assert.match(text, /data-theme-toggle/);
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
    assert.match(html, /data-theme-toggle/);
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});

test('browser tab icon is linked from all page layouts and served as SVG', async () => {
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    for (const [route, request] of [
      ['/', (url) => authedFetch(url)],
      ['/expenses', (url) => authedFetch(url)],
      ['/login', (url) => fetch(url)],
    ]) {
      const response = await request(`${baseUrl}${route}`);
      assert.equal(response.status, 200);
      assert.match(await response.text(), /<link rel="icon" type="image\/svg\+xml" href="\/favicon\.svg"\s*\/>/);
    }

    const icon = await fetch(`${baseUrl}/favicon.svg`);
    assert.equal(icon.status, 200);
    assert.match(icon.headers.get('content-type'), /image\/svg\+xml/);
    assert.match(await icon.text(), /<svg[^>]+viewBox="0 0 64 64"/);
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

test('grocery page filters the selected month and accepts total-only pricing', async () => {
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;
  const totalPricedItem = 'Marchandise total prix test';
  const unitPricedItem = 'Marchandise prix unitaire test';

  try {
    const totalPriceResponse = await authedFetch(`${baseUrl}/groceries`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        item_name: totalPricedItem,
        purchase_date: '2040-04-04',
        selected_month: '2040-04',
        quantity: '3',
        unit: 'sacs',
        unit_price: '',
        total_amount: '8500',
      }),
    });
    assert.equal(totalPriceResponse.status, 200);
    const totalPriceHtml = await totalPriceResponse.text();
    assert.match(totalPriceHtml, /name="month" value="2040-04"/);
    assert.match(totalPriceHtml, /<input type="text" name="item_name"[^>]*data-modal-initial-focus/);
    assert.match(totalPriceHtml, new RegExp(totalPricedItem));
    assert.doesNotMatch(totalPriceHtml, new RegExp(unitPricedItem));

    const totalPriceRow = await db.prepare(`
      SELECT id, quantity, unit_price, total_amount, unit_price_known
      FROM grocery_items WHERE item_name = ?
    `).get(totalPricedItem);
    const groceryId = Number(totalPriceRow.id);
    assert.match(totalPriceHtml, new RegExp(`data-modal-template="grocery-edit-${groceryId}"`));
    assert.equal(Number(totalPriceRow.quantity), 3);
    assert.equal(Number(totalPriceRow.unit_price), 0);
    assert.equal(Number(totalPriceRow.total_amount), 8500);
    assert.equal(Number(totalPriceRow.unit_price_known), 0);

    const editGroceryResponse = await authedFetch(`${baseUrl}/groceries/${groceryId}/update`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        item_name: totalPricedItem,
        purchase_date: '2040-04-04',
        selected_month: '2040-04',
        quantity: '3',
        unit: 'sacs',
        unit_price: '',
        total_amount: '9000',
        notes: 'Montant corrigé',
        return_to: '/history',
      }),
    });
    assert.equal(editGroceryResponse.status, 200);
    const historyHtml = await editGroceryResponse.text();
    assert.match(historyHtml, new RegExp(`data-modal-template="grocery-edit-${groceryId}"`));
    assert.match(historyHtml, /class="table-actions history-operation-actions"/);
    assert.match(historyHtml, new RegExp(`data-modal-confirm="/groceries/${groceryId}/delete"`));
    const updatedGrocery = await db.prepare(`
      SELECT total_amount, notes FROM grocery_items WHERE id = ?
    `).get(groceryId);
    assert.equal(Number(updatedGrocery.total_amount), 9000);
    assert.equal(updatedGrocery.notes, 'Montant corrigé');

    const unitPriceResponse = await authedFetch(`${baseUrl}/groceries`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        item_name: unitPricedItem,
        purchase_date: '2040-05-04',
        selected_month: '2040-05',
        quantity: '2',
        unit: 'paquets',
        unit_price: '275',
        total_amount: '',
      }),
    });
    assert.equal(unitPriceResponse.status, 200);
    const unitPriceHtml = await unitPriceResponse.text();
    assert.match(unitPriceHtml, /name="month" value="2040-05"/);
    assert.match(unitPriceHtml, new RegExp(unitPricedItem));
    assert.doesNotMatch(unitPriceHtml, new RegExp(totalPricedItem));
    assert.equal(Number((await db.prepare(`
      SELECT total_amount FROM grocery_items WHERE item_name = ?
    `).get(unitPricedItem)).total_amount), 550);

    const filteredAprilResponse = await authedFetch(`${baseUrl}/groceries?month=2040-04`);
    const filteredAprilHtml = await filteredAprilResponse.text();
    assert.match(filteredAprilHtml, new RegExp(totalPricedItem));
    assert.doesNotMatch(filteredAprilHtml, new RegExp(unitPricedItem));

    const invalidPriceResponse = await authedFetch(`${baseUrl}/groceries`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        item_name: 'Marchandise sans prix test',
        purchase_date: '2040-05-05',
        quantity: '1',
        unit_price: '',
        total_amount: '',
      }),
    });
    assert.equal(invalidPriceResponse.status, 400);
  } finally {
    await db.prepare('DELETE FROM grocery_items WHERE item_name IN (?, ?)').run(totalPricedItem, unitPricedItem);
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
      assert.match(text, /class="side-nav" aria-label="Sections de l’application"/);
      assert.match(text, /class="sidebar-user"/);
      assert.match(text, /id="main-sidebar"/);
      assert.match(text, /data-mobile-sidebar-toggle/);
      assert.match(text, /id="global-search"/);
      assert.match(text, /aria-label="Rechercher sur cette page"/);
      assert.match(text, /class="mobile-nav-overlay" data-mobile-nav-close hidden/);
      assert.doesNotMatch(text, /mobile-primary-nav|mobile-more-trigger/);
      assert.match(text, /src="\/app.js"/);
    }

    const expenses = await authedFetch(`http://127.0.0.1:${port}/expenses`);
    const expensesHtml = await expenses.text();
    assert.match(expensesHtml, /data-modal-template="expense-create"/);
    assert.match(expensesHtml, /class="close-modal"[^>]*data-close-modal/);
    assert.match(expensesHtml, /class="sidebar-logout"/);
    assert.doesNotMatch(expensesHtml, /href="\/groceries" data-modal-link/);
    const dashboard = await authedFetch(`http://127.0.0.1:${port}/`);
    const dashboardHtml = await dashboard.text();
    assert.match(dashboardHtml, /class="close-modal"[^>]*data-close-modal/);
    assert.match(dashboardHtml, /data-mobile-sidebar-toggle/);
    assert.match(dashboardHtml, /aria-label="Rechercher dans l’application"/);
    const incomePage = await authedFetch(`http://127.0.0.1:${port}/income?month=2026-11`);
    const incomeHtml = await incomePage.text();
    assert.match(incomeHtml, /Revenus mensuels/);
    assert.match(incomeHtml, /class="income-source-card"/);
    assert.match(incomeHtml, /class="income-source-card income-source-card-clickable" data-modal-template="rent-payments-detail"/);
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
    const octoberFinance = await (
      await authedFetch(`${baseUrl}/finance/monthly-data?month=2026-10`)
    ).json();
    assert.equal(
      exportedMonth.Sheets.Novembre.B23.v,
      octoberFinance.closingBalance - octoberFinance.openingBalance,
    );
    assert.equal(exportedMonth.Sheets.Novembre.A24.v, 'Total des ressources disponibles');
    assert.equal(exportedMonth.Sheets.Novembre.B24.v, finance.incomeTotal);
    const novemberExpenseTotalRow = findXlsxColumnRow(exportedMonth.Sheets.Novembre, 'D', 'Total général');
    assert.ok(novemberExpenseTotalRow);
    assert.ok(exportedMonth.Sheets.Novembre.B22.f === 'B24-0' ||
      /^B24-SUM\(F/.test(exportedMonth.Sheets.Novembre.B22.f));
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
    assert.equal(
      uncheckedReportRow.getCell(2).value,
      finance.closingBalance - finance.openingBalance,
    );
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
    assert.match(historyHtml, /data-row-href="\/rents\?month=2026-11&amp;open=rent-payment-details-\d+"/);
    assert.match(historyHtml, /data-row-href="\/income\?month=2026-11&amp;open=income-details-\d+"/);
    assert.match(historyHtml, /data-row-href="\/savings\?open=savings-details-\d+"/);
  } finally {
    server.close();
  }
});

test('rent payments can omit dates and mark all remaining obligations paid', async () => {
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;
  const postForm = (route, values) => authedFetch(`${baseUrl}${route}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(values),
  });

  try {
    await postForm('/units', { floor: 'Rez-de-chaussée', label: 'BULK-PAYMENT-1', housing_type: 'Studio' });
    const unit = await db.prepare("SELECT id FROM units WHERE label = 'BULK-PAYMENT-1'").get();
    assert.ok(unit);
    await postForm('/tenants', {
      unit_id: String(unit.id),
      first_name: 'Paiement',
      last_name: 'Groupé',
      people_count: '1',
      monthly_rent: '100000',
      entry_date: '2026-01-01',
      phone: '',
    });
    await postForm('/rents/generate', { month_key: '2027-01' });
    const obligation = await db.prepare(`
      SELECT o.id FROM rent_obligations o JOIN tenants t ON t.id = o.tenant_id
      WHERE t.first_name = 'Paiement' AND t.last_name = 'Groupé' AND o.month_key = '2027-01'
    `).get();
    assert.ok(obligation);

    const partialPayment = await postForm(`/rents/${obligation.id}/payments`, { amount: '25000' });
    assert.equal(partialPayment.status, 200);
    const rentsPage = await (await authedFetch(`${baseUrl}/rents?month=2027-01`)).text();
    assert.match(rentsPage, /Marquer tous les loyers comme payés/);
    assert.match(rentsPage, /Date du paiement \(facultative\)/);
    assert.match(rentsPage, /name="payment_date"[^>]*\/>/);
    assert.match(rentsPage, /name="amount"[^>]*max="75000"[^>]*value="75000"/);

    const remainingObligations = await db.prepare(`
      SELECT COUNT(*) AS count FROM (
        SELECT o.id, o.expected_amount, COALESCE(SUM(p.amount), 0) AS amount_paid
        FROM rent_obligations o
        LEFT JOIN rent_payments p ON p.obligation_id = o.id
        WHERE o.month_key = ?
        GROUP BY o.id
      ) AS month_obligations
      WHERE expected_amount > amount_paid
    `).get('2027-01');
    const markAllResponses = await Promise.all([
      postForm('/rents/mark-all-paid', { month_key: '2027-01' }),
      postForm('/rents/mark-all-paid', { month_key: '2027-01' }),
    ]);
    assert.ok(markAllResponses.every((response) => response.status === 200));
    const markAllHtml = await Promise.all(markAllResponses.map((response) => response.text()));
    const markedCount = Number(remainingObligations.count);
    const markedNotice = `${markedCount} loyer${markedCount === 1 ? '' : 's'} marqué${markedCount === 1 ? '' : 's'} comme payé${markedCount === 1 ? '' : 's'}.`;
    assert.ok(markAllHtml.some((html) => html.includes(markedNotice)));
    assert.ok(markAllHtml.some((html) => html.includes('Aucun loyer restant à payer.')));

    const payments = await db.prepare(`
      SELECT amount, payment_date, received_month_key FROM rent_payments
      WHERE obligation_id = ? ORDER BY id
    `).all(obligation.id);
    assert.deepEqual(payments.map((payment) => Number(payment.amount)), [25000, 75000]);
    assert.ok(payments.every((payment) => payment.payment_date === null));
    assert.ok(payments.every((payment) => payment.received_month_key === '2027-01'));

    for (const index of [1, 2]) {
      const unitLabel = `BULK-SELECT-${index}`;
      await postForm('/units', { floor: 'Rez-de-chaussée', label: unitLabel, housing_type: 'Studio' });
      const selectedUnit = await db.prepare('SELECT id FROM units WHERE label = ?').get(unitLabel);
      await postForm('/tenants', {
        unit_id: String(selectedUnit.id),
        first_name: `Choix${index}`,
        last_name: 'Paiement',
        people_count: '1',
        monthly_rent: String(50000 * index),
        entry_date: '2026-01-01',
        phone: '',
      });
    }
    await postForm('/rents/generate', { month_key: '2027-03' });
    const selectedObligations = await db.prepare(`
      SELECT o.id, o.expected_amount FROM rent_obligations o
      JOIN tenants t ON t.id = o.tenant_id
      WHERE t.first_name IN ('Choix1', 'Choix2') AND o.month_key = '2027-03'
      ORDER BY t.first_name
    `).all();
    assert.equal(selectedObligations.length, 2);
    const selectedPage = await (await authedFetch(`${baseUrl}/rents?month=2027-03`)).text();
    assert.match(selectedPage, /data-rent-select-all/);
    assert.match(selectedPage, /data-rent-selected-trigger disabled/);
    assert.match(selectedPage, /id="rent-mark-selected-paid"/);
    const selectedForm = selectedPage.split('id="rent-mark-selected-paid"')[1].split('</template>')[0];
    assert.match(selectedForm, /type="date" name="payment_date"/);
    assert.doesNotMatch(selectedForm, /name="payment_date"[^>]*required/);

    const selectedBody = new URLSearchParams({ month_key: '2027-03' });
    selectedObligations.forEach((row) => selectedBody.append('obligation_ids', String(row.id)));
    const selectedResponse = await authedFetch(`${baseUrl}/rents/mark-selected-paid`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: selectedBody,
    });
    assert.equal(selectedResponse.status, 200);
    assert.match(await selectedResponse.text(), /2 loyers marqués comme payés/);
    const selectedPayments = await db.prepare(`
      SELECT p.amount, p.payment_date, p.received_month_key FROM rent_payments p
      WHERE p.obligation_id = ANY(?)
      ORDER BY p.obligation_id
    `).all(selectedObligations.map((row) => row.id));
    assert.equal(selectedPayments.length, 2);
    assert.deepEqual(
      selectedPayments.map((payment) => Number(payment.amount)),
      selectedObligations.map((row) => Number(row.expected_amount)),
    );
    assert.ok(selectedPayments.every((payment) => payment.payment_date === null));
    assert.ok(selectedPayments.every((payment) => payment.received_month_key === '2027-03'));
    const unselectedObligation = await db.prepare(`
      SELECT o.id FROM rent_obligations o
      JOIN tenants t ON t.id = o.tenant_id
      LEFT JOIN rent_payments p ON p.obligation_id = o.id
      WHERE t.first_name = 'Paiement' AND t.last_name = 'Groupé' AND o.month_key = '2027-03'
      GROUP BY o.id HAVING COALESCE(SUM(p.amount), 0) = 0
    `).get();
    assert.ok(unselectedObligation);

    const incomePage = await (await authedFetch(`${baseUrl}/income?month=2027-03`)).text();
    assert.match(incomePage, /2 paiements reçus ce mois/);
    assert.match(incomePage, /data-modal-confirm="\/rents\/payments\/\d+\/delete"/);
    const historyPage = await (await authedFetch(`${baseUrl}/history?type=rent&month=2027-03`)).text();
    assert.match(historyPage, /data-modal-confirm="\/rents\/payments\/\d+\/delete"/);

    const paymentToDelete = await db.prepare(`
      SELECT id FROM rent_payments WHERE obligation_id = ? ORDER BY id LIMIT 1
    `).get(selectedObligations[0].id);
    assert.ok(paymentToDelete);
    const deleteResponse = await postForm(`/rents/payments/${paymentToDelete.id}/delete`, {
      return_to: '/income?month=2027-03',
    });
    assert.equal(deleteResponse.status, 200);
    assert.match(await deleteResponse.text(), /1 paiement reçu ce mois/);
    assert.equal(Number((await db.prepare(`
      SELECT COUNT(*) AS count FROM rent_payments WHERE id = ?
    `).get(paymentToDelete.id)).count), 0);
    const deletionAudit = await db.prepare(`
      SELECT before_data, after_data FROM audit_log
      WHERE entity_type = 'rent_payment' AND entity_id = ? AND action = 'delete'
    `).get(String(paymentToDelete.id));
    assert.ok(deletionAudit);
    assert.equal(deletionAudit.before_data.id, paymentToDelete.id);
    assert.equal(deletionAudit.after_data, null);
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});

test('rent Excel statuses match each month obligation and the rents page', async () => {
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;
  const cases = [
    { key: 'fully-paid', firstName: 'Loyer', lastName: 'Intégral', expected: 50_000, paid: 50_000, status: 'PAYÉ' },
    { key: 'unpaid', firstName: 'Loyer', lastName: 'Absent', expected: 50_000, paid: 0, status: 'IMPAYÉ' },
    { key: 'partial', firstName: 'Loyer', lastName: 'Partiel', expected: 50_000, paid: 30_000, status: 'PARTIELLEMENT PAYÉ' },
    {
      key: 'installments',
      firstName: 'Loyer',
      lastName: 'Versements',
      expected: 100_000,
      payments: [40_000, 60_000],
      status: 'PAYÉ',
    },
    { key: 'other-month', firstName: 'Loyer', lastName: 'AutreMois', expected: 70_000, paid: 0, status: 'IMPAYÉ' },
    { key: 'overpaid', firstName: 'Loyer', lastName: 'TropPerçu', expected: 50_000, paid: 60_000, status: 'PAYÉ' },
    {
      key: 'changed-rent',
      firstName: 'Loyer',
      lastName: 'Modifié',
      expected: 50_000,
      paid: 50_000,
      monthlyRent: 80_000,
      status: 'PAYÉ',
    },
  ];
  const unitIds = [];
  const tenantIds = [];

  try {
    for (const renter of cases) {
      const unit = await db.prepare(`
        INSERT INTO units (floor, label, housing_type) VALUES (?, ?, 'Studio')
      `).run('Rez-de-chaussée', `STATUS-${renter.key}`);
      unitIds.push(unit.lastInsertRowid);
      const tenant = await db.prepare(`
        INSERT INTO tenants (
          unit_id, first_name, last_name, people_count, monthly_rent, entry_date
        ) VALUES (?, ?, ?, 1, ?, '2026-01-01')
      `).run(
        unit.lastInsertRowid,
        renter.firstName,
        renter.lastName,
        renter.monthlyRent || renter.expected,
      );
      tenantIds.push(tenant.lastInsertRowid);

      const obligation = await db.prepare(`
        INSERT INTO rent_obligations (tenant_id, month_key, expected_amount)
        VALUES (?, '2027-06', ?)
      `).run(tenant.lastInsertRowid, renter.expected);
      let payments = renter.payments || (renter.paid ? [renter.paid] : []);
      if (renter.key === 'other-month') {
        const mayObligation = await db.prepare(`
          INSERT INTO rent_obligations (tenant_id, month_key, expected_amount)
          VALUES (?, '2027-05', ?)
        `).run(tenant.lastInsertRowid, renter.expected);
        await db.prepare(`
          INSERT INTO rent_payments (
            obligation_id, amount, payment_date, received_month_key
          ) VALUES (?, ?, '2027-06-02', '2027-06')
        `).run(mayObligation.lastInsertRowid, renter.expected);
        payments = [];
      }
      for (const [index, amount] of payments.entries()) {
        await db.prepare(`
          INSERT INTO rent_payments (
            obligation_id, amount, payment_date, received_month_key
          ) VALUES (?, ?, ?, '2027-07')
        `).run(obligation.lastInsertRowid, amount, `2027-07-${String(index + 1).padStart(2, '0')}`);
      }
    }

    const rentsHtml = await (await authedFetch(`${baseUrl}/rents?month=2027-06`)).text();
    const visibleRows = [...rentsHtml.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/g)].map((match) => match[1]);
    for (const renter of cases) {
      const fullName = `${renter.firstName} ${renter.lastName}`;
      const visibleRow = visibleRows.find((row) => row.includes(`<strong>${fullName}</strong>`));
      assert.ok(visibleRow, `missing rents-page row for ${fullName}`);
      assert.match(visibleRow, new RegExp(`<span class="category-pill">${renter.status}</span>`));
    }

    async function exportedStatuses(url) {
      const response = await authedFetch(`${baseUrl}${url}`);
      assert.equal(response.status, 200);
      const workbook = await loadExcelJsWorkbook(Buffer.from(await response.arrayBuffer()));
      const sheet = workbook.getWorksheet('Juillet');
      return new Map(cases.map((renter) => {
        const fullName = `${renter.firstName} ${renter.lastName}`;
        let rowNumber = 3;
        while (rowNumber <= sheet.rowCount && sheet.getCell(rowNumber, 2).value !== fullName) rowNumber += 1;
        assert.ok(rowNumber <= sheet.rowCount, `missing Excel row for ${fullName}`);
        assert.equal(Number(sheet.getCell(rowNumber, 6).value), renter.expected);
        return [renter.key, sheet.getCell(rowNumber, 7).value];
      }));
    }

    const monthlyStatuses = await exportedStatuses('/export/month/2027/7');
    assert.deepEqual([...monthlyStatuses.values()], cases.map((renter) => renter.status));
    assert.deepEqual(
      [...await exportedStatuses('/export/month?year=2027&month=7').then((statuses) => statuses.values())],
      cases.map((renter) => renter.status),
    );
    assert.deepEqual(
      [...await exportedStatuses('/export/months?year=2027&months=7&months=8').then((statuses) => statuses.values())],
      cases.map((renter) => renter.status),
    );
    assert.deepEqual(
      [...await exportedStatuses('/export/2027').then((statuses) => statuses.values())],
      cases.map((renter) => renter.status),
    );
    assert.equal(monthlyStatuses.get('other-month'), 'IMPAYÉ');
  } finally {
    await new Promise((resolve) => server.close(resolve));
    if (tenantIds.length) {
      await db.prepare(`
        DELETE FROM rent_payments
        WHERE obligation_id IN (SELECT id FROM rent_obligations WHERE tenant_id = ANY(?))
      `).run(tenantIds);
      await db.prepare('DELETE FROM rent_obligations WHERE tenant_id = ANY(?)').run(tenantIds);
      await db.prepare('DELETE FROM tenants WHERE id = ANY(?)').run(tenantIds);
    }
    if (unitIds.length) await db.prepare('DELETE FROM units WHERE id = ANY(?)').run(unitIds);
  }
});

test('protected-savings expenses stay separate from current balances and appear in exports', async () => {
  const protectedAccount = await db.prepare(`
    SELECT id FROM savings_accounts WHERE is_protected = TRUE
  `).get();
  const category = await db.prepare("SELECT id, name FROM categories WHERE type <> 'boutique' ORDER BY id LIMIT 1").get();
  assert.ok(protectedAccount);
  assert.ok(category);
  const monthKey = '2027-02';
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const before = await (await authedFetch(
      `${baseUrl}/finance/monthly-data?month=${monthKey}`,
    )).json();
    const unconfirmed = await authedFetch(`${baseUrl}/savings/protected-expenses`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        title: 'Dépense protégée sans confirmation',
        category_id: String(category.id),
        expense_date: '2027-02-12',
        month_key: monthKey,
        amount: '2500',
      }),
    });
    assert.equal(unconfirmed.status, 400);
    const response = await authedFetch(`${baseUrl}/savings/protected-expenses`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        title: 'Dépense test épargne protégée',
        category_id: String(category.id),
        expense_date: '2027-02-12',
        month_key: monthKey,
        amount: '2500',
        protected_confirmation: '1',
      }),
      redirect: 'manual',
    });
    assert.equal(response.status, 302);
    const linked = await db.prepare(`
      SELECT e.funding_source, e.amount, t.transaction_type, t.related_expense_id
      FROM expenses e JOIN savings_transactions t ON t.related_expense_id = e.id
      WHERE e.title = 'Dépense test épargne protégée'
    `).get();
    assert.equal(linked.funding_source, 'protected_savings');
    assert.equal(Number(linked.amount), 2500);
    assert.equal(linked.transaction_type, 'withdrawal');
    assert.ok(linked.related_expense_id);

    const after = await (await authedFetch(
      `${baseUrl}/finance/monthly-data?month=${monthKey}`,
    )).json();
    assert.equal(after.totalExpenses, before.totalExpenses);
    assert.equal(after.closingBalance, before.closingBalance);
    assert.equal(after.protectedSavingsExpenses, before.protectedSavingsExpenses + 2500);

    const workbookResponse = await authedFetch(`${baseUrl}/export/month/2027/2`);
    const workbook = await loadExcelJsWorkbook(Buffer.from(await workbookResponse.arrayBuffer()));
    const sheet = workbook.getWorksheet('Février');
    let protectedExpenseRow = null;
    for (let rowNumber = 22; rowNumber <= sheet.rowCount; rowNumber += 1) {
      if (sheet.getCell(rowNumber, 4).value === category.name &&
        sheet.getCell(rowNumber, 5).value === 'Dépense test épargne protégée') {
        protectedExpenseRow = rowNumber;
        break;
      }
    }
    assert.ok(protectedExpenseRow);
    assert.equal(sheet.getCell(protectedExpenseRow, 6).value, 2500);

    const csvResponse = await authedFetch(`${baseUrl}/export/csv/2027/expenses`);
    const csv = await csvResponse.text();
    assert.match(csv, /^"Date";"Mois";"Désignation";"Catégorie";"Montant";"Source de financement"/);
    assert.match(csv, /Dépense test épargne protégée/);
    assert.match(csv, /Épargne protégée/);
  } finally {
    const createdExpense = await db.prepare(`
      SELECT id FROM expenses WHERE title = 'Dépense test épargne protégée'
    `).get();
    if (createdExpense) {
      await db.prepare('DELETE FROM savings_transactions WHERE related_expense_id = ?')
        .run(createdExpense.id);
      await db.prepare('DELETE FROM expenses WHERE id = ?').run(createdExpense.id);
    }
    await new Promise((resolve) => server.close(resolve));
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
    const allSearchResults = await (await authedFetch(
      `${baseUrl}/history?type=expense&month=2026-11&search=pagination&format=json`,
    )).json();
    assert.equal(allSearchResults.totalCount, 51);
    assert.equal(allSearchResults.operations.length, 51);
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
    const expensesHtml = await response.text();
    assert.match(expensesHtml, /<tr data-modal-template="expense-details-\d+" tabindex="0"/);
    const expense = await db.prepare("SELECT * FROM expenses WHERE title = 'Electricité'").get();
    const detailsStart = expensesHtml.indexOf(`<template id="expense-details-${expense.id}">`);
    const detailsEnd = expensesHtml.indexOf('</template>', detailsStart);
    const detailsHtml = expensesHtml.slice(detailsStart, detailsEnd);
    assert.match(expensesHtml, /aria-label="Afficher le détail de Electricité"/);
    assert.match(detailsHtml, /aria-label="Modifier Electricité"/);
    assert.match(detailsHtml, /aria-label="Supprimer Electricité"/);
    assert.match(detailsHtml, /data-modal-template="expense-edit-\d+">Modifier<\/button>/);
    assert.match(detailsHtml, /data-modal-danger="true">Supprimer<\/button>/);
    assert.equal(Number(expense.amount), 25000);
    assert.equal(expense.beneficiary, 'Fournisseur test');
    assert.equal(expense.month_key, '2026-11');
    assert.ok(expense.created_by);
    const attachment = await db.prepare('SELECT * FROM expense_attachments WHERE expense_id = ?').get(expense.id);
    assert.equal(attachment.file_name, 'preuve.png');
    const dashboardHtml = await (await authedFetch(`http://127.0.0.1:${port}/?month=2026-11`)).text();
    assert.match(dashboardHtml, new RegExp(`data-row-href="/expenses\\?[^"]*open=expense-details-${expense.id}"`));
    assert.match(dashboardHtml, /id="global-search-results"/);
    const searchResponse = await authedFetch(`http://127.0.0.1:${port}/history?format=json&search=Electricité`);
    const searchData = await searchResponse.json();
    const searchResult = searchData.operations.find((item) => item.type === 'expense' && item.id === expense.id);
    assert.equal(searchResponse.status, 200);
    assert.equal(searchResult.href, `/expenses?month=2026-11&search=Electricit%C3%A9&open=expense-details-${expense.id}`);
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
    return_to: '/history',
  };

  try {
    const historyResponse = await authedFetch(`${baseUrl}/history?type=expense&search=Electricité`);
    const historyHtml = await historyResponse.text();
    assert.match(historyHtml, new RegExp(`data-modal-template="expense-edit-${expense.id}"`));
    assert.match(historyHtml, new RegExp(`action="/expenses/${expense.id}/update"`));
    assert.match(historyHtml, new RegExp(`data-row-href="/expenses\\?month=2026-11&amp;search=Electricit%C3%A9&amp;open=expense-details-${expense.id}"`));
    assert.match(historyHtml, /name="return_to" value="\/history"/);

    const updateResponse = await authedFetch(`${baseUrl}/expenses/${expense.id}/update`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams(fields),
    });
    assert.equal(updateResponse.status, 200);
    assert.equal(new URL(updateResponse.url).pathname, '/history');
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

test('expenses support optional dates and bulk duplicate/delete actions from their history', async () => {
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;
  const category = await db.prepare("SELECT id FROM categories WHERE type <> 'boutique' ORDER BY id LIMIT 1").get();
  const first = await db.prepare(`
    INSERT INTO expenses (title, category_id, amount, expense_date, month_key)
    VALUES (?, ?, ?, ?, ?)
  `).run('Bulk source dated', category.id, 1200, '2027-01-31', '2027-01');
  const second = await db.prepare(`
    INSERT INTO expenses (title, category_id, amount, expense_date, month_key)
    VALUES (?, ?, ?, ?, ?)
  `).run('Bulk source undated', category.id, 800, null, '2027-01');
  const protectedExpense = await db.prepare(`
    INSERT INTO expenses (title, category_id, amount, expense_date, month_key, funding_source)
    VALUES (?, ?, ?, ?, ?, 'protected_savings')
  `).run('Protected bulk source', category.id, 500, null, '2027-01');

  try {
    const expensesPage = await authedFetch(`${baseUrl}/expenses?month=2027-01`);
    const expensesHtml = await expensesPage.text();
    assert.match(expensesHtml, /Date de la dépense \(facultative\)/);
    assert.match(expensesHtml, new RegExp(`name="expense_ids" value="${first.lastInsertRowid}"`));
    assert.match(expensesHtml, new RegExp(`name="expense_ids" value="${second.lastInsertRowid}"`));
    assert.match(expensesHtml, /data-expense-select-all/);
    assert.match(expensesHtml, /Dupliquer vers le mois/);

    const historyPage = await authedFetch(`${baseUrl}/history?type=expense&month=2027-01`);
    const historyHtml = await historyPage.text();
    assert.match(historyHtml, /Ajouter une dépense/);
    assert.match(historyHtml, new RegExp(`name="expense_ids" value="${first.lastInsertRowid}"`));
    assert.match(historyHtml, /Sans date/);
    assert.doesNotMatch(historyHtml, new RegExp(`name="expense_ids" value="${protectedExpense.lastInsertRowid}"`));

    const createWithoutDate = await authedFetch(`${baseUrl}/expenses`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        title: 'Expense without date',
        category_id: String(category.id),
        amount: '300',
        expense_date: '',
        month_key: '2027-01',
      }),
      redirect: 'manual',
    });
    assert.equal(createWithoutDate.status, 302);
    const undatedExpense = await db.prepare("SELECT * FROM expenses WHERE title = 'Expense without date'").get();
    assert.equal(undatedExpense.expense_date, null);
    assert.equal(undatedExpense.month_key, '2027-01');

    const invalidDate = await authedFetch(`${baseUrl}/expenses`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        title: 'Expense with invalid date',
        category_id: String(category.id),
        amount: '300',
        expense_date: '2027-02-31',
        month_key: '2027-02',
      }),
      redirect: 'manual',
    });
    assert.equal(invalidDate.status, 400);
    const missingMonth = await authedFetch(`${baseUrl}/expenses`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        title: 'Undated expense without month',
        category_id: String(category.id),
        amount: '300',
        expense_date: '',
      }),
      redirect: 'manual',
    });
    assert.equal(missingMonth.status, 400);

    const duplicateFields = new URLSearchParams({
      action: 'duplicate',
      target_month: '2027-02',
      return_to: '/history',
    });
    [first.lastInsertRowid, second.lastInsertRowid].forEach((id) => duplicateFields.append('expense_ids', String(id)));
    const duplicateResponse = await authedFetch(`${baseUrl}/expenses/bulk`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: duplicateFields,
      redirect: 'manual',
    });
    assert.equal(duplicateResponse.status, 302);
    assert.equal(duplicateResponse.headers.get('location'), '/history?month=2027-02&type=expense');
    const duplicates = await db.prepare(`
      SELECT expense_date, month_key, amount FROM expenses
      WHERE title IN ('Bulk source dated', 'Bulk source undated')
        AND month_key = '2027-02'
      ORDER BY title
    `).all();
    assert.equal(duplicates.length, 2);
    assert.equal(duplicates[0].expense_date, '2027-02-28');
    assert.equal(duplicates[1].expense_date, null);
    assert.ok(duplicates.every((expense) => expense.month_key === '2027-02'));
    assert.equal(Number((await db.prepare(`
      SELECT COUNT(*) AS count FROM expenses WHERE id IN (?, ?)
    `).get(first.lastInsertRowid, second.lastInsertRowid)).count), 2);
    const protectedDeleteResponse = await authedFetch(`${baseUrl}/expenses/bulk`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        action: 'delete',
        expense_ids: String(protectedExpense.lastInsertRowid),
        return_to: '/history',
      }),
      redirect: 'manual',
    });
    assert.equal(protectedDeleteResponse.status, 409);
    assert.ok(await db.prepare('SELECT id FROM expenses WHERE id = ?').get(protectedExpense.lastInsertRowid));
    const staleSelectionResponse = await authedFetch(`${baseUrl}/expenses/bulk`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        action: 'duplicate',
        expense_ids: '9007199254740991',
        target_month: '2027-02',
        return_to: '/history',
      }),
      redirect: 'manual',
    });
    assert.equal(staleSelectionResponse.status, 404);
    assert.match(await staleSelectionResponse.text(), /Actualisez l’historique puis réessayez/);

    const deleteFields = new URLSearchParams({
      action: 'delete',
      month_key: '2027-01',
      return_to: '/history',
    });
    [first.lastInsertRowid, second.lastInsertRowid].forEach((id) => deleteFields.append('expense_ids', String(id)));
    const deleteResponse = await authedFetch(`${baseUrl}/expenses/bulk`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: deleteFields,
      redirect: 'manual',
    });
    assert.equal(deleteResponse.status, 302);
    assert.equal(deleteResponse.headers.get('location'), '/history?month=2027-01&type=expense');
    assert.equal(Number((await db.prepare(`
      SELECT COUNT(*) AS count FROM expenses WHERE id IN (?, ?)
    `).get(first.lastInsertRowid, second.lastInsertRowid)).count), 0);
    assert.equal(Number((await db.prepare(`
      SELECT COUNT(*) AS count FROM expenses
      WHERE title IN ('Bulk source dated', 'Bulk source undated') AND month_key = '2027-02'
    `).get()).count), 2);
  } finally {
    await db.prepare("DELETE FROM expenses WHERE title IN ('Bulk source dated', 'Bulk source undated', 'Expense without date', 'Protected bulk source')").run();
    await new Promise((resolve) => server.close(resolve));
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
    assert.match(exportHtml, /name="includeAnnualSummary" value="1" checked/);
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
    assert.match(workbook.Sheets.Janvier.D20.v, /DÉPENSES DU MOIS DE JANVIER/);
    const januaryExpenseTotalRow = findXlsxColumnRow(workbook.Sheets.Janvier, 'D', 'Total général');
    assert.ok(januaryExpenseTotalRow);
    assert.equal(workbook.Sheets.Janvier[`F${januaryExpenseTotalRow}`].v, 0);
    assert.equal(workbook.Sheets['Récapitulatif annuel'].A1.v, 'Récapitulatif annuel 2026');
    assert.ok(workbook.Sheets.Janvier['!merges'].some((range) => range.s.r === 0 && range.s.c === 0 && range.e.c === 6));
    assert.ok(workbook.Sheets.Janvier['!merges'].some((range) => range.s.r === 19 && range.s.c === 3 && range.e.c === 5));
    const uncheckedAnnualResponse = await authedFetch(
      `http://127.0.0.1:${port}/export/2026?includeCarryover=0&includeAnnualSummary=0`,
    );
    const uncheckedAnnual = await loadExcelJsWorkbook(
      Buffer.from(await uncheckedAnnualResponse.arrayBuffer()),
    );
    assert.equal(uncheckedAnnual.getWorksheet('Récapitulatif annuel'), undefined);
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
      const monthIndex = workbook.SheetNames.indexOf(monthName);
      const previousMonthKey = monthIndex > 0
        ? `2026-${String(monthIndex).padStart(2, '0')}`
        : null;
      const previousFinance = previousMonthKey ? exportedFinances.get(previousMonthKey) : null;
      assert.equal(
        reportRow.getCell(2).value,
        previousFinance ? previousFinance.closingBalance - previousFinance.openingBalance : 0,
      );
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
    assert.match(exportHtml, /class="export-carryover-option"/);
    assert.match(exportHtml, /Le report reste visible dans le fichier, même si cette option est désactivée/);
    assert.match(exportHtml, /name="includeCarryover" value="1"/);
    assert.doesNotMatch(exportHtml, /name="includeCarryover" value="1" checked/);

    const workbookResponse = await authedFetch(`${baseUrl}/export/month/2026/11`);
    assert.equal(workbookResponse.status, 200);
    const workbook = XLSX.read(Buffer.from(await workbookResponse.arrayBuffer()), { type: 'buffer' });
    assert.deepEqual(workbook.SheetNames, ['Novembre']);
    assert.match(workbook.Sheets.Novembre.A1.v, /Loyers du mois d'Octobre/);
    const novemberExpenseTotalRow = findXlsxColumnRow(workbook.Sheets.Novembre, 'D', 'Total général');
    assert.ok(novemberExpenseTotalRow);
    assert.equal(workbook.Sheets.Novembre[`F${novemberExpenseTotalRow}`].v, 0);
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

test('expenses page shows the selected month shop-account total', async () => {
  const result = await db.prepare(`
    INSERT INTO grocery_items (item_name, purchase_date, month_key, quantity, unit, unit_price, total_amount)
    VALUES ('Compte boutique affichage test', '2027-12-03', '2027-12', 1, 'unité', 1234567, 1234567)
  `).run();
  const server = await startServer();
  const { port } = server.address();

  try {
    const response = await authedFetch(`http://127.0.0.1:${port}/expenses?month=2027-12`);
    const html = await response.text();
    const total = Number((await db.prepare(`
      SELECT SUM(total_amount) AS total FROM grocery_items
      WHERE COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM')) = '2027-12'
    `).get()).total);
    const formattedTotal = new Intl.NumberFormat('fr-FR', {
      style: 'currency',
      currency: 'XAF',
      minimumFractionDigits: 0,
      maximumFractionDigits: 0,
    }).format(total);

    assert.equal(response.status, 200);
    assert.match(html, /Compte boutique du mois/);
    assert.match(html, /Liste des dépenses/);
    assert.match(html, /class="expense-table-header-actions"[\s\S]*?data-expense-month="2027-12"/);
    assert.match(html, /href="\/export\/month\/2027\/12"/);
    assert.match(html, /data-expense-month-offset="-1"/);
    assert.match(html, /data-expense-current-month/);
    assert.match(html, /data-category-option-search/);
    const createExpenseTemplate = html.match(/<template id="expense-create">([\s\S]*?)<\/template>/);
    assert.ok(createExpenseTemplate);
    assert.match(createExpenseTemplate[1], /aria-label="Rechercher une catégorie"/);
    assert.match(createExpenseTemplate[1], /data-category-select-trigger/);
    assert.match(createExpenseTemplate[1], /data-category-select-menu hidden/);
    assert.match(createExpenseTemplate[1], /role="option" data-category-option/);
    assert.match(createExpenseTemplate[1], /data-category-search-status role="status"/);
    const fallbackCategory = await db.prepare("SELECT id FROM categories WHERE name = 'Autres dépenses'").get();
    assert.ok(fallbackCategory);
    assert.match(html, new RegExp(`<option value="${fallbackCategory.id}"\\s*>Autres dépenses</option>`));
    assert.match(html, /Reste du mois \(hors report\)/);
    assert.match(html, /Solde final du mois \(report inclus\)/);
    const tableSection = html.match(/<section class="panel-card page-table-card monthly-expense-history">([\s\S]*?)<form id="expense-bulk-form"/);
    assert.ok(tableSection);
    assert.match(tableSection[1], /data-live-search data-expense-list-search/);
    assert.match(tableSection[1], /name="month" value="2027-12"/);
    assert.match(tableSection[1], /Rechercher dans les dépenses du mois/);
    assert.ok(html.includes(formattedTotal));
  } finally {
    await db.prepare('DELETE FROM grocery_items WHERE id = ?').run(result.lastInsertRowid);
    await new Promise((resolve) => server.close(resolve));
  }
});

test('monthly CSV filters remain available and expenses Excel uses the reference export workbook', async () => {
  const category = await db.prepare("SELECT id FROM categories WHERE name = 'Factures'").get();
  const titles = [
    'Export mois dépense ciblée',
    'Export mois autre dépense',
    'Export mois précédent',
  ];
  const septemberTitles = [
    'Export septembre dépense ciblée',
    'Export septembre autre dépense',
    'Export août hors période',
  ];
  const groceryTitle = 'Marché export de septembre';
  await db.prepare(`
    INSERT INTO expenses (title, category_id, amount, expense_date, month_key)
    VALUES (?, ?, 1250, '2027-12-03', '2027-12'),
      (?, ?, 2500, '2027-12-04', '2027-12'),
      (?, ?, 5000, '2027-11-03', '2027-11')
  `).run(titles[0], category.id, titles[1], category.id, titles[2], category.id);
  await db.prepare(`
    INSERT INTO expenses (title, category_id, amount, expense_date, month_key)
    VALUES (?, ?, 1250, '2026-09-03', '2026-09'),
      (?, ?, 2500, '2026-09-04', '2026-09'),
      (?, ?, 5000, '2026-08-03', '2026-08')
  `).run(septemberTitles[0], category.id, septemberTitles[1], category.id, septemberTitles[2], category.id);
  const grocery = await db.prepare(`
    INSERT INTO grocery_items
      (item_name, purchase_date, month_key, quantity, unit, unit_price, total_amount, notes)
    VALUES (?, '2026-09-06', '2026-09', 2, 'kg', 750, 1500, 'Produit de test')
  `).run(groceryTitle);
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const response = await authedFetch(
      `http://127.0.0.1:${port}/export/csv/month/2027/12/expenses?category=${category.id}&search=${encodeURIComponent('dépense ciblée')}`,
    );
    assert.equal(response.status, 200);
    assert.match(response.headers.get('content-type'), /text\/csv/);
    assert.match(response.headers.get('content-disposition'), /MAISON_DEPENSES_2027-12\.csv/);
    const csv = await response.text();
    assert.match(csv, /Export mois dépense ciblée/);
    assert.doesNotMatch(csv, /Export mois autre dépense/);
    assert.doesNotMatch(csv, /Export mois précédent/);

    const expensesPage = await authedFetch(`http://127.0.0.1:${port}/expenses?month=2026-09`);
    assert.match(await expensesPage.text(), /href="\/export\/month\/2026\/9"/);
    const expensesExport = await authedFetch(`http://127.0.0.1:${port}/export/month/2026/9`);
    const exportPage = await authedFetch(`http://127.0.0.1:${port}/export/month?year=2026&month=9`);
    assert.equal(expensesExport.status, 200);
    assert.equal(exportPage.status, 200);
    assert.match(expensesExport.headers.get('content-type'), /spreadsheetml\.sheet/);
    assert.match(expensesExport.headers.get('content-disposition'), /DEPENSES_MAISON_2026-09\.xlsx/);
    const expensesWorkbook = await loadExcelJsWorkbook(Buffer.from(await expensesExport.arrayBuffer()));
    const exportPageWorkbook = await loadExcelJsWorkbook(Buffer.from(await exportPage.arrayBuffer()));
    assert.deepEqual(expensesWorkbook.worksheets.map((sheet) => sheet.name), ['Septembre']);
    const expensesSheet = expensesWorkbook.getWorksheet('Septembre');
    const exportPageSheet = exportPageWorkbook.getWorksheet('Septembre');
    assert.deepEqual(
      worksheetContentAndStyles(expensesSheet),
      worksheetContentAndStyles(exportPageSheet),
    );
    const renderedValues = Array.from({ length: expensesSheet.rowCount }, (_, rowIndex) => (
      Array.from({ length: expensesSheet.columnCount }, (_, columnIndex) => (
        expensesSheet.getCell(rowIndex + 1, columnIndex + 1).value
      ))
    )).flat();
    assert.equal(renderedValues.filter((value) => value === septemberTitles[0]).length, 1);
    assert.equal(renderedValues.filter((value) => value === septemberTitles[1]).length, 1);
    assert.ok(!renderedValues.includes(septemberTitles[2]));
    const expenseRow = findExcelColumnRow(expensesSheet, 5, septemberTitles[0]);
    assert.ok(expenseRow);
    const boutiqueRow = findExcelColumnRow(expensesSheet, 4, 'Compte boutique');
    assert.ok(boutiqueRow);
    assert.equal(expensesSheet.getCell(boutiqueRow, 4).value, 'Compte boutique');
    assert.equal(expensesSheet.getCell(boutiqueRow, 6).value, 1500);
    const totalRow = findExcelColumnRow(expensesSheet, 4, 'Total général');
    assert.ok(totalRow);
    const financeResponse = await authedFetch(`${baseUrl}/finance/monthly-data?month=2026-09`);
    const finance = await financeResponse.json();
    assert.equal(expensesSheet.getCell(totalRow, 6).value.result, finance.totalExpenses + finance.protectedSavingsExpenses);
  } finally {
    await db.prepare('DELETE FROM expenses WHERE title IN (?, ?, ?, ?, ?, ?)').run(...titles, ...septemberTitles);
    await db.prepare('DELETE FROM grocery_items WHERE id = ?').run(grocery.lastInsertRowid);
    await new Promise((resolve) => server.close(resolve));
  }
});

test('new expense categories appear in the filters and category summary on the next page load', async () => {
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;
  const categoryName = 'Catégorie dynamique dépenses';
  const updatedName = 'Catégorie dynamique modifiée';
  const exportTitles = [
    'Dépense catégorie dynamique export 1',
    'Dépense catégorie dynamique export 2',
  ];

  try {
    const createResponse = await authedFetch(`${baseUrl}/categories`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        name: categoryName,
        color: '#147967',
        return_to: '/expenses?month=2026-11',
      }),
      redirect: 'manual',
    });
    assert.equal(createResponse.status, 302);
    assert.equal(createResponse.headers.get('location'), '/expenses?month=2026-11');
    const category = await db.prepare('SELECT id FROM categories WHERE name = ?').get(categoryName);
    assert.ok(category);

    const expensesResponse = await authedFetch(`${baseUrl}/expenses?month=2026-11`);
    const expensesHtml = await expensesResponse.text();
    assert.match(expensesHtml, /data-modal-template="category-create"/);
    assert.match(expensesHtml, /href="\/categories\?return_to=%2Fexpenses%3Fmonth%3D2026-11">Modifier ou supprimer/);
    assert.match(expensesHtml, /name="return_to" value="\/expenses\?month=2026-11"/);
    assert.match(expensesHtml, new RegExp(`<option value="${category.id}"\\s*>${categoryName}</option>`));
    assert.match(expensesHtml, new RegExp(`<h3>${categoryName}</h3>`));
    assert.match(expensesHtml, /0\sFCFA/u);
    const emptyWorkbookResponse = await authedFetch(`${baseUrl}/export/month/2026/11`);
    const emptyWorkbook = await loadExcelJsWorkbook(Buffer.from(await emptyWorkbookResponse.arrayBuffer()));
    const emptySheet = emptyWorkbook.getWorksheet('Novembre');
    const emptyCategoryRow = findExcelColumnRow(emptySheet, 4, categoryName);
    assert.equal(emptyCategoryRow, null);

    await db.prepare(`
      INSERT INTO expenses (title, category_id, amount, expense_date, month_key)
      VALUES (?, ?, 2500, '2026-11-05', '2026-11'),
        (?, ?, 5000, '2026-11-12', '2026-11')
    `).run(exportTitles[0], category.id, exportTitles[1], category.id);
    const createdWorkbookResponse = await authedFetch(`${baseUrl}/export/month/2026/11`);
    const createdWorkbook = await loadExcelJsWorkbook(Buffer.from(await createdWorkbookResponse.arrayBuffer()));
    const createdSheet = createdWorkbook.getWorksheet('Novembre');
    const createdCategoryRow = findExcelColumnRow(createdSheet, 4, categoryName);
    assert.ok(createdCategoryRow);

    const managementHtml = await (await authedFetch(
      `${baseUrl}/categories?return_to=${encodeURIComponent('/expenses?month=2026-11')}`,
    )).text();
    assert.match(managementHtml, /data-modal-return="\/expenses\?month=2026-11"/);
    assert.match(managementHtml, /name="return_to" value="\/expenses\?month=2026-11"/);

    const updateResponse = await authedFetch(`${baseUrl}/categories/${category.id}/update`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        name: updatedName,
        color: '#147967',
        return_to: '/expenses?month=2026-11',
      }),
      redirect: 'manual',
    });
    assert.equal(updateResponse.status, 302);
    assert.equal(updateResponse.headers.get('location'), '/expenses?month=2026-11');
    const updatedHtml = await (await authedFetch(`${baseUrl}/expenses?month=2026-11`)).text();
    assert.match(updatedHtml, new RegExp(`<h3>${updatedName}</h3>`));
    assert.doesNotMatch(updatedHtml, new RegExp(`<h3>${categoryName}</h3>`));
    const renamedWorkbookResponse = await authedFetch(`${baseUrl}/export/month/2026/11`);
    const renamedWorkbook = await loadExcelJsWorkbook(Buffer.from(await renamedWorkbookResponse.arrayBuffer()));
    const renamedSheet = renamedWorkbook.getWorksheet('Novembre');
    assert.ok(findExcelColumnRow(renamedSheet, 4, updatedName));
    assert.equal(findExcelColumnRow(renamedSheet, 4, categoryName), null);

    const deleteResponse = await authedFetch(`${baseUrl}/categories/${category.id}/delete`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ return_to: '/expenses?month=2026-11' }),
      redirect: 'manual',
    });
    assert.equal(deleteResponse.status, 302);
    assert.equal(deleteResponse.headers.get('location'), '/expenses?month=2026-11');
    const deletedHtml = await (await authedFetch(`${baseUrl}/expenses?month=2026-11`)).text();
    assert.doesNotMatch(deletedHtml, new RegExp(`<h3>${updatedName}</h3>`));
    assert.match(deletedHtml, /Autres dépenses/);
    const deletedWorkbookResponse = await authedFetch(`${baseUrl}/export/month/2026/11`);
    const deletedWorkbook = await loadExcelJsWorkbook(Buffer.from(await deletedWorkbookResponse.arrayBuffer()));
    const deletedSheet = deletedWorkbook.getWorksheet('Novembre');
    assert.equal(findExcelColumnRow(deletedSheet, 4, updatedName), null);
    assert.equal(findExcelColumnRow(deletedSheet, 4, categoryName), null);
    const deletedValues = Array.from({ length: deletedSheet.rowCount }, (_, rowIndex) => (
      Array.from({ length: deletedSheet.columnCount }, (_, columnIndex) => (
        deletedSheet.getCell(rowIndex + 1, columnIndex + 1).value
      ))
    )).flat();
    assert.equal(exportTitles.reduce(
      (count, title) => count + deletedValues.filter((value) => value === title).length,
      0,
    ), 2);
  } finally {
    await db.prepare('DELETE FROM expenses WHERE title IN (?, ?)').run(...exportTitles);
    await db.prepare('DELETE FROM categories WHERE name IN (?, ?)').run(categoryName, updatedName);
    await new Promise((resolve) => server.close(resolve));
  }
});

test('deleting a default category reassigns its expenses and keeps it deleted after initialization', async () => {
  const categoryName = 'Factures';
  const renamedCategoryName = 'Factures test renommées';
  const fallback = await db.prepare("SELECT id FROM categories WHERE name = 'Autres dépenses'").get();
  const category = await db.prepare('SELECT id FROM categories WHERE name = ?').get(categoryName);
  assert.ok(category);
  const expense = await db.prepare(`
    INSERT INTO expenses (title, category_id, amount, expense_date)
    VALUES ('Dépense catégorie supprimée', ?, 1250, '2026-11-19')
  `).run(category.id);
  const server = await startServer();
  const { port } = server.address();
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const updateResponse = await authedFetch(`${baseUrl}/categories/${category.id}/update`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ name: renamedCategoryName, color: '#147967' }),
      redirect: 'manual',
    });
    assert.equal(updateResponse.status, 302);

    const response = await authedFetch(`http://127.0.0.1:${port}/categories/${category.id}/delete`, {
      method: 'POST',
      redirect: 'manual',
    });
    assert.equal(response.status, 302);
    assert.equal(
      Number((await db.prepare('SELECT category_id FROM expenses WHERE id = ?').get(expense.lastInsertRowid)).category_id),
      Number(fallback.id),
    );
    assert.equal(await db.prepare('SELECT id FROM categories WHERE name = ?').get(categoryName), undefined);

    await initializeDatabase();
    assert.equal(await db.prepare('SELECT id FROM categories WHERE name = ?').get(categoryName), undefined);
    assert.equal(await db.prepare('SELECT id FROM categories WHERE name = ?').get(renamedCategoryName), undefined);

    const recreated = await authedFetch(`http://127.0.0.1:${port}/categories`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ name: categoryName, color: '#147967' }),
      redirect: 'manual',
    });
    assert.equal(recreated.status, 302);
    await initializeDatabase();
    assert.ok(await db.prepare('SELECT id FROM categories WHERE name = ?').get(categoryName));
  } finally {
    await db.prepare('DELETE FROM expenses WHERE id = ?').run(expense.lastInsertRowid);
    await db.prepare('DELETE FROM categories WHERE name IN (?, ?)').run(categoryName, renamedCategoryName);
    await db.prepare('DELETE FROM category_seed_exclusions WHERE name IN (?, ?)').run(categoryName, renamedCategoryName);
    await initializeDatabase();
    await new Promise((resolve) => server.close(resolve));
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
    assert.ok(november.some((row) => row.includes('Ajout manuel après import')));
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});
