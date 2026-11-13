const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const XLSX = require('xlsx');
const Database = require('better-sqlite3');

const testDataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'app-depenses-app-test-'));
process.env.DB_PATH = path.join(testDataDir, 'app.db');
process.env.APP_DATA_DIR = path.join(testDataDir, 'data');
const { app, db } = require('../src/app');

test.after(() => {
  fs.rmSync(testDataDir, { recursive: true, force: true });
});

const startServer = async () => {
  const server = app.listen(0);
  await new Promise((resolve) => server.once('listening', resolve));
  return server;
};

test('GET / returns dashboard HTML', async () => {
  const server = await startServer();
  const { port } = server.address();

  try {
    const response = await fetch(`http://127.0.0.1:${port}/`);
    assert.equal(response.status, 200);
    const text = await response.text();
    assert.match(text, /Gestion des dépenses familiales/i);
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});

test('GET / renders the selected dashboard month and trend period', async () => {
  const server = await startServer();
  const { port } = server.address();

  try {
    const response = await fetch(`http://127.0.0.1:${port}/?month=2026-11&period=3`);
    assert.equal(response.status, 200);
    const text = await response.text();
    assert.match(text, /name="month" value="2026-11"/);
    assert.match(text, /Comparaison mensuelle des dépenses sur 3 mois/);
    assert.match(text, /Dépenses totales du mois/);
    assert.match(text, /Dernières opérations/);
  } finally {
    server.close();
  }
});

test('sidebar links open pages and page actions expose modal triggers', async () => {
  const server = await startServer();
  const { port } = server.address();

  try {
    for (const route of ['/expenses', '/groceries', '/history', '/monthly', '/import', '/export', '/backup', '/categories']) {
      const response = await fetch(`http://127.0.0.1:${port}${route}`);
      assert.equal(response.status, 200, `${route} should render as a page`);
      const text = await response.text();
      assert.match(text, /class="sidebar"/);
      assert.match(text, /href="\/history"/);
      assert.match(text, /src="\/app.js"/);
    }

    const expenses = await fetch(`http://127.0.0.1:${port}/expenses`);
    const expensesHtml = await expenses.text();
    assert.match(expensesHtml, /data-modal-template="expense-create"/);
    assert.doesNotMatch(expensesHtml, /href="\/groceries" data-modal-link/);
  } finally {
    server.close();
  }
});

test('POST /expenses stores an ordinary expense', async () => {
  const server = await startServer();
  const { port } = server.address();

  try {
    const response = await fetch(`http://127.0.0.1:${port}/expenses`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        title: 'Electricité',
        category_id: '1',
        amount: '25000',
        expense_date: '2026-11-13',
        description: 'Test',
        reference: 'facture-1',
      }).toString(),
    });

    assert.equal(response.status, 200);
  } finally {
    server.close();
  }
});

test('GET /export/2026 returns the 13-sheet annual workbook', async () => {
  const server = await startServer();
  const { port } = server.address();

  try {
    const response = await fetch(`http://127.0.0.1:${port}/export/2026`);
    assert.equal(response.status, 200);
    assert.match(response.headers.get('content-disposition'), /DEPENSES_MAISON_2026\.xlsx/);
    const workbook = XLSX.read(Buffer.from(await response.arrayBuffer()), { type: 'buffer' });
    assert.equal(workbook.SheetNames.length, 13);
    assert.equal(workbook.SheetNames[0], 'RÉCAPITULATIF ANNUEL');
    assert.ok(workbook.Sheets.JANVIER['!autofilter']);
  } finally {
    server.close();
  }
});

test('historical import backs up before one-time cleanup, filters dates and flags reimport duplicates', async () => {
  const categoryId = db.prepare('SELECT id FROM categories ORDER BY id LIMIT 1').get().id;
  db.prepare(`
    INSERT INTO expenses (title, category_id, amount, expense_date)
    VALUES ('Ancienne dépense de test', ?, 100, '2026-01-02')
  `).run(categoryId);
  db.prepare(`
    INSERT INTO grocery_items (item_name, purchase_date, quantity, unit, unit_price, total_amount)
    VALUES ('Ancien produit de test', '2026-01-03', 1, 'pièce', 200, 200)
  `).run();
  db.prepare(`
    INSERT INTO import_batches (file_name, rows_json, status)
    VALUES ('old-test.xlsx', '[]', 'imported')
  `).run();
  db.prepare('INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)')
    .run('setting_to_preserve', 'yes');

  const source = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(source, XLSX.utils.aoa_to_sheet([
    ['Date', 'Intitulé', 'Catégorie', 'Montant'],
    ['2026-01-06', 'Dépense catégorie nouvelle', 'Catégorie inconnue', 2000],
    ['2026-09-06', 'Dépense septembre interdite', 'Eau', 700],
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
    const uploadResponse = await fetch(`${baseUrl}/import`, { method: 'POST', body: uploadForm, redirect: 'manual' });
    assert.equal(uploadResponse.status, 302);
    const previewPath = uploadResponse.headers.get('location');
    const previewResponse = await fetch(`${baseUrl}${previewPath}`);
    const previewHtml = await previewResponse.text();
    assert.equal(previewResponse.status, 200);
    assert.match(previewHtml, /Catégories classées dans Autres/);
    assert.match(previewHtml, /Dépense septembre interdite/);
    assert.match(previewHtml, /Sauvegarder et remplacer l’historique/);

    const importResponse = await fetch(`${baseUrl}/import/confirm/${previewPath.split('/').at(-1)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ reinitialize: '1' }),
      redirect: 'manual',
    });
    assert.equal(importResponse.status, 302);
    assert.equal(db.prepare('SELECT COUNT(*) AS count FROM expenses').get().count, 1);
    assert.equal(db.prepare('SELECT COUNT(*) AS count FROM grocery_items').get().count, 1);
    assert.equal(db.prepare('SELECT COUNT(*) AS count FROM import_batches').get().count, 1);
    assert.equal(db.prepare("SELECT value FROM settings WHERE key = 'setting_to_preserve'").get().value, 'yes');
    assert.equal(db.prepare("SELECT COUNT(*) AS count FROM expenses WHERE title = 'Dépense septembre interdite'").get().count, 0);
    assert.equal(db.prepare("SELECT name FROM categories WHERE id = (SELECT category_id FROM expenses LIMIT 1)").get().name, 'Autres');
    assert.equal(db.prepare("SELECT value FROM settings WHERE key = 'historical_2026_reinitialized'").get().value.length > 0, true);

    const backupRecord = db.prepare("SELECT file_name FROM backup_runs WHERE source = 'pre-import-cleanup'").get();
    assert.ok(backupRecord);
    const backupPath = path.join(testDataDir, 'data', 'backups', backupRecord.file_name);
    const backupDb = new Database(backupPath, { readonly: true, fileMustExist: true });
    try {
      assert.equal(backupDb.pragma('integrity_check', { simple: true }), 'ok');
      assert.deepEqual(backupDb.pragma('foreign_key_check'), []);
      assert.equal(backupDb.prepare("SELECT COUNT(*) AS count FROM expenses WHERE title = 'Ancienne dépense de test'").get().count, 1);
    } finally {
      backupDb.close();
    }

    const duplicateForm = new FormData();
    duplicateForm.append('file', new Blob([fileBuffer]), 'history-again.xlsx');
    const duplicateUpload = await fetch(`${baseUrl}/import`, { method: 'POST', body: duplicateForm, redirect: 'manual' });
    const duplicatePreview = await fetch(`${baseUrl}${duplicateUpload.headers.get('location')}`);
    assert.match(await duplicatePreview.text(), /Doublons potentiels/);
    const duplicateConfirm = await fetch(`${baseUrl}${duplicateUpload.headers.get('location').replace('/preview/', '/confirm/')}`, {
      method: 'POST',
      redirect: 'manual',
    });
    assert.equal(duplicateConfirm.status, 400);
    assert.equal(db.prepare('SELECT COUNT(*) AS count FROM expenses').get().count, 1);

    await fetch(`${baseUrl}/expenses`, {
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
    const exportResponse = await fetch(`${baseUrl}/export/2026`);
    const exported = XLSX.read(Buffer.from(await exportResponse.arrayBuffer()), { type: 'buffer' });
    const november = XLSX.utils.sheet_to_json(exported.Sheets.NOVEMBRE, { header: 1, raw: true });
    assert.ok(november.some((row) => row[1] === 'Ajout manuel après import'));
  } finally {
    server.close();
  }
});
