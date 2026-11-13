require('dotenv').config();
const fs = require('fs');
const path = require('path');
const express = require('express');
const multer = require('multer');
const Database = require('better-sqlite3');
const XLSX = require('xlsx');
const { S3Client, PutObjectCommand } = require('@aws-sdk/client-s3');
const { db, initializeDatabase, dbPath } = require('./db');
const {
  buildAnnualWorkbook,
  END_DATE,
  markDuplicates,
  parseHistoricalWorkbook,
  START_DATE,
  summarizeRecords,
} = require('./excel');

const app = express();
const PORT = process.env.PORT || 3000;
const upload = multer({ storage: multer.memoryStorage() });
const appDataDir = process.env.APP_DATA_DIR || path.join(__dirname, '..', 'data');
const backupDir = path.join(appDataDir, 'backups');
const importDir = path.join(appDataDir, 'imports');

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, '..', 'views'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, '..', 'public')));

const r2Configured = Boolean(
  process.env.R2_BUCKET &&
  process.env.R2_ACCOUNT_ID &&
  process.env.R2_ACCESS_KEY_ID &&
  process.env.R2_SECRET_ACCESS_KEY
);

const ensureStorageDirectories = () => {
  [backupDir, importDir].forEach((dir) => {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  });
};

initializeDatabase();
ensureStorageDirectories();

const formatCurrency = (value) => {
  const amount = Number(value || 0);
  return new Intl.NumberFormat('fr-FR', {
    style: 'currency',
    currency: 'XAF',
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(amount);
};

const formatDate = (value) => {
  if (!value) return '-';
  const date = new Date(value + 'T00:00:00');
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleDateString('fr-FR');
};

const isModalRequest = (req) => (req && (req.query?.modal === '1' || req.query?.modal === 'true'));

const readCategories = () => db.prepare('SELECT * FROM categories ORDER BY name ASC').all();

const getMonthKey = (year, month) => `${year}-${String(month).padStart(2, '0')}`;

const getMonthlyTotals = (monthKey) => {
  const groceries = db.prepare(`
    SELECT COALESCE(SUM(total_amount), 0) AS total
    FROM grocery_items
    WHERE strftime('%Y-%m', purchase_date) = ?
  `).get(monthKey);

  const expenses = db.prepare(`
    SELECT COALESCE(SUM(amount), 0) AS total
    FROM expenses
    WHERE strftime('%Y-%m', expense_date) = ?
  `).get(monthKey);

  return {
    groceries: Number(groceries?.total || 0),
    expenses: Number(expenses?.total || 0),
  };
};

const currentDashboard = (monthKey, period) => {
  const [selectedYear, selectedMonth] = monthKey.split('-').map(Number);
  const selectedDate = new Date(selectedYear, selectedMonth - 1, 1);
  const previousMonthDate = new Date(selectedYear, selectedMonth - 2, 1);
  const previousMonth = getMonthKey(previousMonthDate.getFullYear(), previousMonthDate.getMonth() + 1);
  const startDate = new Date(selectedYear, selectedMonth - period, 1);
  const startMonth = getMonthKey(startDate.getFullYear(), startDate.getMonth() + 1);
  const monthlyRows = db.prepare(`
    SELECT month, SUM(groceries) AS groceries, SUM(expenses) AS expenses
    FROM (
      SELECT strftime('%Y-%m', purchase_date) AS month, SUM(total_amount) AS groceries, 0 AS expenses
      FROM grocery_items
      WHERE strftime('%Y-%m', purchase_date) BETWEEN ? AND ?
      GROUP BY month
      UNION ALL
      SELECT strftime('%Y-%m', expense_date) AS month, 0 AS groceries, SUM(amount) AS expenses
      FROM expenses
      WHERE strftime('%Y-%m', expense_date) BETWEEN ? AND ?
      GROUP BY month
    )
    GROUP BY month
  `).all(startMonth, monthKey, startMonth, monthKey);
  const totalsByMonth = new Map(monthlyRows.map((row) => [row.month, row]));
  const trend = Array.from({ length: period }, (_, index) => {
    const date = new Date(startDate.getFullYear(), startDate.getMonth() + index, 1);
    const key = getMonthKey(date.getFullYear(), date.getMonth() + 1);
    const totals = totalsByMonth.get(key) || { groceries: 0, expenses: 0 };
    return {
      key,
      label: date.toLocaleString('fr-FR', { month: 'short' }).replace('.', ''),
      groceries: Number(totals.groceries || 0),
      expenses: Number(totals.expenses || 0),
    };
  });

  const monthlyTotals = getMonthlyTotals(monthKey);
  const previousTotals = getMonthlyTotals(previousMonth);
  const monthlyOperationCount = Number(db.prepare(`
    SELECT
      (SELECT COUNT(*) FROM expenses WHERE strftime('%Y-%m', expense_date) = ?) +
      (SELECT COUNT(*) FROM grocery_items WHERE strftime('%Y-%m', purchase_date) = ?) AS count
  `).get(monthKey, monthKey).count);

  const recentOperations = db.prepare(`
    SELECT * FROM (
      SELECT e.id, e.title, e.amount, e.expense_date AS date, c.name AS category, 'expense' AS type
      FROM expenses e
      JOIN categories c ON c.id = e.category_id
      UNION ALL
      SELECT g.id, g.item_name AS title, g.total_amount AS amount, g.purchase_date AS date,
        'Compte boutique' AS category, 'grocery' AS type
      FROM grocery_items g
    )
    WHERE strftime('%Y-%m', date) = ?
    ORDER BY date DESC, id DESC
    LIMIT 6
  `).all(monthKey);

  const categoryBreakdown = db.prepare(`
    SELECT c.name, c.color, COALESCE(SUM(e.amount), 0) AS total
    FROM categories c
    LEFT JOIN expenses e ON e.category_id = c.id
      AND strftime('%Y-%m', e.expense_date) = ?
    GROUP BY c.id, c.name
    HAVING total > 0
    ORDER BY total DESC
    LIMIT 5
  `).all(monthKey);
  const categoryTotal = categoryBreakdown.reduce((sum, entry) => sum + Number(entry.total), 0);

  const groceryItems = db.prepare(`
    SELECT item_name, purchase_date, quantity, unit, total_amount
    FROM grocery_items
    WHERE strftime('%Y-%m', purchase_date) = ?
    ORDER BY purchase_date DESC, id DESC
    LIMIT 4
  `).all(monthKey);
  const groceryCount = Number(db.prepare(`
    SELECT COUNT(*) AS count FROM grocery_items WHERE strftime('%Y-%m', purchase_date) = ?
  `).get(monthKey).count);

  const latestYear = selectedDate.getFullYear();
  const boardTotal = monthlyTotals.groceries + monthlyTotals.expenses;

  return {
    monthKey,
    monthLabel: selectedDate.toLocaleString('fr-FR', { month: 'long', year: 'numeric' }),
    year: latestYear,
    period,
    currentMonthTotal: boardTotal,
    groceryTotal: monthlyTotals.groceries,
    otherExpenses: monthlyTotals.expenses,
    operationCount: monthlyOperationCount,
    monthDelta: boardTotal - (previousTotals.groceries + previousTotals.expenses),
    trend,
    trendMax: Math.max(...trend.map((item) => item.groceries + item.expenses), 1),
    categoryBreakdown,
    categoryTotal,
    recentOperations: recentOperations.map((item) => ({ ...item, amount: Number(item.amount) })),
    groceryItems,
    groceryCount,
  };
};

const fileList = (directory) => {
  if (!fs.existsSync(directory)) return [];
  return fs.readdirSync(directory)
    .sort((a, b) => fs.statSync(path.join(directory, b)).mtimeMs - fs.statSync(path.join(directory, a)).mtimeMs);
};

const getLatestImportBatch = () => db.prepare('SELECT * FROM import_batches ORDER BY id DESC LIMIT 1').get();

const hasInitialHistoricalImport = () => Boolean(
  db.prepare("SELECT value FROM settings WHERE key = 'historical_2026_reinitialized'").get(),
);

const existingHistoricalRows = () => [
  ...db.prepare(`
    SELECT title, expense_date AS date, amount, 'expense' AS type
    FROM expenses WHERE expense_date BETWEEN ? AND ?
  `).all(START_DATE, END_DATE),
  ...db.prepare(`
    SELECT item_name AS title, purchase_date AS date, total_amount AS amount, 'grocery' AS type
    FROM grocery_items WHERE purchase_date BETWEEN ? AND ?
  `).all(START_DATE, END_DATE),
];

const createVerifiedBackup = async (source) => {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const fileName = `${source}_${timestamp}.db`;
  const target = path.join(backupDir, fileName);
  await db.backup(target);

  const backupDb = new Database(target, { readonly: true, fileMustExist: true });
  let integrity;
  let foreignKeyViolations;
  try {
    integrity = backupDb.pragma('integrity_check', { simple: true });
    foreignKeyViolations = backupDb.pragma('foreign_key_check');
  } finally {
    backupDb.close();
  }
  const fileSize = fs.statSync(target).size;
  if (integrity !== 'ok' || foreignKeyViolations.length || fileSize === 0) {
    throw new Error(`La sauvegarde ${fileName} n'a pas passé les contrôles d'intégrité.`);
  }
  db.prepare(`
    INSERT INTO backup_runs (file_name, storage, status, metadata, source)
    VALUES (?, 'local', 'done', ?, ?)
  `).run(fileName, JSON.stringify({ path: target, integrity, foreignKeyViolations: 0, fileSize }), source);
  return { fileName, target, integrity, fileSize };
};

const ensureCategoryByName = (name) => {
  const trimmed = String(name || '').trim();
  if (!trimmed) return null;
  const existing = db.prepare('SELECT id FROM categories WHERE LOWER(name) = LOWER(?)').get(trimmed);
  if (existing) return existing.id;
  const color = ['#22c55e', '#38bdf8', '#a78bfa', '#f97316', '#facc15'][Math.floor(Math.random() * 5)];
  const result = db.prepare('INSERT INTO categories (name, color, type) VALUES (?, ?, ?)').run(trimmed, color, 'ordinary');
  return Number(result.lastInsertRowid);
};

const uploadBackupToR2 = async (filePath, fileName) => {
  if (!r2Configured) return null;

  const client = new S3Client({
    region: 'auto',
    endpoint: `https://${process.env.R2_ACCOUNT_ID}.r2.cloudflarestorage.com`,
    credentials: {
      accessKeyId: process.env.R2_ACCESS_KEY_ID,
      secretAccessKey: process.env.R2_SECRET_ACCESS_KEY,
    },
  });

  const fileBuffer = fs.readFileSync(filePath);
  await client.send(new PutObjectCommand({
    Bucket: process.env.R2_BUCKET,
    Key: `backups/${fileName}`,
    Body: fileBuffer,
    ContentType: 'application/octet-stream',
  }));

  return `https://${process.env.R2_BUCKET}.r2.cloudflarestorage.com/backups/${fileName}`;
};

app.use((req, res, next) => {
  res.locals.formatCurrency = formatCurrency;
  res.locals.formatDate = formatDate;
  res.locals.isModalRequest = isModalRequest(req);
  next();
});

app.get('/', (req, res) => {
  const today = new Date();
  const requestedMonth = (req.query.month || '').toString();
  const monthKey = /^\d{4}-(0[1-9]|1[0-2])$/.test(requestedMonth)
    ? requestedMonth
    : getMonthKey(today.getFullYear(), today.getMonth() + 1);
  const requestedPeriod = Number(req.query.period);
  const period = [3, 6, 12].includes(requestedPeriod) ? requestedPeriod : 6;
  const dashboard = currentDashboard(monthKey, period);
  res.render('index', { dashboard, categories: readCategories() });
});

app.get('/groceries', (req, res) => {
  const search = (req.query.search || '').toString().trim();
  let query = 'SELECT * FROM grocery_items ORDER BY purchase_date DESC, id DESC';
  const params = [];

  if (search) {
    query = 'SELECT * FROM grocery_items WHERE item_name LIKE ? OR unit LIKE ? OR notes LIKE ? ORDER BY purchase_date DESC, id DESC';
    params.push(`%${search}%`, `%${search}%`, `%${search}%`);
  }

  const groceryList = db.prepare(query).all(...params);
  res.render('groceries', { groceryList, search });
});

app.post('/groceries', (req, res) => {
  const { item_name, purchase_date, quantity, unit, unit_price, notes } = req.body;
  const q = Number(quantity || 0);
  const price = Number(unit_price || 0);
  const total = q * price;

  db.prepare(`
    INSERT INTO grocery_items (item_name, purchase_date, quantity, unit, unit_price, total_amount, notes)
    VALUES (?, ?, ?, ?, ?, ?, ?)
  `).run(item_name, purchase_date, q, unit || '', price, total, notes || '');

  res.redirect(req.body.return_to === '/history' ? '/history' : '/groceries');
});

app.post('/groceries/:id/delete', (req, res) => {
  db.prepare('DELETE FROM grocery_items WHERE id = ?').run(req.params.id);
  res.redirect(req.body.return_to === '/history' ? '/history' : '/groceries');
});

app.get('/expenses', (req, res) => {
  const search = (req.query.search || '').toString().trim();
  const categories = readCategories();
  const today = new Date();
  const requestedMonth = (req.query.month || '').toString();
  const monthKey = /^\d{4}-(0[1-9]|1[0-2])$/.test(requestedMonth)
    ? requestedMonth
    : getMonthKey(today.getFullYear(), today.getMonth() + 1);

  let query = `
    SELECT e.*, c.name AS category_name
    FROM expenses e
    JOIN categories c ON c.id = e.category_id
    WHERE strftime('%Y-%m', e.expense_date) = ?
  `;
  const params = [monthKey];

  if (search) {
    query += `
      AND (e.title LIKE ? OR e.description LIKE ? OR e.reference LIKE ? OR c.name LIKE ?)
    `;
    params.push(`%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`);
  }

  query += `
    ORDER BY e.expense_date DESC, e.id DESC
  `;

  const expenseList = db.prepare(query).all(...params);
  res.render('expenses', { categories, expenseList, search, monthKey });
});

app.post('/expenses', (req, res) => {
  const { title, category_id, amount, expense_date, description, reference } = req.body;
  db.prepare(`
    INSERT INTO expenses (title, category_id, amount, expense_date, description, reference)
    VALUES (?, ?, ?, ?, ?, ?)
  `).run(title, Number(category_id), Number(amount || 0), expense_date, description || '', reference || '');
  res.redirect(req.body.return_to === '/history' ? '/history' : '/expenses');
});

app.post('/expenses/:id/delete', (req, res) => {
  db.prepare('DELETE FROM expenses WHERE id = ?').run(req.params.id);
  res.redirect(req.body.return_to === '/history' ? '/history' : '/expenses');
});

app.get('/history', (req, res) => {
  const search = (req.query.search || '').toString().trim();
  let query = `
    SELECT * FROM (
      SELECT e.id, e.title, e.amount, e.expense_date AS date, c.name AS category, 'expense' AS type
      FROM expenses e
      JOIN categories c ON c.id = e.category_id
      UNION ALL
      SELECT g.id, g.item_name AS title, g.total_amount AS amount, g.purchase_date AS date,
        'Compte boutique' AS category, 'grocery' AS type
      FROM grocery_items g
    )
  `;
  const params = [];

  if (search) {
    query += ' WHERE title LIKE ? OR category LIKE ?';
    params.push(`%${search}%`, `%${search}%`);
  }

  query += ' ORDER BY date DESC, id DESC';
  const operationList = db.prepare(query).all(...params);
  res.render('history', { operationList, search });
});

app.get('/categories', (req, res) => {
  const categories = readCategories();
  res.render('categories', { categories });
});

app.post('/categories', (req, res) => {
  const { name, color } = req.body;
  if (!name) return res.redirect('/categories');
  db.prepare('INSERT INTO categories (name, color) VALUES (?, ?)').run(name.trim(), color || '#7c3aed');
  res.redirect('/categories');
});

app.post('/categories/:id/delete', (req, res) => {
  let fallback = db.prepare('SELECT id FROM categories WHERE name = ?').get('Autres');
  if (!fallback) {
    const inserted = db.prepare('INSERT INTO categories (name, color, type) VALUES (?, ?, ?)').run('Autres', '#94a3b8', 'ordinary');
    fallback = { id: Number(inserted.lastInsertRowid) };
  }

  if (!fallback || Number(req.params.id) === Number(fallback.id)) return res.redirect('/categories');

  db.prepare('UPDATE expenses SET category_id = ? WHERE category_id = ?').run(fallback.id, Number(req.params.id));

  db.prepare('DELETE FROM categories WHERE id = ?').run(req.params.id);
  res.redirect('/categories');
});

app.get('/monthly', (req, res) => {
  const today = new Date();
  const requestedYear = Number(req.query.year);
  const year = Number.isInteger(requestedYear) && requestedYear >= 2000 && requestedYear <= 2100
    ? requestedYear
    : today.getFullYear();
  const requestedMonth = Number(req.query.month);
  const month = Number.isInteger(requestedMonth) && requestedMonth >= 1 && requestedMonth <= 12
    ? requestedMonth
    : today.getMonth() + 1;
  const reportDate = getMonthKey(year, month);
  const monthLabel = new Date(Date.UTC(year, month - 1, 1))
    .toLocaleString('fr-FR', { month: 'long', year: 'numeric', timeZone: 'UTC' });

  const monthlyOther = db.prepare(`
    SELECT e.title, e.amount, e.expense_date, c.name AS category
    FROM expenses e
    JOIN categories c ON c.id = e.category_id
    WHERE strftime('%Y-%m', e.expense_date) = ?
    ORDER BY e.expense_date ASC
  `).all(reportDate);

  const monthlyGroceries = db.prepare(`
    SELECT *
    FROM grocery_items
    WHERE strftime('%Y-%m', purchase_date) = ?
    ORDER BY purchase_date ASC
  `).all(reportDate);

  const totalGroceries = monthlyGroceries.reduce((sum, row) => sum + Number(row.total_amount || 0), 0);
  const totalOther = monthlyOther.reduce((sum, row) => sum + Number(row.amount || 0), 0);

  res.render('monthly', {
    year,
    month,
    monthLabel,
    monthlyOther,
    monthlyGroceries,
    totalGroceries,
    totalOther,
    totalGeneral: totalGroceries + totalOther,
  });
});

app.get('/import', (req, res) => {
  const latestBatch = getLatestImportBatch();
  res.render('import', { latestBatch, initialImportComplete: hasInitialHistoricalImport(), r2Configured });
});

app.post('/import', upload.single('file'), (req, res) => {
  const file = req.file;
  if (!file) return res.redirect('/import');

  let parsed;
  try {
    parsed = parseHistoricalWorkbook(file.buffer, readCategories());
  } catch (error) {
    console.error('Excel import parse failed:', error.message);
    return res.status(400).send('Le fichier Excel ne peut pas être lu. Vérifiez le format puis réessayez.');
  }
  const rows = markDuplicates(parsed.records, existingHistoricalRows());
  const payload = JSON.stringify({
    ...parsed,
    records: rows,
    monthly: summarizeRecords(rows.map((row) => ({
      ...row,
      duplicateSource: row.duplicateSource === 'fichier' ? 'fichier' : null,
    }))),
  });
  const result = db.prepare('INSERT INTO import_batches (file_name, rows_json, status) VALUES (?, ?, ?)')
    .run(path.basename(file.originalname), payload, 'preview');

  fs.writeFileSync(path.join(importDir, `${Date.now()}_${path.basename(file.originalname)}`), file.buffer);
  res.redirect(`/import/preview/${result.lastInsertRowid}`);
});

app.get('/import/preview/:id', (req, res) => {
  const previewBatch = db.prepare('SELECT * FROM import_batches WHERE id = ?').get(req.params.id);
  if (!previewBatch) return res.redirect('/import');

  const payload = JSON.parse(previewBatch.rows_json || '{}');
  const records = payload.records || [];
  const duplicates = records.filter((row) => row.duplicateSource);
  const validCount = records.filter((row) => row.valid).length;
  const reinitializeAvailable = !hasInitialHistoricalImport() && validCount > 0;
  const monthly = payload.monthly || summarizeRecords(records);

  res.render('import-preview', {
    previewBatch,
    records,
    duplicates,
    monthly,
    excludedRows: payload.excludedRows || [],
    ignoredRows: payload.ignoredRows || [],
    summaryLines: payload.summaryLines || [],
    sheetNames: payload.sheetNames || [],
    reinitializeAvailable,
    initialImportComplete: hasInitialHistoricalImport(),
  });
});

app.post('/import/confirm/:id', async (req, res) => {
  const previewBatch = db.prepare('SELECT * FROM import_batches WHERE id = ?').get(req.params.id);
  if (!previewBatch) return res.redirect('/import');

  const payload = JSON.parse(previewBatch.rows_json || '{}');
  const reinitialize = req.body?.reinitialize === '1';
  const alreadyInitialized = hasInitialHistoricalImport();
  if (!alreadyInitialized && !reinitialize) {
    return res.status(409).send('La première importation 2026 doit être confirmée avec sauvegarde et nettoyage.');
  }
  if (alreadyInitialized && reinitialize) {
    return res.status(409).send('Le nettoyage initial a déjà été effectué. Il ne peut pas être répété.');
  }

  const sourceRecords = (payload.records || []).filter((row) => row.valid);
  const unresolvedBoutiqueSummaries = (payload.summaryLines || [])
    .filter((summary) => summary.unresolved || summary.mismatch);
  if (unresolvedBoutiqueSummaries.length) {
    return res.status(400).send('Importation annulée : un total boutique ne correspond pas à ses marchandises détaillées.');
  }
  const records = markDuplicates(
    sourceRecords,
    reinitialize ? [] : existingHistoricalRows(),
  );
  const toImport = records.filter((row) => !row.duplicateSource);
  if (!toImport.length) {
    return res.status(400).send('Aucune ligne valide et non dupliquée à importer pour la période demandée.');
  }

  let backup;
  if (reinitialize) {
    try {
      backup = await createVerifiedBackup('pre-import-cleanup');
    } catch (error) {
      console.error('Pre-import backup failed:', error.message);
      return res.status(500).send(`Nettoyage annulé : sauvegarde non vérifiée (${error.message})`);
    }
  }

  const transaction = db.transaction(() => {
    if (reinitialize) {
      if (hasInitialHistoricalImport()) throw new Error('Le nettoyage initial a déjà été effectué.');
      db.prepare('DELETE FROM expenses').run();
      db.prepare('DELETE FROM grocery_items').run();
      db.prepare('DELETE FROM import_batches').run();
    }

    const expenseIds = [];
    const groceryIds = [];
    toImport.forEach((row) => {
      if (row.type === 'grocery') {
        const result = db.prepare(`
          INSERT INTO grocery_items (
            item_name, purchase_date, quantity, unit, unit_price, total_amount, notes,
            quantity_known, unit_price_known
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        `).run(
          row.title,
          row.date,
          row.quantityKnown ? row.quantity : 0,
          row.unit || '',
          row.unitPriceKnown ? row.unitPrice : 0,
          row.amount,
          row.notes || '',
          row.quantityKnown ? 1 : 0,
          row.unitPriceKnown ? 1 : 0,
        );
        groceryIds.push(Number(result.lastInsertRowid));
        return;
      }

      const categoryId = ensureCategoryByName(row.categoryName || 'Autres');
      const result = db.prepare(`
        INSERT INTO expenses (title, category_id, amount, expense_date, description, reference)
        VALUES (?, ?, ?, ?, ?, ?)
      `).run(row.title, categoryId, row.amount, row.date, row.notes || '', row.reference || '');
      expenseIds.push(Number(result.lastInsertRowid));
    });

    const expectedExpenses = toImport.filter((row) => row.type === 'expense');
    const expectedGroceries = toImport.filter((row) => row.type === 'grocery');
    const storedExpenses = expenseIds.map((id) => db.prepare(`
      SELECT e.title, e.expense_date AS date, e.amount, c.name AS category
      FROM expenses e LEFT JOIN categories c ON c.id = e.category_id WHERE e.id = ?
    `).get(id));
    const storedGroceries = groceryIds.map((id) => db.prepare(`
      SELECT item_name AS title, purchase_date AS date, total_amount AS amount,
        quantity, unit, unit_price AS unitPrice, quantity_known AS quantityKnown,
        unit_price_known AS unitPriceKnown
      FROM grocery_items WHERE id = ?
    `).get(id));
    const expenseTotal = storedExpenses.reduce((sum, row) => sum + Number(row.amount), 0);
    const groceryTotal = storedGroceries.reduce((sum, row) => sum + Number(row.amount), 0);
    const rowsMatch = (expected, stored, type) => expected.length === stored.length
      && expected.every((row, index) => {
        const actual = stored[index];
        const titleMatches = row.title === actual.title;
        const dateMatches = row.date === actual.date;
        const amountMatches = Math.abs(row.amount - Number(actual.amount)) <= 0.001;
        const categoryMatches = type !== 'expense' || row.categoryName === actual.category;
        const quantityMatches = type !== 'grocery'
          || (row.quantityKnown === Boolean(actual.quantityKnown)
            && row.unitPriceKnown === Boolean(actual.unitPriceKnown)
            && (!row.quantityKnown || row.quantity === actual.quantity)
            && (!row.unitPriceKnown || row.unitPrice === actual.unitPrice)
            && (row.unit || '') === (actual.unit || ''));
        return titleMatches && dateMatches && amountMatches && categoryMatches && quantityMatches;
      });
    if (!rowsMatch(expectedExpenses, storedExpenses, 'expense')
      || !rowsMatch(expectedGroceries, storedGroceries, 'grocery')
      || Math.abs(expenseTotal - expectedExpenses.reduce((sum, row) => sum + row.amount, 0)) > 0.001
      || Math.abs(groceryTotal - expectedGroceries.reduce((sum, row) => sum + row.amount, 0)) > 0.001) {
      throw new Error('Les contrôles après insertion ne correspondent pas aux lignes prévisualisées.');
    }
    const expectedMonthly = summarizeRecords(toImport);
    const actualMonthly = new Map();
    [...storedExpenses.map((row) => ({ ...row, type: 'expense' })),
      ...storedGroceries.map((row) => ({ ...row, type: 'grocery' }))].forEach((row) => {
      const month = actualMonthly.get(row.date.slice(0, 7)) || {
        ordinaryCount: 0, groceryCount: 0, totalOther: 0, totalBoutique: 0,
      };
      if (row.type === 'expense') {
        month.ordinaryCount += 1;
        month.totalOther += Number(row.amount);
      } else {
        month.groceryCount += 1;
        month.totalBoutique += Number(row.amount);
      }
      actualMonthly.set(row.date.slice(0, 7), month);
    });
    expectedMonthly.forEach((month) => {
      const actual = actualMonthly.get(month.month) || {
        ordinaryCount: 0, groceryCount: 0, totalOther: 0, totalBoutique: 0,
      };
      if (actual.ordinaryCount !== month.ordinaryCount
        || actual.groceryCount !== month.groceryCount
        || Math.abs(actual.totalOther - month.totalOther) > 0.001
        || Math.abs(actual.totalBoutique - month.totalBoutique) > 0.001
        || Math.abs(month.totalGeneral - (month.totalBoutique + month.totalOther)) > 0.001) {
        throw new Error(`Le contrôle des totaux de ${month.label} ne correspond pas à la prévisualisation.`);
      }
    });

    if (reinitialize) {
      db.prepare(`
        INSERT INTO import_batches (file_name, rows_json, status) VALUES (?, ?, 'imported')
      `).run(previewBatch.file_name, JSON.stringify({ ...payload, records, backup: backup.fileName }));
      db.prepare(`
        INSERT OR REPLACE INTO settings (key, value) VALUES ('historical_2026_reinitialized', ?)
      `).run(new Date().toISOString());
    } else {
      db.prepare('UPDATE import_batches SET rows_json = ?, status = ? WHERE id = ?')
        .run(JSON.stringify({ ...payload, records }), 'imported', previewBatch.id);
    }
  });

  try {
    transaction();
  } catch (error) {
    console.error('Excel import transaction failed:', error.message);
    return res.status(500).send(`Importation annulée, aucune donnée n’a été modifiée : ${error.message}`);
  }
  res.redirect('/import');
});

app.get('/backup', (req, res) => {
  const backups = fileList(backupDir).filter((file) => /\.(db|sqlite|bak|sql)$/i.test(file));
  const latestBackup = backups[0] || null;
  res.render('backup', { backups, latestBackup, r2Configured });
});

app.get('/export', (req, res) => {
  const year = Number(req.query.year) || new Date().getFullYear();
  res.render('export', { year });
});

app.post('/backup/create', async (req, res) => {
  let backup;
  try {
    backup = await createVerifiedBackup('manual');
  } catch (error) {
    console.error('Database backup failed:', error.message);
    return res.status(500).send(`La sauvegarde n’a pas pu être vérifiée : ${error.message}`);
  }

  let r2Url = null;
  try {
    r2Url = await uploadBackupToR2(backup.target, backup.fileName);
  } catch (error) {
    console.error('Cloudflare R2 backup failed:', error.message);
  }

  db.prepare('UPDATE backup_runs SET storage = ?, metadata = ? WHERE file_name = ?')
    .run(r2Url ? 'cloudflare-r2' : 'local', JSON.stringify({
      path: backup.target,
      url: r2Url,
      integrity: backup.integrity,
      fileSize: backup.fileSize,
    }), backup.fileName);

  res.redirect('/backup');
});

app.post('/backup/restore/:fileName', (req, res) => {
  const fileName = decodeURIComponent(req.params.fileName);
  const source = path.join(backupDir, fileName);
  if (!fs.existsSync(source)) {
    return res.status(404).send('Sauvegarde introuvable');
  }

  fs.copyFileSync(source, dbPath);
  res.redirect('/backup');
});

app.get('/export/:year', (req, res) => {
  const year = Number(req.params.year);
  if (!Number.isInteger(year) || year < 2000 || year > 2100) {
    return res.status(400).send('Année d’exportation invalide.');
  }
  const startDate = `${year}-01-01`;
  const endDate = `${year + 1}-01-01`;
  const expenses = db.prepare(`
    SELECT e.title, e.amount, e.expense_date, COALESCE(c.name, 'Autres') AS category
    FROM expenses e
    LEFT JOIN categories c ON c.id = e.category_id
    WHERE e.expense_date >= ? AND e.expense_date < ?
    ORDER BY e.expense_date ASC, e.id ASC
  `).all(startDate, endDate);
  const groceryItems = db.prepare(`
    SELECT item_name, purchase_date, quantity, unit, unit_price, total_amount,
      quantity_known AS quantityKnown, unit_price_known AS unitPriceKnown
    FROM grocery_items
    WHERE purchase_date >= ? AND purchase_date < ?
    ORDER BY purchase_date ASC, id ASC
  `).all(startDate, endDate);
  const { workbook } = buildAnnualWorkbook(year, expenses, groceryItems);
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const fileName = `DEPENSES_MAISON_${year}.xlsx`;
  res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
});

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`App running on http://localhost:${PORT}`);
  });
}

module.exports = { app, db, initializeDatabase, formatCurrency, formatDate };
