require('dotenv').config();
const fs = require('fs');
const path = require('path');
const { randomUUID } = require('node:crypto');
const express = require('express');
const multer = require('multer');
const { promisify } = require('node:util');
const { execFile } = require('node:child_process');
const XLSX = require('xlsx-js-style');
const PDFDocument = require('pdfkit');
const { S3Client, PutObjectCommand } = require('@aws-sdk/client-s3');
const { db, initializeDatabase } = require('./db');
const { router: financeRouter, getMonthlyFinance, writeAudit } = require('./finance');
const {
  clearFailedLogins,
  clearSessionCookie,
  createSession,
  deleteSession,
  hashPassword,
  loginIsLimited,
  noteFailedLogin,
  resolveSession,
  safeRedirect,
  setSessionCookie,
  verifyPassword,
} = require('./auth');
const {
  buildAnnualWorkbook,
  buildMonthlyWorkbook,
  END_DATE,
  markDuplicates,
  parseAmount,
  parseHistoricalFile,
  START_DATE,
  summarizeRecords,
} = require('./excel');
const { buildReferenceWorkbook, MONTH_NAMES } = require('./template-export');
const {
  buildCategoryPresentation,
  filterCategoryPresentation,
  resolvePriorityCategory,
  sortCategoriesForPresentation,
} = require('./category-presentation');

const app = express();
const PORT = process.env.PORT || 3000;
const execFileAsync = promisify(execFile);
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024, files: 1 },
  fileFilter(req, file, callback) {
    if (!/\.(xlsx|xls|csv)$/i.test(file.originalname)) {
      return callback(new Error('Formats acceptés : XLSX, XLS ou CSV.'));
    }
    callback(null, true);
  },
});
const importFileUpload = (req, res, next) => upload.single('file')(req, res, (error) => {
  if (error) return res.status(400).send(`Fichier invalide : ${error.message}`);
  next();
});
const receiptUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024, files: 1 },
  fileFilter(req, file, callback) {
    const allowed = new Set(['application/pdf', 'image/jpeg', 'image/png', 'image/webp']);
    if (!allowed.has(file.mimetype)) return callback(new Error('Justificatif accepté : PDF, JPEG, PNG ou WEBP.'));
    callback(null, true);
  },
});
const appDataDir = process.env.APP_DATA_DIR || path.join(__dirname, '..', 'data');
const backupDir = path.join(appDataDir, 'backups');
const importDir = path.join(appDataDir, 'imports');
const receiptDir = path.join(appDataDir, 'receipts');
const normalizeIncomeSource = (value) => String(value || '')
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g, '')
  .toLocaleLowerCase('fr')
  .trim();
const isPrimaryIncomeSource = (value) => ['pape', 'pension pere'].includes(normalizeIncomeSource(value));
const isRentIncomeSource = (value) => /loyer|location|locataire/i.test(normalizeIncomeSource(value));

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
  [backupDir, importDir, receiptDir].forEach((dir) => {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  });
};

const initialization = initializeDatabase();
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
const isValidIsoDate = (value) => {
  const date = String(value || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return false;
  const parsed = new Date(`${date}T00:00:00.000Z`);
  return !Number.isNaN(parsed.getTime()) && parsed.toISOString().slice(0, 10) === date;
};
const csvCell = (value) => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') return String(value);
  let text = String(value);
  if (/^[=+\-@\t\r]/.test(text)) text = `'${text}`;
  return `"${text.replace(/"/g, '""')}"`;
};
const makeCsv = (headers, rows) => `\uFEFF${[headers, ...rows]
  .map((row) => row.map(csvCell).join(';'))
  .join('\r\n')}`;

const getMonthlyTotals = async (monthKey) => {
  const groceries = await db.prepare(`
    SELECT COALESCE(SUM(total_amount), 0) AS total
    FROM grocery_items
    WHERE COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM')) = ?
  `).get(monthKey);

  const expenses = await db.prepare(`
    SELECT COALESCE(SUM(amount), 0) AS total
    FROM expenses
    WHERE funding_source <> 'protected_savings'
      AND COALESCE(NULLIF(month_key, ''), to_char(expense_date, 'YYYY-MM')) = ?
  `).get(monthKey);

  return {
    groceries: Number(groceries?.total || 0),
    expenses: Number(expenses?.total || 0),
  };
};

const currentDashboard = async (monthKey, period) => {
  const [selectedYear, selectedMonth] = monthKey.split('-').map(Number);
  const selectedDate = new Date(selectedYear, selectedMonth - 1, 1);
  const previousMonthDate = new Date(selectedYear, selectedMonth - 2, 1);
  const previousMonth = getMonthKey(previousMonthDate.getFullYear(), previousMonthDate.getMonth() + 1);
  const startDate = new Date(selectedYear, selectedMonth - period, 1);
  const startMonth = getMonthKey(startDate.getFullYear(), startDate.getMonth() + 1);
  const monthlyRows = await db.prepare(`
    SELECT month, SUM(groceries) AS groceries, SUM(expenses) AS expenses
    FROM (
      SELECT COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM')) AS month,
        SUM(total_amount) AS groceries, 0 AS expenses
      FROM grocery_items
      WHERE COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM')) BETWEEN ? AND ?
      GROUP BY month
      UNION ALL
      SELECT COALESCE(NULLIF(month_key, ''), to_char(expense_date, 'YYYY-MM')) AS month,
        0 AS groceries, SUM(amount) AS expenses
      FROM expenses
      WHERE funding_source <> 'protected_savings'
        AND COALESCE(NULLIF(month_key, ''), to_char(expense_date, 'YYYY-MM')) BETWEEN ? AND ?
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
  const financialTrend = await Promise.all(trend.map(async (item) => {
    const finance = await getMonthlyFinance(item.key);
    return {
      key: item.key,
      label: item.label,
      income: finance.incomeTotal,
      expenses: finance.totalExpenses,
      closingBalance: finance.closingBalance,
    };
  }));
  const financialTrendMax = Math.max(
    ...financialTrend.flatMap((item) => [item.income, item.expenses]),
    1,
  );
  const balanceTrendMaxAbs = Math.max(...financialTrend.map((item) => Math.abs(item.closingBalance)), 1);

  const monthlyTotals = await getMonthlyTotals(monthKey);
  const previousTotals = await getMonthlyTotals(previousMonth);
  const monthlyOperationCount = Number((await db.prepare(`
    SELECT
      (SELECT COUNT(*) FROM expenses WHERE COALESCE(NULLIF(month_key, ''), to_char(expense_date, 'YYYY-MM')) = ?) +
      (SELECT COUNT(*) FROM grocery_items WHERE COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM')) = ?) AS count
  `).get(monthKey, monthKey)).count);

  const recentOperations = await db.prepare(`
    SELECT * FROM (
      SELECT e.id,
        CASE WHEN e.funding_source = 'protected_savings'
          THEN (SELECT s.id FROM savings_transactions s WHERE s.related_expense_id = e.id ORDER BY s.id DESC LIMIT 1)
          ELSE e.id
        END AS source_id,
        e.title, e.amount, e.expense_date AS date, e.funding_source,
        COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) AS month,
        c.name AS category, 'expense' AS type, e.beneficiary,
        CONCAT_WS(' ', e.title, c.name, e.beneficiary, e.paid_by, e.funding_source) AS search_text
      FROM expenses e
      JOIN categories c ON c.id = e.category_id
      UNION ALL
      SELECT g.id, g.id AS source_id, g.item_name AS title, g.total_amount AS amount, g.purchase_date AS date,
        'current' AS funding_source,
        COALESCE(NULLIF(g.month_key, ''), to_char(g.purchase_date, 'YYYY-MM')) AS month,
        'Compte boutique' AS category, 'grocery' AS type, NULL AS beneficiary,
        CONCAT_WS(' ', g.item_name, g.notes, 'Compte boutique') AS search_text
      FROM grocery_items g
    )
    WHERE month = ?
    ORDER BY date DESC, id DESC
    LIMIT 6
  `).all(monthKey);

  const categoryBreakdown = await db.prepare(`
    SELECT c.name, c.color, COALESCE(SUM(e.amount), 0) AS total
    FROM categories c
    LEFT JOIN expenses e ON e.category_id = c.id
      AND e.funding_source <> 'protected_savings'
      AND COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) = ?
    GROUP BY c.id, c.name
    HAVING SUM(e.amount) > 0
    ORDER BY total DESC
    LIMIT 5
  `).all(monthKey);
  const categoryTotal = categoryBreakdown.reduce((sum, entry) => sum + Number(entry.total), 0);

  const groceryItems = await db.prepare(`
    SELECT item_name, purchase_date, month_key, quantity, unit, total_amount
    FROM grocery_items
    WHERE COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM')) = ?
    ORDER BY purchase_date DESC NULLS LAST, id DESC
    LIMIT 4
  `).all(monthKey);
  const groceryCount = Number((await db.prepare(`
    SELECT COUNT(*) AS count FROM grocery_items
    WHERE COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM')) = ?
  `).get(monthKey)).count);

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
    financialTrend,
    financialTrendMax,
    balanceTrendMaxAbs,
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

const hasInitialHistoricalImport = async () => Boolean(
  await db.prepare("SELECT value FROM settings WHERE key = 'historical_2026_reinitialized'").get(),
);

const existingHistoricalRows = async () => [
  ...await db.prepare(`
    SELECT e.title, e.expense_date AS date, e.amount, e.month_key AS "monthKey",
      c.name AS category, 'expense' AS type
    FROM expenses e LEFT JOIN categories c ON c.id = e.category_id
    WHERE e.expense_date BETWEEN ? AND ?
      OR (e.expense_date IS NULL AND e.month_key BETWEEN ? AND ?)
  `).all(START_DATE, END_DATE, START_DATE.slice(0, 7), END_DATE.slice(0, 7)),
  ...await db.prepare(`
    SELECT item_name AS title, purchase_date AS date, month_key AS "monthKey",
      total_amount AS amount, 'grocery' AS type
    FROM grocery_items
    WHERE purchase_date BETWEEN ? AND ?
      OR (purchase_date IS NULL AND month_key BETWEEN ? AND ?)
  `).all(START_DATE, END_DATE, START_DATE.slice(0, 7), END_DATE.slice(0, 7)),
];

const createVerifiedBackup = async (source) => {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const fileName = `${source}_${timestamp}.dump`;
  const target = path.join(backupDir, fileName);
  const databaseUrl = process.env.DATABASE_URL;
  const args = ['--format=custom', '--no-owner', '--file', target];
  if (databaseUrl) args.push(databaseUrl);
  else {
    args.push('--host', process.env.PGHOST || '/var/run/postgresql');
    if (process.env.PGPORT) args.push('--port', process.env.PGPORT);
    args.push('--username', process.env.PGUSER || process.env.USER);
    args.push('--dbname', process.env.PGDATABASE || 'app_depenses');
  }
  await execFileAsync('pg_dump', args);
  const { stdout } = await execFileAsync('pg_restore', ['--list', target]);
  const tableNames = [
    'categories', 'expenses', 'grocery_items', 'import_batches', 'backup_runs', 'settings',
    'units', 'tenants', 'rent_obligations', 'rent_payments', 'income_categories', 'incomes',
    'savings_accounts', 'savings_transactions', 'savings_balance_snapshots',
  ];
  const sourceCounts = [];
  for (const name of tableNames) {
    const result = await db.prepare(`SELECT COUNT(*) AS count FROM ${name}`).get();
    sourceCounts.push([name, Number(result.count)]);
  }
  const tableDefinitionsPresent = tableNames.every((name) => stdout.includes(`TABLE public ${name} `));
  const tableDataPresent = sourceCounts.every(([name, count]) => count === 0
    || stdout.includes(`TABLE DATA public ${name} `));
  const archiveVerified = tableDefinitionsPresent && tableDataPresent;
  const fileSize = fs.statSync(target).size;
  if (!archiveVerified || fileSize === 0) {
    throw new Error(`La sauvegarde ${fileName} n'a pas passé les contrôles d'intégrité.`);
  }
  await db.prepare(`
    INSERT INTO backup_runs (file_name, storage, status, metadata, source)
    VALUES (?, 'local', 'done', ?, ?)
  `).run(fileName, JSON.stringify({
    path: target,
    integrity: 'pg_restore --list',
    archiveVerified,
    tablesVerified: tableNames.length,
    rowCounts: Object.fromEntries(sourceCounts),
    fileSize,
  }), source);
  return { fileName, target, integrity: 'pg_restore --list', fileSize };
};

const ensureCategoryByName = async (name) => {
  const trimmed = String(name || '').trim();
  if (!trimmed) return null;
  const existing = await db.prepare('SELECT id FROM categories WHERE LOWER(name) = LOWER(?)').get(trimmed);
  if (existing) return existing.id;
  const color = ['#22c55e', '#38bdf8', '#a78bfa', '#f97316', '#facc15'][Math.floor(Math.random() * 5)];
  const result = await db.prepare('INSERT INTO categories (name, color, type) VALUES (?, ?, ?)').run(trimmed, color, 'ordinary');
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

app.use(async (req, res, next) => {
  try {
    await initialization;
  } catch (error) {
    console.error('PostgreSQL initialization failed:', error.message);
    return res.status(500).send('La base de données PostgreSQL est indisponible.');
  }
  res.locals.formatCurrency = formatCurrency;
  res.locals.formatDate = formatDate;
  res.locals.isModalRequest = isModalRequest(req);
  next();
});

app.get('/login', async (req, res) => {
  res.setHeader('Cache-Control', 'no-store, private');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  const next = safeRedirect(req.query.next);
  if (await resolveSession(req)) return res.redirect(next);
  res.render('login', {
    error: null,
    notice: req.query.loggedOut === '1' ? 'Vous êtes déconnecté(e).' : null,
    identifier: '',
    next,
  });
});

app.post('/login', async (req, res) => {
  res.setHeader('Cache-Control', 'no-store, private');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  const ip = req.ip || req.socket.remoteAddress || 'unknown';
  if (loginIsLimited(ip)) {
    return res.status(429).render('login', {
      error: 'Trop de tentatives. Réessayez dans 15 minutes.',
      notice: null,
      identifier: String(req.body.identifier || req.body.email || ''),
      next: safeRedirect(req.body.next),
    });
  }
  const identifier = String(req.body.identifier || req.body.email || '').trim().toLowerCase();
  const user = await db.prepare(`
    SELECT id, password_hash FROM users
    WHERE is_active = TRUE AND (LOWER(email) = ? OR id::text = ?)
  `).get(identifier, identifier);
  const valid = user && await verifyPassword(String(req.body.password || ''), user.password_hash);
  if (!valid) {
    noteFailedLogin(ip);
    return res.status(401).render('login', {
      error: 'Adresse ou mot de passe incorrect.',
      notice: null,
      identifier: String(req.body.identifier || req.body.email || ''),
      next: safeRedirect(req.body.next),
    });
  }
  clearFailedLogins(ip);
  setSessionCookie(res, await createSession(user.id));
  res.redirect(safeRedirect(req.body.next));
});

app.post('/logout', async (req, res) => {
  res.setHeader('Cache-Control', 'no-store, private');
  await deleteSession(req);
  clearSessionCookie(res);
  res.redirect('/login?loggedOut=1');
});

app.use(async (req, res, next) => {
  if (req.path === '/login' && ['GET', 'POST'].includes(req.method)) return next();
  res.setHeader('Cache-Control', 'no-store, private');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  try {
    const user = await resolveSession(req);
    if (!user) return res.redirect(`/login?next=${encodeURIComponent(req.originalUrl)}`);
    req.user = user;
    res.locals.user = user;
    next();
  } catch (error) {
    next(error);
  }
});

app.get('/account', (req, res) => {
  res.render('account', {
    email: req.user.email,
    emailError: null,
    passwordError: null,
    notice: req.query.updated === 'email'
      ? 'Adresse e-mail mise à jour.'
      : req.query.updated === 'password'
        ? 'Mot de passe modifié. Les autres sessions ont été déconnectées.'
        : null,
  });
});

app.post('/account/email', async (req, res) => {
  const email = String(req.body.email || '').trim().toLowerCase();
  const currentPassword = String(req.body.current_password || '');
  const renderError = (message, status = 400) => res.status(status).render('account', {
    email,
    emailError: message,
    passwordError: null,
    notice: null,
  });

  if (email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return renderError('Saisissez une adresse e-mail valide.');
  }
  const account = await db.prepare(
    'SELECT password_hash FROM users WHERE id = ? AND is_active = TRUE',
  ).get(req.user.id);
  if (!account || !await verifyPassword(currentPassword, account.password_hash)) {
    return renderError('Le mot de passe actuel est incorrect.', 401);
  }
  if (email === req.user.email.toLowerCase()) {
    return res.redirect('/account?updated=email');
  }
  const existing = await db.prepare(
    'SELECT id FROM users WHERE LOWER(email) = ? AND id <> ?',
  ).get(email, req.user.id);
  if (existing) return renderError('Cette adresse e-mail est déjà utilisée.', 409);

  try {
    await db.prepare(
      'UPDATE users SET email = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
    ).run(email, req.user.id);
  } catch (error) {
    if (error.code === '23505') {
      return renderError('Cette adresse e-mail est déjà utilisée.', 409);
    }
    throw error;
  }
  res.redirect('/account?updated=email');
});

app.post('/account/password', async (req, res) => {
  const email = req.user.email;
  const currentPassword = String(req.body.current_password || '');
  const newPassword = String(req.body.new_password || '');
  const confirmation = String(req.body.confirm_password || '');
  const renderError = (message, status = 400) => res.status(status).render('account', {
    email,
    emailError: null,
    passwordError: message,
    notice: null,
  });

  if (newPassword.length < 12 || newPassword.length > 128) {
    return renderError('Le nouveau mot de passe doit contenir entre 12 et 128 caractères.');
  }
  if (newPassword !== confirmation) {
    return renderError('La confirmation du nouveau mot de passe ne correspond pas.');
  }
  const account = await db.prepare(
    'SELECT password_hash FROM users WHERE id = ? AND is_active = TRUE',
  ).get(req.user.id);
  if (!account || !await verifyPassword(currentPassword, account.password_hash)) {
    return renderError('Le mot de passe actuel est incorrect.', 401);
  }

  const passwordHash = await hashPassword(newPassword);
  let newToken;
  await db.transaction(async () => {
    await db.prepare(
      'UPDATE users SET password_hash = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
    ).run(passwordHash, req.user.id);
    await db.prepare('DELETE FROM user_sessions WHERE user_id = ?').run(req.user.id);
    newToken = await createSession(req.user.id);
  });
  setSessionCookie(res, newToken);
  res.redirect('/account?updated=password');
});

app.use(financeRouter);

app.get('/', async (req, res) => {
  const today = new Date();
  const requestedMonth = (req.query.month || '').toString();
  const monthKey = /^\d{4}-(0[1-9]|1[0-2])$/.test(requestedMonth)
    ? requestedMonth
    : getMonthKey(today.getFullYear(), today.getMonth() + 1);
  const requestedPeriod = Number(req.query.period);
  const period = [3, 6, 12].includes(requestedPeriod) ? requestedPeriod : 6;
  const [dashboard, allCategories, finance] = await Promise.all([
    currentDashboard(monthKey, period),
    readCategories(),
    getMonthlyFinance(monthKey),
  ]);
  const categories = sortCategoriesForPresentation(
    allCategories.filter((category) => category.type !== 'boutique'),
  );
  res.render('index', { dashboard, categories, finance });
});

app.get('/groceries', async (req, res) => {
  const search = (req.query.search || '').toString().trim();
  const requestedMonth = (req.query.month || '').toString();
  const today = new Date();
  const selectedMonth = /^\d{4}-(0[1-9]|1[0-2])$/.test(requestedMonth)
    ? requestedMonth
    : getMonthKey(today.getFullYear(), today.getMonth() + 1);
  let query = `
    SELECT * FROM grocery_items
    WHERE COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM')) = ?
  `;
  const params = [selectedMonth];

  if (search) {
    query += ' AND (item_name ILIKE ? OR unit ILIKE ? OR notes ILIKE ?)';
    params.push(`%${search}%`, `%${search}%`, `%${search}%`);
  }

  query += ' ORDER BY purchase_date DESC NULLS LAST, id DESC';
  const groceryList = await db.prepare(query).all(...params);
  const defaultPurchaseDate = selectedMonth === getMonthKey(today.getFullYear(), today.getMonth() + 1)
    ? today.toISOString().slice(0, 10)
    : `${selectedMonth}-01`;
  res.render('groceries', { groceryList, search, selectedMonth, defaultPurchaseDate });
});

function validateGroceryInput(body, {
  allowUnknownQuantity = false,
  allowUnknownDate = false,
  fallbackMonth = '',
} = {}) {
  const itemName = String(body.item_name || '').trim();
  const purchaseDateInput = String(body.purchase_date || '').trim();
  const purchaseDate = purchaseDateInput || null;
  const quantityInput = String(body.quantity ?? '').trim();
  const quantityKnown = quantityInput !== '';
  const quantity = quantityKnown ? Number(quantityInput) : 0;
  const unitPriceInput = String(body.unit_price || '').trim();
  const totalAmountInput = String(body.total_amount || '').trim();
  const unitPrice = unitPriceInput ? Number(unitPriceInput) : 0;
  const enteredTotal = totalAmountInput ? Number(totalAmountInput) : null;
  const selectedMonth = /^\d{4}-(0[1-9]|1[0-2])$/.test(String(body.selected_month || ''))
    ? String(body.selected_month)
    : purchaseDate?.slice(0, 7) || fallbackMonth;

  if (
    !itemName
    || (purchaseDate ? !isValidIsoDate(purchaseDate) : !allowUnknownDate)
    || (!quantityKnown && (!allowUnknownQuantity || enteredTotal === null))
    || (quantityKnown && (!Number.isFinite(quantity) || quantity <= 0))
    || (!unitPriceInput && enteredTotal === null)
    || (unitPriceInput && (!Number.isFinite(unitPrice) || unitPrice < 0))
    || (totalAmountInput && (!Number.isFinite(enteredTotal) || enteredTotal < 0))
  ) {
    return null;
  }

  return {
    itemName,
    purchaseDate,
    monthKey: purchaseDate?.slice(0, 7) || selectedMonth,
    quantity,
    quantityKnown: quantityKnown ? 1 : 0,
    unit: String(body.unit || ''),
    unitPrice,
    unitPriceKnown: unitPriceInput ? 1 : 0,
    totalAmount: enteredTotal ?? quantity * unitPrice,
    notes: String(body.notes || ''),
    selectedMonth,
  };
}

app.post('/groceries', async (req, res) => {
  const grocery = validateGroceryInput(req.body);
  if (!grocery) {
    return res.status(400).send('Saisissez une marchandise, une date valide, une quantité positive et un prix unitaire ou total valide.');
  }

  await db.prepare(`
    INSERT INTO grocery_items (
      item_name, purchase_date, month_key, quantity, unit, unit_price, total_amount, notes,
      quantity_known, unit_price_known
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(
    grocery.itemName,
    grocery.purchaseDate,
    grocery.monthKey,
    grocery.quantity,
    grocery.unit,
    grocery.unitPrice,
    grocery.totalAmount,
    grocery.notes,
    grocery.quantityKnown,
    grocery.unitPriceKnown,
  );

  res.redirect(req.body.return_to === '/history' ? '/history' : `/groceries?month=${grocery.selectedMonth}`);
});

app.post('/groceries/:id/update', async (req, res) => {
  const existing = await db.prepare('SELECT * FROM grocery_items WHERE id = ?').get(req.params.id);
  if (!existing) return res.status(404).send('Ravitaillement introuvable.');
  const existingMonth = existing.month_key || (existing.purchase_date ? String(existing.purchase_date).slice(0, 7) : '');
  const grocery = validateGroceryInput(req.body, {
    allowUnknownQuantity: !Number(existing.quantity_known),
    allowUnknownDate: !existing.purchase_date,
    fallbackMonth: existingMonth,
  });
  if (!grocery) {
    return res.status(400).send('Saisissez une marchandise, une date valide, une quantité positive et un prix unitaire ou total valide.');
  }

  await db.prepare(`
    UPDATE grocery_items SET
      item_name = ?, purchase_date = ?, month_key = ?, quantity = ?, unit = ?,
      unit_price = ?, total_amount = ?, notes = ?, quantity_known = ?, unit_price_known = ?
    WHERE id = ?
  `).run(
    grocery.itemName,
    grocery.purchaseDate,
    grocery.monthKey,
    grocery.quantity,
    grocery.unit,
    grocery.unitPrice,
    grocery.totalAmount,
    grocery.notes,
    grocery.quantityKnown,
    grocery.unitPriceKnown,
    req.params.id,
  );
  res.redirect(req.body.return_to === '/history'
    ? '/history'
    : `/groceries?month=${encodeURIComponent(grocery.monthKey)}`);
});

app.post('/groceries/:id/delete', async (req, res) => {
  await db.prepare('DELETE FROM grocery_items WHERE id = ?').run(req.params.id);
  res.redirect(req.body.return_to === '/history' ? '/history' : '/groceries');
});

app.get('/expenses', async (req, res) => {
  const search = (req.query.search || '').toString().trim();
  const allCategories = await readCategories();
  const categories = sortCategoriesForPresentation(
    allCategories.filter((category) => category.type !== 'boutique'),
  );
  const today = new Date();
  const requestedMonth = (req.query.month || '').toString();
  const monthKey = /^\d{4}-(0[1-9]|1[0-2])$/.test(requestedMonth)
    ? requestedMonth
    : getMonthKey(today.getFullYear(), today.getMonth() + 1);
  const page = Math.max(1, Number.parseInt(req.query.page, 10) || 1);
  const pageSize = 50;

  let where = `
    WHERE e.funding_source <> 'protected_savings'
      AND COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) = ?
  `;
  const params = [monthKey];
  if (req.query.category && /^\d+$/.test(String(req.query.category))) {
    where += ' AND e.category_id = ?';
    params.push(Number(req.query.category));
  }
  if (req.query.category_alias) {
    const alias = String(req.query.category_alias);
    const matchingIds = allCategories
      .filter((category) => resolvePriorityCategory(category.name) === alias
        || (alias === 'Compte boutique' && category.type === 'boutique')
        || (resolvePriorityCategory(category.name) === null && category.name === alias))
      .map((category) => Number(category.id));
    if (matchingIds.length) {
      where += ` AND e.category_id IN (${matchingIds.map(() => '?').join(', ')})`;
      params.push(...matchingIds);
    } else {
      where += ' AND 1 = 0';
    }
  }
  if (search) {
    where += `
      AND (e.title ILIKE ? OR e.description ILIKE ? OR e.reference ILIKE ?
        OR e.beneficiary ILIKE ? OR e.paid_by ILIKE ? OR c.name ILIKE ?)
    `;
    params.push(...Array(6).fill(`%${search}%`));
  }
  const totalCount = Number((await db.prepare(`
    SELECT COUNT(*) AS count
    FROM expenses e
    JOIN categories c ON c.id = e.category_id
    ${where}
  `).get(...params)).count);
  const expenseList = await db.prepare(`
    SELECT e.*, c.name AS category_name,
      (SELECT COUNT(*) FROM expense_attachments a WHERE a.expense_id = e.id) AS attachment_count
    FROM expenses e
    JOIN categories c ON c.id = e.category_id
    ${where}
    ORDER BY e.expense_date DESC NULLS LAST, e.id DESC
    LIMIT ? OFFSET ?
  `).all(...params, pageSize, (page - 1) * pageSize);
  const summaryExpenses = await db.prepare(`
    SELECT e.id, e.title, e.amount, e.expense_date,
      e.category_id, e.funding_source, c.name AS category_name
    FROM expenses e
    JOIN categories c ON c.id = e.category_id
    WHERE COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) = ?
    ORDER BY e.expense_date DESC, e.id DESC
  `).all(monthKey);
  const groceryItems = await db.prepare(`
    SELECT id, item_name, total_amount, purchase_date, month_key
    FROM grocery_items
    WHERE COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM')) = ?
  `).all(monthKey);
  const finance = await getMonthlyFinance(monthKey);
  const categoryPresentation = monthKey >= '2026-09'
    ? filterCategoryPresentation(
      buildCategoryPresentation({ monthKey, categories: allCategories, expenses: summaryExpenses, groceryItems }),
      allCategories,
    )
    : null;
  res.render('expenses', {
    categories, expenseList, search, monthKey, page, pageSize, totalCount, finance,
    selectedCategory: String(req.query.category || ''),
    selectedCategoryAlias: String(req.query.category_alias || ''),
    categoryPresentation,
  });
});

function validateExpenseInput(body) {
  const title = String(body.title || '').trim();
  const categoryId = Number(body.category_id);
  const amount = Number(body.amount);
  const date = String(body.expense_date || '').trim();
  const parsedDate = date ? new Date(`${date}T00:00:00.000Z`) : null;
  const monthInput = String(body.month_key || '');
  const monthKey = /^\d{4}-(0[1-9]|1[0-2])$/.test(monthInput)
    ? monthInput
    : date
      ? date.slice(0, 7)
      : '';
  if (!title || title.length > 250 || !Number.isSafeInteger(categoryId)
    || !Number.isFinite(amount) || amount <= 0 || !monthKey
    || (date && (!/^\d{4}-\d{2}-\d{2}$/.test(date)
      || Number.isNaN(parsedDate.getTime()) || parsedDate.toISOString().slice(0, 10) !== date))) {
    return null;
  }
  return {
    title,
    categoryId,
    amount,
    date: date || null,
    monthKey,
    description: String(body.description || '').trim(),
    beneficiary: String(body.beneficiary || '').trim(),
    paidBy: String(body.paid_by || '').trim(),
    paymentMethod: String(body.payment_method || '').trim(),
    reference: String(body.reference || '').trim(),
  };
}

function parseExpenseIds(value) {
  const submittedIds = Array.isArray(value) ? value : [value];
  if (!submittedIds.length || submittedIds.length > 50) return null;
  const ids = submittedIds.map((id) => /^\d+$/.test(String(id || '')) ? Number(id) : NaN);
  if (ids.some((id) => !Number.isSafeInteger(id) || id <= 0) || new Set(ids).size !== ids.length) {
    return null;
  }
  return ids;
}

function dateForMonth(date, monthKey) {
  if (!date) return null;
  const day = Number(String(date).slice(8, 10));
  const [year, month] = monthKey.split('-').map(Number);
  const lastDay = new Date(Date.UTC(year, month, 0)).getUTCDate();
  return `${monthKey}-${String(Math.min(day, lastDay)).padStart(2, '0')}`;
}

function includeCarryoverInExport(value) {
  if (value === undefined) return false;
  const submittedValues = Array.isArray(value) ? value : [value];
  return submittedValues.some((item) => String(item) === '1');
}

async function storeReceipt(file) {
  if (!file) return null;
  const signatureValid = {
    'application/pdf': file.buffer.subarray(0, 5).toString() === '%PDF-',
    'image/jpeg': file.buffer[0] === 0xff && file.buffer[1] === 0xd8 && file.buffer.at(-2) === 0xff && file.buffer.at(-1) === 0xd9,
    'image/png': file.buffer.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10])),
    'image/webp': file.buffer.subarray(0, 4).toString() === 'RIFF'
      && file.buffer.subarray(8, 12).toString() === 'WEBP',
  }[file.mimetype];
  if (!signatureValid) {
    throw Object.assign(new Error('Le contenu du justificatif ne correspond pas à son format.'), { status: 400 });
  }
  const extension = {
    'application/pdf': '.pdf',
    'image/jpeg': '.jpg',
    'image/png': '.png',
    'image/webp': '.webp',
  }[file.mimetype];
  const filename = `${randomUUID()}${extension}`;
  await fs.promises.writeFile(path.join(receiptDir, filename), file.buffer, { flag: 'wx' });
  return { filename, mimetype: file.mimetype, originalName: path.basename(file.originalname).slice(0, 255) };
}

app.post('/expenses', receiptUpload.single('receipt'), async (req, res) => {
  const expense = validateExpenseInput(req.body);
  if (!expense) return res.status(400).send('Vérifiez la date, la désignation, la catégorie et le montant.');
  const category = await db.prepare('SELECT id FROM categories WHERE id = ? AND type <> ?').get(expense.categoryId, 'boutique');
  if (!category) return res.status(400).send('Catégorie de dépense invalide.');
  const receipt = await storeReceipt(req.file);
  try {
    await db.transaction(async () => {
      const result = await db.prepare(`
        INSERT INTO expenses (
          title, category_id, amount, expense_date, description, reference, beneficiary,
          paid_by, payment_method, month_key, created_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).run(
        expense.title, expense.categoryId, expense.amount, expense.date,
        expense.description, expense.reference, expense.beneficiary, expense.paidBy,
        expense.paymentMethod, expense.monthKey, req.user.id,
      );
      if (receipt) {
        await db.prepare(`
          INSERT INTO expense_attachments (expense_id, file_name, storage_path, mime_type, file_size, uploaded_by)
          VALUES (?, ?, ?, ?, ?, ?)
        `).run(result.lastInsertRowid, receipt.originalName, receipt.filename, receipt.mimetype, req.file.size, req.user.id);
      }
      await writeAudit(req.user.id, 'expense', result.lastInsertRowid, 'create', null, expense);
    });
  } catch (error) {
    if (receipt) await fs.promises.unlink(path.join(receiptDir, receipt.filename)).catch(() => {});
    throw error;
  }
  res.redirect(req.body.return_to === '/history' ? '/history' : `/expenses?month=${encodeURIComponent(expense.monthKey)}`);
});

app.post('/expenses/:id/update', receiptUpload.single('receipt'), async (req, res) => {
  const id = Number(req.params.id);
  const expense = validateExpenseInput(req.body);
  if (!Number.isSafeInteger(id) || !expense) return res.status(400).send('Données de dépense invalides.');
  const existing = await db.prepare('SELECT funding_source FROM expenses WHERE id = ?').get(id);
  if (existing?.funding_source === 'protected_savings') {
    return res.status(409).send('Cette dépense est liée à un retrait de l’épargne protégée et ne peut pas être modifiée ici.');
  }
  const category = await db.prepare('SELECT id FROM categories WHERE id = ? AND type <> ?').get(expense.categoryId, 'boutique');
  if (!category) return res.status(400).send('Catégorie de dépense invalide.');
  const before = await db.prepare('SELECT * FROM expenses WHERE id = ?').get(id);
  if (!before) return res.status(404).send('Dépense introuvable.');
  const receipt = await storeReceipt(req.file);
  try {
    await db.transaction(async () => {
      await db.prepare(`
        UPDATE expenses SET title = ?, category_id = ?, amount = ?, expense_date = ?,
          description = ?, reference = ?, beneficiary = ?, paid_by = ?, payment_method = ?,
          month_key = ?, updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
      `).run(
        expense.title, expense.categoryId, expense.amount, expense.date, expense.description,
        expense.reference, expense.beneficiary, expense.paidBy, expense.paymentMethod,
        expense.monthKey, id,
      );
      if (receipt) {
        await db.prepare(`
          INSERT INTO expense_attachments (expense_id, file_name, storage_path, mime_type, file_size, uploaded_by)
          VALUES (?, ?, ?, ?, ?, ?)
        `).run(id, receipt.originalName, receipt.filename, receipt.mimetype, req.file.size, req.user.id);
      }
      const after = await db.prepare('SELECT * FROM expenses WHERE id = ?').get(id);
      await writeAudit(req.user.id, 'expense', id, 'update', before, after);
    });
  } catch (error) {
    if (receipt) await fs.promises.unlink(path.join(receiptDir, receipt.filename)).catch(() => {});
    throw error;
  }
  res.redirect(req.body.return_to === '/history' ? '/history' : `/expenses?month=${encodeURIComponent(expense.monthKey)}`);
});

app.post('/expenses/:id/delete', async (req, res) => {
  const id = Number(req.params.id);
  const existing = await db.prepare('SELECT funding_source FROM expenses WHERE id = ?').get(id);
  if (existing?.funding_source === 'protected_savings') {
    return res.status(409).send('Cette dépense est liée à un retrait de l’épargne protégée et ne peut pas être supprimée ici.');
  }
  const removedFiles = await db.prepare('SELECT storage_path FROM expense_attachments WHERE expense_id = ?').all(id);
  await db.transaction(async () => {
    const before = await db.prepare('SELECT * FROM expenses WHERE id = ? FOR UPDATE').get(id);
    if (!before) return;
    await db.prepare('DELETE FROM expenses WHERE id = ?').run(id);
    await writeAudit(req.user.id, 'expense', id, 'delete', before, null);
  });
  await Promise.all(removedFiles.map(({ storage_path }) => fs.promises
    .unlink(path.join(receiptDir, path.basename(storage_path))).catch(() => {})));
  res.redirect(req.body.return_to === '/history' ? '/history' : '/expenses');
});

app.post('/expenses/bulk', async (req, res) => {
  const ids = parseExpenseIds(req.body.expense_ids);
  const action = String(req.body.action || '');
  const monthKey = String(req.body.month_key || '');
  const returnTo = String(req.body.return_to || '');
  if (!ids || !['delete', 'duplicate'].includes(action)) {
    return res.status(400).send('Sélectionnez entre 1 et 50 dépenses valides.');
  }
  const targetMonth = String(req.body.target_month || '');
  if (action === 'duplicate' && !/^\d{4}-(0[1-9]|1[0-2])$/.test(targetMonth)) {
    return res.status(400).send('Choisissez un mois de destination valide.');
  }

  let removedFiles = [];
  try {
    await db.transaction(async () => {
      const expenses = await db.prepare(`
        SELECT * FROM expenses WHERE id IN (${ids.map(() => '?').join(', ')}) FOR UPDATE
      `).all(...ids);
      if (expenses.length !== ids.length) {
        throw Object.assign(
          new Error('Certaines dépenses sélectionnées ne sont plus disponibles. Actualisez l’historique puis réessayez.'),
          { status: 404 },
        );
      }
      if (expenses.some((expense) => expense.funding_source === 'protected_savings')) {
        throw Object.assign(
          new Error('Les dépenses financées par l’épargne protégée ne peuvent pas être modifiées par ces actions.'),
          { status: 409 },
        );
      }

      if (action === 'delete') {
        removedFiles = await db.prepare(`
          SELECT storage_path FROM expense_attachments WHERE expense_id IN (${ids.map(() => '?').join(', ')})
        `).all(...ids);
        for (const expense of expenses) {
          await db.prepare('DELETE FROM expenses WHERE id = ?').run(expense.id);
          await writeAudit(req.user.id, 'expense', expense.id, 'delete', expense, null);
        }
        return;
      }

      for (const expense of expenses) {
        const newDate = dateForMonth(expense.expense_date, targetMonth);
        const result = await db.prepare(`
          INSERT INTO expenses (
            title, category_id, amount, expense_date, description, reference, beneficiary,
            paid_by, payment_method, month_key, created_by
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `).run(
          expense.title, expense.category_id, expense.amount, newDate, expense.description,
          expense.reference, expense.beneficiary, expense.paid_by, expense.payment_method,
          targetMonth, req.user.id,
        );
        const after = await db.prepare('SELECT * FROM expenses WHERE id = ?').get(result.lastInsertRowid);
        await writeAudit(req.user.id, 'expense', result.lastInsertRowid, 'duplicate', expense, after);
      }
    });
  } catch (error) {
    if (error.status) return res.status(error.status).send(error.message);
    throw error;
  }

  await Promise.all(removedFiles.map(({ storage_path }) => fs.promises
    .unlink(path.join(receiptDir, path.basename(storage_path))).catch(() => {})));
  if (returnTo === '/history') {
    const historyMonth = action === 'duplicate' ? targetMonth : monthKey;
    return res.redirect(`/history${/^\d{4}-(0[1-9]|1[0-2])$/.test(historyMonth) ? `?month=${encodeURIComponent(historyMonth)}&type=expense` : ''}`);
  }
  const returnMonth = returnTo.match(/^\/expenses\?month=(\d{4}-(?:0[1-9]|1[0-2]))$/)?.[1];
  const destinationMonth = action === 'duplicate' ? targetMonth : (returnMonth || monthKey);
  const redirectMonth = /^\d{4}-(0[1-9]|1[0-2])$/.test(destinationMonth)
    ? destinationMonth
    : getMonthKey(new Date().getFullYear(), new Date().getMonth() + 1);
  res.redirect(`/expenses?month=${encodeURIComponent(redirectMonth)}`);
});

app.get('/expenses/:id/receipt', async (req, res) => {
  const attachment = await db.prepare(`
    SELECT storage_path, mime_type, file_name
    FROM expense_attachments WHERE expense_id = ? ORDER BY created_at DESC LIMIT 1
  `).get(Number(req.params.id));
  if (!attachment) return res.status(404).send('Justificatif introuvable.');
  res.type(attachment.mime_type);
  res.set('Content-Disposition', `inline; filename="${attachment.file_name.replace(/["\\\r\n]/g, '_')}"`);
  res.sendFile(path.basename(attachment.storage_path), { root: receiptDir });
});

app.get('/history', async (req, res) => {
  const search = (req.query.search || '').toString().trim();
  const isSearchApi = req.query.format === 'json';
  if (isSearchApi && !search) return res.json({ totalCount: 0, operations: [] });
  const month = /^\d{4}-(0[1-9]|1[0-2])$/.test(String(req.query.month || ''))
    ? String(req.query.month)
    : '';
  const type = ['expense', 'grocery', 'income', 'rent', 'savings'].includes(String(req.query.type || ''))
    ? String(req.query.type)
    : '';
  const startDate = isValidIsoDate(req.query.start)
    ? String(req.query.start)
    : '';
  const endDate = isValidIsoDate(req.query.end)
    ? String(req.query.end)
    : '';
  const page = Math.max(1, Number.parseInt(req.query.page, 10) || 1);
  const pageSize = 50;
  const operationQuery = `
    SELECT * FROM (
      SELECT e.id, e.title, e.amount, e.expense_date AS date, e.funding_source,
        COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) AS month,
        COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) AS source_month,
        CASE WHEN e.funding_source = 'protected_savings'
          THEN (SELECT s.id FROM savings_transactions s WHERE s.related_expense_id = e.id ORDER BY s.id DESC LIMIT 1)
          ELSE e.id
        END AS source_id,
        c.name AS category, 'expense' AS type, e.beneficiary,
        CONCAT_WS(' ', e.title, c.name, e.beneficiary, e.paid_by, e.funding_source) AS search_text
      FROM expenses e
      JOIN categories c ON c.id = e.category_id
      UNION ALL
      SELECT g.id, g.item_name AS title, g.total_amount AS amount, g.purchase_date AS date, 'current' AS funding_source,
        COALESCE(NULLIF(g.month_key, ''), to_char(g.purchase_date, 'YYYY-MM')) AS month,
        COALESCE(NULLIF(g.month_key, ''), to_char(g.purchase_date, 'YYYY-MM')) AS source_month,
        g.id AS source_id, 'Compte boutique' AS category,
        'grocery' AS type, NULL AS beneficiary,
        CONCAT_WS(' ', g.item_name, g.notes, 'Compte boutique') AS search_text
      FROM grocery_items g
      UNION ALL
      SELECT i.id, i.title, i.amount, i.income_date AS date, NULL AS funding_source, i.month_key AS month,
        i.month_key AS source_month, i.id AS source_id, c.name AS category, 'income' AS type, i.source AS beneficiary,
        CONCAT_WS(' ', i.title, c.name, i.source) AS search_text
      FROM incomes i JOIN income_categories c ON c.id = i.category_id
      UNION ALL
      SELECT p.id, CONCAT('Loyer ', t.first_name, ' ', t.last_name) AS title,
        p.amount, p.payment_date AS date, NULL AS funding_source,
        COALESCE(NULLIF(p.received_month_key, ''), to_char(p.payment_date, 'YYYY-MM'), o.month_key) AS month,
        o.month_key AS source_month, p.id AS source_id, 'Loyer' AS category,
        'rent' AS type, CONCAT(u.floor, ' · ', u.label) AS beneficiary,
        CONCAT_WS(' ', t.first_name, t.last_name, u.floor, u.label, p.payment_method) AS search_text
      FROM rent_payments p
      JOIN rent_obligations o ON o.id = p.obligation_id
      JOIN tenants t ON t.id = o.tenant_id
      JOIN units u ON u.id = t.unit_id
      UNION ALL
      SELECT s.id, COALESCE(e.title, a.name) AS title, s.amount, s.transaction_date AS date,
        CASE WHEN s.related_expense_id IS NOT NULL THEN 'protected_savings' ELSE NULL END AS funding_source,
        to_char(s.transaction_date, 'YYYY-MM') AS month, to_char(s.transaction_date, 'YYYY-MM') AS source_month,
        s.id AS source_id, 'Épargne' AS category,
        'savings' AS type, s.transaction_type AS beneficiary,
        CONCAT_WS(' ', a.name, s.transaction_type, s.comment) AS search_text
      FROM savings_transactions s JOIN savings_accounts a ON a.id = s.account_id
      LEFT JOIN expenses e ON e.id = s.related_expense_id
    )
  `;
  const params = [];
  let where = ' WHERE TRUE';
  if (month) {
    where += ' AND month = ?';
    params.push(month);
  }
  if (startDate) {
    where += ' AND date >= ?';
    params.push(startDate);
  }
  if (endDate) {
    where += ' AND date <= ?';
    params.push(endDate);
  }
  if (type) {
    where += ' AND type = ?';
    params.push(type);
  }

  if (search) {
    where += ' AND search_text ILIKE ?';
    params.push(`%${search}%`);
  }

  const totalCount = Number((await db.prepare(`
    SELECT COUNT(*) AS count FROM (${operationQuery}) operations ${where}
  `).get(...params)).count);
  const operationList = isSearchApi
    ? await db.prepare(`
      SELECT * FROM (${operationQuery}) operations ${where}
      ORDER BY date DESC NULLS LAST, type, id DESC
    `).all(...params)
    : await db.prepare(`
      SELECT * FROM (${operationQuery}) operations ${where}
      ORDER BY date DESC NULLS LAST, type, id DESC
      LIMIT ? OFFSET ?
    `).all(...params, pageSize, (page - 1) * pageSize);
  if (isSearchApi) {
    const operations = operationList.map((item) => {
      let href;
      if (item.type === 'expense' && item.funding_source === 'protected_savings') {
        href = item.source_id
          ? `/savings?open=savings-details-${encodeURIComponent(item.source_id)}`
          : '/savings';
      } else if (item.type === 'expense') {
        href = `/expenses?month=${encodeURIComponent(item.source_month)}&search=${encodeURIComponent(item.title)}&open=expense-details-${item.id}`;
      } else if (item.type === 'grocery') {
        href = `/groceries?month=${encodeURIComponent(item.source_month)}&search=${encodeURIComponent(item.title)}&open=grocery-details-${item.id}`;
      } else if (item.type === 'income') {
        href = `/income?month=${encodeURIComponent(item.source_month)}&open=income-details-${item.id}`;
      } else if (item.type === 'rent') {
        href = `/rents?month=${encodeURIComponent(item.source_month)}&open=rent-payment-details-${item.id}`;
      } else {
        href = `/savings?open=savings-details-${item.source_id}`;
      }
      return {
        id: item.id,
        type: item.type,
        title: item.title,
        amount: Number(item.amount),
        date: item.date,
        category: item.category,
        source_month: item.source_month,
        source_id: item.source_id,
        funding_source: item.funding_source,
        search_text: item.search_text,
        href,
      };
    });
    return res.json({ totalCount, operations });
  }
  const editableExpenseIds = operationList
    .filter((item) => item.type === 'expense' && item.funding_source !== 'protected_savings')
    .map((item) => item.id);
  const editableExpenses = editableExpenseIds.length
    ? await db.prepare(`
      SELECT * FROM expenses WHERE id IN (${editableExpenseIds.map(() => '?').join(', ')})
    `).all(...editableExpenseIds)
    : [];
  const editableExpensesById = Object.fromEntries(editableExpenses.map((expense) => [expense.id, expense]));
  const editableGroceryIds = operationList
    .filter((item) => item.type === 'grocery')
    .map((item) => item.source_id);
  const editableGroceries = editableGroceryIds.length
    ? await db.prepare(`
      SELECT * FROM grocery_items WHERE id IN (${editableGroceryIds.map(() => '?').join(', ')})
    `).all(...editableGroceryIds)
    : [];
  const editableGroceriesById = Object.fromEntries(editableGroceries.map((grocery) => [grocery.id, grocery]));
  const categories = sortCategoriesForPresentation(
    (await readCategories()).filter((category) => category.type !== 'boutique'),
  );
  const pageCount = Math.max(1, Math.ceil(totalCount / pageSize));
  const pageUrl = (pageNumber) => {
    const queryParams = new URLSearchParams();
    if (search) queryParams.set('search', search);
    if (month) queryParams.set('month', month);
    if (type) queryParams.set('type', type);
    if (startDate) queryParams.set('start', startDate);
    if (endDate) queryParams.set('end', endDate);
    queryParams.set('page', String(pageNumber));
    return `/history?${queryParams.toString()}`;
  };
  res.render('history', {
    operationList, editableExpensesById, editableGroceriesById, categories, search, month, type, startDate, endDate,
    page: Math.min(page, pageCount), pageCount, totalCount, pageUrl,
  });
});

app.get('/categories', async (req, res) => {
  const categories = sortCategoriesForPresentation(await readCategories());
  const requestedReturnTo = String(req.query.return_to || '');
  const returnTo = /^\/expenses\?month=\d{4}-(0[1-9]|1[0-2])$/.test(requestedReturnTo)
    ? requestedReturnTo
    : '/categories';
  res.render('categories', { categories, returnTo });
});

app.post('/categories', async (req, res) => {
  const { name, color } = req.body;
  const categoryName = String(name || '').trim();
  if (!categoryName || categoryName.length > 100) return res.status(400).send('Le nom de la catégorie est obligatoire (100 caractères maximum).');
  const categoryColor = /^#[0-9a-f]{6}$/i.test(String(color || '')) ? color : '#147967';
  await db.transaction(async () => {
    const result = await db.prepare('INSERT INTO categories (name, color, type) VALUES (?, ?, ?)').run(categoryName, categoryColor, 'ordinary');
    await db.prepare('DELETE FROM category_seed_exclusions WHERE name = ?').run(categoryName);
    await writeAudit(req.user.id, 'category', result.lastInsertRowid, 'create', null, { name: categoryName, color: categoryColor });
  });
  const returnTo = String(req.body.return_to || '');
  const expensesReturnTo = /^\/expenses\?month=\d{4}-(0[1-9]|1[0-2])$/.test(returnTo) ? returnTo : null;
  res.redirect(expensesReturnTo || '/categories');
});

app.get('/monthly', async (req, res) => {
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

  const monthlyOther = await db.prepare(`
    SELECT e.title, e.amount, e.expense_date, c.name AS category
    FROM expenses e
    JOIN categories c ON c.id = e.category_id
    WHERE e.funding_source <> 'protected_savings'
      AND COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) = ?
    ORDER BY e.expense_date ASC
  `).all(reportDate);
  const protectedSavingsExpenses = await db.prepare(`
    SELECT e.title, e.amount, e.expense_date, c.name AS category
    FROM expenses e
    JOIN categories c ON c.id = e.category_id
    WHERE e.funding_source = 'protected_savings'
      AND COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) = ?
    ORDER BY e.expense_date ASC
  `).all(reportDate);

  const monthlyGroceries = await db.prepare(`
    SELECT *
    FROM grocery_items
    WHERE COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM')) = ?
    ORDER BY purchase_date ASC NULLS LAST
  `).all(reportDate);

  const totalGroceries = monthlyGroceries.reduce((sum, row) => sum + Number(row.total_amount || 0), 0);
  const totalOther = monthlyOther.reduce((sum, row) => sum + Number(row.amount || 0), 0);
  const finance = await getMonthlyFinance(reportDate);
  const allCategories = await readCategories();
  const summaryExpenses = [
    ...monthlyOther.map((row) => ({ ...row, category_name: row.category })),
    ...protectedSavingsExpenses.map((row) => ({ ...row, category_name: row.category, funding_source: 'protected_savings' })),
  ];
  const categoryPresentation = reportDate >= '2026-09'
    ? filterCategoryPresentation(buildCategoryPresentation({
      monthKey: reportDate,
      categories: allCategories,
      expenses: summaryExpenses,
      groceryItems: monthlyGroceries,
    }), allCategories)
    : null;

  res.render('monthly', {
    year,
    month,
    monthLabel,
    monthlyOther,
    protectedSavingsExpenses,
    monthlyGroceries,
    totalGroceries,
    totalOther,
    totalGeneral: totalGroceries + totalOther,
    finance,
    categoryPresentation,
    categories: sortCategoriesForPresentation(
      allCategories.filter((category) => category.type !== 'boutique'),
    ),
  });
});

app.get('/import', async (req, res) => {
  const latestBatch = await getLatestImportBatch();
  let latestReport = null;
  if (latestBatch?.status === 'imported') {
    const payload = JSON.parse(latestBatch.rows_json || '{}');
    const records = payload.records || [];
    latestReport = {
      detectedCount: (payload.classifiedRows || []).filter((row) => row.classification !== 'blank').length,
      expenseCount: records.filter((row) => row.valid && !row.duplicateSource && row.type === 'expense').length,
      groceryCount: records.filter((row) => row.valid && !row.duplicateSource && row.type === 'grocery').length,
      tenantCount: (payload.tenants || []).filter((row) => row.valid).length,
      rentPaymentCount: (payload.rentPayments || []).filter((row) => row.valid && !row.duplicateSource).length,
      incomeCount: (payload.incomes || []).filter((row) => row.valid && !row.duplicateSource).length,
      savingsCount: (payload.savings || []).filter((row) => row.valid && !row.duplicateSource).length,
      totalCount: (payload.totals || []).length,
      recognizedCount: (payload.classifiedRows || []).filter((row) => row.classification !== 'blank').length,
      unknownCategoryCount: records.filter((row) => row.valid && !row.duplicateSource
        && row.type === 'expense' && !row.categoryRecognized).length,
      duplicateCount: records.filter((row) => row.duplicateSource).length
        + (payload.rentPayments || []).filter((row) => row.duplicateSource).length
        + (payload.incomes || []).filter((row) => row.duplicateSource).length
        + (payload.savings || []).filter((row) => row.duplicateSource).length,
      incompleteCount: records.filter((row) => !row.valid).length,
      excludedCount: (payload.excludedRows || []).length,
      unparsedCount: (payload.ignoredRows || []).filter((row) => row.values).length,
      monthly: payload.monthly || [],
    };
  }
  res.render('import', {
    latestBatch,
    latestReport,
    initialImportComplete: await hasInitialHistoricalImport(),
    r2Configured,
  });
});

app.post('/import', importFileUpload, async (req, res) => {
  const file = req.file;
  if (!file) return res.redirect('/import');

  let parsed;
  try {
    parsed = parseHistoricalFile(file.buffer, file.originalname, await readCategories());
  } catch (error) {
    console.error('Excel import parse failed:', error.message);
    return res.status(400).send('Le fichier Excel ne peut pas être lu. Vérifiez le format puis réessayez.');
  }
  const rows = markDuplicates(parsed.records, await existingHistoricalRows());
  const payload = JSON.stringify({
    ...parsed,
    records: rows,
    monthly: summarizeRecords(rows.map((row) => ({
      ...row,
      duplicateSource: row.duplicateSource === 'fichier' ? 'fichier' : null,
    }))),
  });
  const fileName = path.basename(file.originalname.replace(/\\/g, '/')).slice(0, 255);
  let previewId;
  let archivePath;
  try {
    await db.transaction(async () => {
      const result = await db.prepare('INSERT INTO import_batches (file_name, rows_json, status) VALUES (?, ?, ?)')
        .run(fileName, payload, 'preview');
      previewId = result.lastInsertRowid;
      archivePath = path.join(importDir, `${randomUUID()}_${fileName}`);
      await fs.promises.writeFile(archivePath, file.buffer, { flag: 'wx' });
    });
  } catch (error) {
    if (archivePath) await fs.promises.unlink(archivePath).catch(() => {});
    throw error;
  }
  res.redirect(`/import/preview/${previewId}`);
});

app.get('/import/preview/:id', async (req, res) => {
  const previewBatch = await db.prepare('SELECT * FROM import_batches WHERE id = ?').get(req.params.id);
  if (!previewBatch) return res.redirect('/import');

  const payload = JSON.parse(previewBatch.rows_json || '{}');
  const records = payload.records || [];
  const duplicates = records.filter((row) => row.duplicateSource);
  const validCount = records.filter((row) => row.valid).length
    + (payload.tenants || []).filter((row) => row.valid).length
    + (payload.incomes || []).filter((row) => row.valid).length
    + (payload.savings || []).filter((row) => row.valid).length;
  const reinitializeAvailable = !await hasInitialHistoricalImport() && validCount > 0;
  const monthly = payload.monthly || summarizeRecords(records);

  res.render('import-preview', {
    previewBatch,
    records,
    duplicates,
    monthly,
    tenants: payload.tenants || [],
    rentPayments: payload.rentPayments || [],
    incomes: payload.incomes || [],
    savings: payload.savings || [],
    totals: payload.totals || [],
    classifiedRows: payload.classifiedRows || [],
    periodWarnings: payload.periodWarnings || [],
    excludedRows: payload.excludedRows || [],
    ignoredRows: payload.ignoredRows || [],
    summaryLines: payload.summaryLines || [],
    sheetNames: payload.sheetNames || [],
    reinitializeAvailable,
    initialImportComplete: await hasInitialHistoricalImport(),
  });
});

app.post('/import/preview/:id/correct', async (req, res) => {
  const previewBatch = await db.prepare('SELECT * FROM import_batches WHERE id = ?').get(req.params.id);
  if (!previewBatch || previewBatch.status === 'imported') return res.redirect('/import');
  const payload = JSON.parse(previewBatch.rows_json || '{}');
  const sheetName = String(req.body.sheetName || '');
  const rowNumber = Number(req.body.rowNumber);
  const sourceRow = (payload.ignoredRows || []).find((row) => (
    row.sheetName === sheetName && row.row === rowNumber && row.values
  ));
  if (!sourceRow) return res.status(404).send('Ligne source introuvable dans cette prévisualisation.');

  const kind = String(req.body.kind || '');
  const title = String(req.body.title || '').trim();
  const categoryName = String(req.body.categoryName || '').trim();
  const monthKey = String(req.body.monthKey || '').trim();
  const amount = parseAmount(req.body.amount);
  if (!['expense', 'income', 'total', 'header'].includes(kind)) {
    return res.status(400).send('Type de correction non pris en charge.');
  }
  if (['expense', 'income'].includes(kind)
    && (!title || !categoryName || !/^\d{4}-(0[1-9]|1[0-2])$/.test(monthKey)
      || amount === null || amount <= 0)) {
    return res.status(400).send('Pour une dépense ou une recette, indiquez un intitulé, une catégorie, un mois valide et un montant positif.');
  }
  if (kind === 'total' && (amount === null || !title)) {
    return res.status(400).send('Pour un total, indiquez son libellé et son montant.');
  }

  if (kind === 'expense') {
    payload.records = payload.records || [];
    payload.records.push({
      lineNumber: payload.records.length + 1,
      sheetName,
      rowNumber,
      title,
      date: null,
      monthKey,
      amount,
      type: 'expense',
      categoryRaw: categoryName,
      categoryName,
      categoryRecognized: true,
      quantity: null,
      unit: null,
      unitPrice: null,
      quantityKnown: false,
      unitPriceKnown: false,
      notes: 'Ligne corrigée manuellement depuis l’aperçu d’import.',
      reference: '',
      beneficiary: '',
      paidBy: '',
      paymentMethod: '',
      issues: [],
      valid: true,
      duplicateSource: null,
    });
  } else if (kind === 'income') {
    payload.incomes = payload.incomes || [];
    payload.incomes.push({
      title,
      source: String(req.body.source || categoryName).trim(),
      categoryName,
      amount,
      monthKey,
      date: null,
      sheetName,
      rowNumber,
      valid: true,
      issues: [],
    });
  } else if (kind === 'total') {
    payload.totals = payload.totals || [];
    payload.totals.push({
      label: title,
      amount,
      monthKey: /^\d{4}-(0[1-9]|1[0-2])$/.test(monthKey) ? monthKey : null,
      kind: 'manual-total',
      sheetName,
      rowNumber,
      formulas: [],
    });
  }

  const classification = kind === 'total' ? 'total' : kind === 'header' ? 'header' : kind;
  payload.classifiedRows = (payload.classifiedRows || []).map((row) => (
    row.sheetName === sheetName && row.row === rowNumber
      ? { ...row, classification, classifications: [classification], reason: 'Classée manuellement.' }
      : row
  ));
  payload.ignoredRows = payload.ignoredRows.filter((row) => (
    row.sheetName !== sheetName || row.row !== rowNumber
  ));
  payload.records = markDuplicates(payload.records || [], await existingHistoricalRows());
  payload.monthly = summarizeRecords(payload.records);
  await db.prepare('UPDATE import_batches SET rows_json = ? WHERE id = ?')
    .run(JSON.stringify(payload), previewBatch.id);
  res.redirect(`/import/preview/${previewBatch.id}`);
});

app.post('/import/confirm/:id', async (req, res) => {
  const previewBatch = await db.prepare('SELECT * FROM import_batches WHERE id = ?').get(req.params.id);
  if (!previewBatch) return res.redirect('/import');
  if (previewBatch.status === 'imported') {
    return res.status(409).send('Ce fichier a déjà été importé.');
  }

  const payload = JSON.parse(previewBatch.rows_json || '{}');
  const reinitialize = req.body?.reinitialize === '1';
  const alreadyInitialized = await hasInitialHistoricalImport();
  if (!alreadyInitialized && !reinitialize) {
    return res.status(409).send('La première importation 2026 doit être confirmée avec sauvegarde et nettoyage.');
  }
  if (alreadyInitialized && reinitialize) {
    return res.status(409).send('Le nettoyage initial a déjà été effectué. Il ne peut pas être répété.');
  }

  const sourceRecords = (payload.records || []).filter((row) => row.valid);
  const tenants = (payload.tenants || []).filter((row) => row.valid);
  const rentPayments = (payload.rentPayments || []).filter((row) => row.valid);
  const incomes = (payload.incomes || []).filter((row) => row.valid);
  const savings = (payload.savings || []).filter((row) => row.valid);
  const invalidTenantCount = (payload.tenants || []).filter((row) => !row.valid).length;
  const invalidSavingsCount = (payload.savings || []).filter((row) => !row.valid).length;
  const unrecognizedRows = (payload.ignoredRows || []).filter((row) => row.values);
  if (unrecognizedRows.length || invalidTenantCount || invalidSavingsCount) {
    return res.status(400).send('Importation bloquée : des lignes non reconnues ou des locataires incomplets doivent être examinés avant validation.');
  }
  const unresolvedBoutiqueSummaries = (payload.summaryLines || [])
    .filter((summary) => summary.unresolved || summary.mismatch);
  if (unresolvedBoutiqueSummaries.length) {
    return res.status(400).send('Importation annulée : un total boutique ne correspond pas à ses marchandises détaillées.');
  }
  const mismatchedExpenseTotals = (payload.totals || [])
    .filter((total) => total.kind === 'expense-total' && total.mismatch);
  if (mismatchedExpenseTotals.length) {
    return res.status(400).send('Importation annulée : un total de dépenses ne correspond pas aux opérations détaillées.');
  }
  const records = markDuplicates(
    sourceRecords,
    reinitialize ? [] : await existingHistoricalRows(),
  );
  if (records.some((row) => row.duplicateSource === 'fichier')) {
    return res.status(400).send('Importation annulée : des lignes identiques dans le fichier doivent être vérifiées avant de poursuivre.');
  }
  const toImport = records.filter((row) => !row.duplicateSource);
  if (!toImport.length && !tenants.length && !rentPayments.length && !incomes.length && !savings.length) {
    return res.status(400).send('Aucune donnée valide et non dupliquée à importer pour la période demandée.');
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

  const transaction = () => db.transaction(async () => {
    if (reinitialize) {
      if (await hasInitialHistoricalImport()) throw new Error('Le nettoyage initial a déjà été effectué.');
      await db.prepare('DELETE FROM expenses').run();
      await db.prepare('DELETE FROM grocery_items').run();
      await db.prepare('DELETE FROM import_batches').run();
    }

    const expenseIds = [];
    const groceryIds = [];
    for (const row of toImport) {
      if (row.type === 'grocery') {
        const result = await db.prepare(`
          INSERT INTO grocery_items (
            item_name, purchase_date, month_key, quantity, unit, unit_price, total_amount, notes,
            quantity_known, unit_price_known
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `).run(
          row.title,
          row.date || null,
          row.monthKey || row.date?.slice(0, 7) || null,
          row.quantityKnown ? row.quantity : 0,
          row.unit || '',
          row.unitPriceKnown ? row.unitPrice : 0,
          row.amount,
          row.notes || '',
          row.quantityKnown ? 1 : 0,
          row.unitPriceKnown ? 1 : 0,
        );
        groceryIds.push(Number(result.lastInsertRowid));
        continue;
      }

      const categoryId = await ensureCategoryByName(row.categoryName || 'Autres');
      const result = await db.prepare(`
        INSERT INTO expenses (
          title, category_id, amount, expense_date, description, reference, beneficiary,
          paid_by, payment_method, month_key, created_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).run(
        row.title,
        categoryId,
        row.amount,
        row.date || null,
        row.notes || '',
        row.reference || '',
        row.beneficiary || '',
        row.paidBy || '',
        row.paymentMethod || '',
        /^\d{4}-(0[1-9]|1[0-2])$/.test(row.monthKey || '')
          ? row.monthKey
          : row.date?.slice(0, 7) || null,
        req.user.id,
      );
      expenseIds.push(Number(result.lastInsertRowid));
    }

    const tenantIds = new Map();
    for (const tenant of tenants) {
      let existingTenant = await db.prepare(`
        SELECT t.id, t.unit_id AS "unitId", t.monthly_rent AS "monthlyRent",
          t.people_count AS "peopleCount"
        FROM tenants t JOIN units u ON u.id = t.unit_id
        WHERE lower(t.first_name) = lower(?) AND lower(t.last_name) = lower(?) AND u.floor = ?
        ORDER BY t.id LIMIT 1
      `).get(tenant.firstName, tenant.lastName, tenant.floor);
      if (existingTenant) {
        if (Number(existingTenant.monthlyRent) !== tenant.monthlyRent
          || Number(existingTenant.peopleCount) !== tenant.peopleCount) {
          throw new Error(`Le locataire ${tenant.fullName} existe avec des informations différentes.`);
        }
      } else {
        const unitLabel = `Import Excel - ${tenant.fullName}`;
        await db.prepare(`
          INSERT INTO units (floor, label, housing_type, notes)
          VALUES (?, ?, ?, 'Logement historique ; aucun numéro de logement dans la source.')
          ON CONFLICT (floor, label) DO NOTHING
        `).run(tenant.floor, unitLabel, tenant.housingType || null);
        const unit = await db.prepare('SELECT id FROM units WHERE floor = ? AND label = ?')
          .get(tenant.floor, unitLabel);
        const result = await db.prepare(`
          INSERT INTO tenants (
            unit_id, first_name, last_name, people_count, monthly_rent, entry_date, is_active
          ) VALUES (?, ?, ?, ?, ?, NULL, TRUE)
        `).run(unit.id, tenant.firstName, tenant.lastName, tenant.peopleCount, tenant.monthlyRent);
        existingTenant = { id: result.lastInsertRowid };
      }
      tenantIds.set(tenant.sourceKey, Number(existingTenant.id));
    }

    for (const payment of rentPayments) {
      const tenantId = tenantIds.get(payment.tenantKey);
      if (!tenantId) throw new Error(`Locataire introuvable pour le loyer de ${payment.tenantName}.`);
      const preexistingObligation = await db.prepare(`
        SELECT id FROM rent_obligations WHERE tenant_id = ? AND month_key = ?
      `).get(tenantId, payment.monthKey);
      await db.prepare(`
        INSERT INTO rent_obligations (tenant_id, month_key, expected_amount)
        VALUES (?, ?, ?)
        ON CONFLICT (tenant_id, month_key) DO NOTHING
      `).run(tenantId, payment.monthKey, payment.expectedAmount);
      const obligation = await db.prepare(`
        SELECT id, expected_amount AS "expectedAmount" FROM rent_obligations
        WHERE tenant_id = ? AND month_key = ?
      `).get(tenantId, payment.monthKey);
      if (Number(obligation.expectedAmount) !== payment.expectedAmount) {
        throw new Error(`Le loyer attendu de ${payment.tenantName} diffère déjà pour ${payment.monthKey}.`);
      }
      const existingPayment = await db.prepare(`
        SELECT COALESCE(SUM(amount), 0) AS total FROM rent_payments WHERE obligation_id = ?
      `).get(obligation.id);
      const paidAmount = payment.paidAmount || 0;
      if (Number(existingPayment.total) > 0) {
        if (Math.abs(Number(existingPayment.total) - paidAmount) > 1e-9) {
          throw new Error(`Le paiement de ${payment.tenantName} diffère déjà pour ${payment.monthKey}.`);
        }
        payment.duplicateSource = 'base';
      } else if (paidAmount > 0) {
        await db.prepare(`
          INSERT INTO rent_payments (
            obligation_id, amount, payment_date, received_month_key, comment, created_by
          ) VALUES (?, ?, NULL, ?, ?, ?)
        `).run(
          obligation.id,
          paidAmount,
          payment.receivedMonthKey || null,
          `Statut source : ${payment.statusText || payment.status} (${payment.sheetName}, ligne ${payment.rowNumber}). Date absente du classeur.`,
          req.user.id,
        );
      } else if (preexistingObligation) {
        payment.duplicateSource = 'base';
      }
    }

    for (const income of incomes) {
      const existingIncome = await db.prepare(`
        SELECT id FROM incomes
        WHERE month_key = ? AND lower(title) = lower(?) AND lower(source) = lower(?) AND amount = ?
        LIMIT 1
      `).get(income.monthKey, income.title, income.source, income.amount);
      if (existingIncome) {
        income.duplicateSource = 'base';
        continue;
      }
      await db.prepare(`
        INSERT INTO income_categories (name) VALUES (?)
        ON CONFLICT (name) DO NOTHING
      `).run(income.categoryName || income.source);
      const category = await db.prepare('SELECT id FROM income_categories WHERE name = ?')
        .get(income.categoryName || income.source);
      await db.prepare(`
        INSERT INTO incomes (
          income_date, title, category_id, amount, source, payment_method, month_key, comment, created_by
        ) VALUES (NULL, ?, ?, ?, ?, NULL, ?, ?, ?)
      `).run(
        income.title,
        category.id,
        income.amount,
        income.source,
        income.monthKey,
        `Importée de ${income.sheetName}, ligne ${income.rowNumber}. Date absente du classeur.`,
        req.user.id,
      );
    }

    for (const saving of savings) {
      if (saving.valueType !== 'balance'
        || !/^\d{4}-(0[1-9]|1[0-2])$/.test(saving.monthKey || '')) {
        throw new Error(`Le solde d’épargne de ${saving.name} n’a pas de mois valide.`);
      }
      let account;
      if (normalizeIncomeSource(saving.name) === 'epargne pension pere') {
        account = await db.prepare(`
          SELECT id FROM savings_accounts WHERE is_protected = TRUE ORDER BY id LIMIT 1
        `).get();
      }
      if (!account) {
        await db.prepare(`
          INSERT INTO savings_accounts (name) VALUES (?)
          ON CONFLICT (name) DO NOTHING
        `).run(saving.name);
        account = await db.prepare('SELECT id FROM savings_accounts WHERE name = ?').get(saving.name);
      }
      const existingSnapshot = await db.prepare(`
        SELECT amount FROM savings_balance_snapshots WHERE account_id = ? AND month_key = ?
      `).get(account.id, saving.monthKey);
      if (existingSnapshot) {
        if (Number(existingSnapshot.amount) !== saving.amount) {
          throw new Error(`Le solde d’épargne de ${saving.name} diffère déjà pour ${saving.monthKey}.`);
        }
        saving.duplicateSource = 'base';
        continue;
      }
      await db.prepare(`
        INSERT INTO savings_balance_snapshots (
          account_id, month_key, amount, source_file_name, source_sheet, source_row
        ) VALUES (?, ?, ?, ?, ?, ?)
      `).run(
        account.id,
        saving.monthKey,
        saving.amount,
        previewBatch.file_name,
        saving.sheetName,
        saving.rowNumber,
      );
      const verifiedSnapshot = await db.prepare(`
        SELECT amount FROM savings_balance_snapshots WHERE account_id = ? AND month_key = ?
      `).get(account.id, saving.monthKey);
      if (Number(verifiedSnapshot?.amount) !== saving.amount) {
        throw new Error(`Le contrôle du solde d’épargne de ${saving.name} ne correspond pas pour ${saving.monthKey}.`);
      }
    }

    const expectedExpenses = toImport.filter((row) => row.type === 'expense');
    const expectedGroceries = toImport.filter((row) => row.type === 'grocery');
    const storedExpenses = [];
    for (const id of expenseIds) {
      storedExpenses.push(await db.prepare(`
        SELECT e.title, e.expense_date AS date, e.amount, e.month_key AS "monthKey",
          e.description, e.reference, e.beneficiary, e.paid_by AS "paidBy",
          e.payment_method AS "paymentMethod", c.name AS category
        FROM expenses e LEFT JOIN categories c ON c.id = e.category_id WHERE e.id = ?
      `).get(id));
    }
    const storedGroceries = [];
    for (const id of groceryIds) {
      storedGroceries.push(await db.prepare(`
        SELECT item_name AS title, purchase_date AS date, total_amount AS amount,
          month_key AS "monthKey",
          quantity, unit, unit_price AS "unitPrice", quantity_known AS "quantityKnown",
          unit_price_known AS "unitPriceKnown"
        FROM grocery_items WHERE id = ?
      `).get(id));
    }
    const expenseTotal = storedExpenses.reduce((sum, row) => sum + Number(row.amount), 0);
    const groceryTotal = storedGroceries.reduce((sum, row) => sum + Number(row.amount), 0);
    const rowsMatch = (expected, stored, type) => expected.length === stored.length
      && expected.every((row, index) => {
        const actual = stored[index];
        const titleMatches = row.title === actual.title;
        const dateMatches = row.date === actual.date;
        const monthMatches = (row.monthKey || row.date?.slice(0, 7)) === actual.monthKey;
        const amountMatches = Math.abs(row.amount - Number(actual.amount)) <= 1e-9;
        const categoryMatches = type !== 'expense' || row.categoryName === actual.category;
        const detailsMatch = type !== 'expense' || (
          (row.monthKey || row.date?.slice(0, 7) || null) === (actual.monthKey || null)
          && (row.notes || '') === (actual.description || '')
          && (row.reference || '') === (actual.reference || '')
          && (row.beneficiary || '') === (actual.beneficiary || '')
          && (row.paidBy || '') === (actual.paidBy || '')
          && (row.paymentMethod || '') === (actual.paymentMethod || '')
        );
        const quantityMatches = type !== 'grocery'
          || (row.quantityKnown === Boolean(actual.quantityKnown)
            && row.unitPriceKnown === Boolean(actual.unitPriceKnown)
            && (!row.quantityKnown || row.quantity === actual.quantity)
            && (!row.unitPriceKnown || row.unitPrice === actual.unitPrice)
            && (row.unit || '') === (actual.unit || ''));
        return titleMatches && dateMatches && monthMatches && amountMatches
          && categoryMatches && detailsMatch && quantityMatches;
      });
    if (!rowsMatch(expectedExpenses, storedExpenses, 'expense')
      || !rowsMatch(expectedGroceries, storedGroceries, 'grocery')
      || Math.abs(expenseTotal - expectedExpenses.reduce((sum, row) => sum + row.amount, 0)) > 1e-9
      || Math.abs(groceryTotal - expectedGroceries.reduce((sum, row) => sum + row.amount, 0)) > 1e-9) {
      throw new Error('Les contrôles après insertion ne correspondent pas aux lignes prévisualisées.');
    }
    const expectedMonthly = summarizeRecords(toImport);
    const actualMonthly = new Map();
    [...storedExpenses.map((row) => ({ ...row, type: 'expense' })),
      ...storedGroceries.map((row) => ({ ...row, type: 'grocery' }))].forEach((row) => {
      const monthKey = row.monthKey || row.date?.slice(0, 7);
      if (!monthKey) throw new Error('Une dépense importée ne contient pas de mois de rattachement.');
      const month = actualMonthly.get(monthKey) || {
        ordinaryCount: 0, groceryCount: 0, totalOther: 0, totalBoutique: 0,
      };
      if (row.type === 'expense') {
        month.ordinaryCount += 1;
        month.totalOther += Number(row.amount);
      } else {
        month.groceryCount += 1;
        month.totalBoutique += Number(row.amount);
      }
      actualMonthly.set(monthKey, month);
    });
    expectedMonthly.forEach((month) => {
      const actual = actualMonthly.get(month.month) || {
        ordinaryCount: 0, groceryCount: 0, totalOther: 0, totalBoutique: 0,
      };
      if (actual.ordinaryCount !== month.ordinaryCount
        || actual.groceryCount !== month.groceryCount
        || Math.abs(actual.totalOther - month.totalOther) > 1e-9
        || Math.abs(actual.totalBoutique - month.totalBoutique) > 1e-9
        || Math.abs(month.totalGeneral - (month.totalBoutique + month.totalOther)) > 1e-9) {
        throw new Error(`Le contrôle des totaux de ${month.label} ne correspond pas à la prévisualisation.`);
      }
    });

    if (reinitialize) {
      await db.prepare(`
        INSERT INTO import_batches (file_name, rows_json, status) VALUES (?, ?, 'imported')
      `).run(previewBatch.file_name, JSON.stringify({ ...payload, records, backup: backup.fileName }));
      await db.prepare(`
        INSERT INTO settings (key, value) VALUES ('historical_2026_reinitialized', ?)
        ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value
      `).run(new Date().toISOString());
    } else {
      await db.prepare('UPDATE import_batches SET rows_json = ?, status = ? WHERE id = ?')
        .run(JSON.stringify({ ...payload, records }), 'imported', previewBatch.id);
    }
  });

  try {
    await transaction();
  } catch (error) {
    console.error('Excel import transaction failed:', error.message);
    return res.status(500).send(`Importation annulée, aucune donnée n’a été modifiée : ${error.message}`);
  }
  res.redirect('/import');
});

app.get('/backup', (req, res) => {
  const backups = fileList(backupDir).filter((file) => /\.dump$/i.test(file));
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

  await db.prepare('UPDATE backup_runs SET storage = ?, metadata = ? WHERE file_name = ?')
    .run(r2Url ? 'cloudflare-r2' : 'local', JSON.stringify({
      path: backup.target,
      url: r2Url,
      integrity: backup.integrity,
      fileSize: backup.fileSize,
    }), backup.fileName);

  res.redirect('/backup');
});

app.post('/backup/restore/:fileName', async (req, res) => {
  const fileName = decodeURIComponent(req.params.fileName);
  if (path.basename(fileName) !== fileName || !/\.dump$/i.test(fileName)) {
    return res.status(400).send('Nom de sauvegarde invalide.');
  }
  const source = path.join(backupDir, fileName);
  if (!fs.existsSync(source)) {
    return res.status(404).send('Sauvegarde introuvable');
  }

  const args = ['--clean', '--if-exists', '--no-owner'];
  if (process.env.DATABASE_URL) args.push('--dbname', process.env.DATABASE_URL);
  else {
    args.push('--host', process.env.PGHOST || '/var/run/postgresql');
    if (process.env.PGPORT) args.push('--port', process.env.PGPORT);
    args.push('--username', process.env.PGUSER || process.env.USER);
    args.push('--dbname', process.env.PGDATABASE || 'app_depenses');
  }
  args.push(source);
  await execFileAsync('pg_restore', args);
  res.redirect('/backup');
});

async function financialExportSheets(year) {
  const firstMonth = `${year}-01`;
  const lastMonth = `${year}-12`;
  const startDate = `${year}-01-01`;
  const endDate = `${year + 1}-01-01`;
  const [tenants, rents, incomes, savings, expenses, savingsSnapshots, monthlyBalances] = await Promise.all([
    db.prepare(`
      SELECT u.floor, u.label, u.housing_type, t.first_name, t.last_name, t.people_count,
        t.monthly_rent, t.entry_date, t.phone, t.is_active
      FROM tenants t JOIN units u ON u.id = t.unit_id
      ORDER BY u.floor, u.label, t.last_name, t.first_name
    `).all(),
    db.prepare(`
      SELECT o.month_key, u.floor, u.label, t.first_name, t.last_name, o.expected_amount,
        p.payment_date, p.amount AS payment_amount, p.payment_method, p.comment
      FROM rent_obligations o
      JOIN tenants t ON t.id = o.tenant_id
      JOIN units u ON u.id = t.unit_id
      LEFT JOIN rent_payments p ON p.obligation_id = o.id
      WHERE o.month_key BETWEEN ? AND ?
      ORDER BY o.month_key, u.floor, u.label, t.last_name, t.first_name, p.payment_date, p.id
    `).all(firstMonth, lastMonth),
    db.prepare(`
      SELECT i.month_key, i.income_date, c.name AS category, i.title, i.amount,
        i.source, i.payment_method, i.comment
      FROM incomes i JOIN income_categories c ON c.id = i.category_id
      WHERE i.month_key BETWEEN ? AND ?
      ORDER BY i.month_key, i.income_date, i.id
    `).all(firstMonth, lastMonth),
    db.prepare(`
      SELECT t.transaction_date, a.name AS account, t.transaction_type, t.amount, t.comment,
        e.title AS related_expense_title
      FROM savings_transactions t JOIN savings_accounts a ON a.id = t.account_id
      LEFT JOIN expenses e ON e.id = t.related_expense_id
      WHERE t.transaction_date >= ? AND t.transaction_date < ?
      ORDER BY t.transaction_date, t.id
    `).all(startDate, endDate),
    db.prepare(`
      SELECT e.expense_date, COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) AS month_key,
        c.name AS category, e.title, e.amount, e.beneficiary, e.paid_by,
        e.payment_method, e.reference, e.description, e.funding_source
      FROM expenses e JOIN categories c ON c.id = e.category_id
      WHERE COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) BETWEEN ? AND ?
      ORDER BY month_key, e.expense_date, e.id
    `).all(firstMonth, lastMonth),
    db.prepare(`
      SELECT s.month_key, a.name AS account, s.amount, s.source_file_name,
        s.source_sheet, s.source_row
      FROM savings_balance_snapshots s JOIN savings_accounts a ON a.id = s.account_id
      WHERE s.month_key BETWEEN ? AND ?
      ORDER BY s.month_key, a.name
    `).all(firstMonth, lastMonth),
    Promise.all(Array.from({ length: 12 }, (_, index) => (
      getMonthlyFinance(getMonthKey(year, index + 1))
    ))),
  ]);
  return [
    {
      name: 'LOCATAIRES',
      headers: ['Étage', 'Logement', 'Type de logement', 'Prénom', 'Nom', 'Personnes', 'Loyer mensuel (FCFA)', 'Date d’entrée', 'Téléphone', 'Statut'],
      rows: tenants.map((row) => [row.floor, row.label, row.housing_type, row.first_name, row.last_name, row.people_count, Number(row.monthly_rent), row.entry_date, row.phone, row.is_active ? 'Actif' : 'Inactif']),
      dateColumns: [7],
      amountColumns: [6],
    },
    {
      name: 'LOYERS ET PAIEMENTS',
      headers: ['Mois', 'Étage', 'Logement', 'Prénom', 'Nom', 'Montant attendu (FCFA)', 'Date du paiement', 'Montant payé (FCFA)', 'Mode de paiement', 'Commentaire'],
      rows: rents.map((row) => [row.month_key, row.floor, row.label, row.first_name, row.last_name, Number(row.expected_amount), row.payment_date, row.payment_amount === null ? '' : Number(row.payment_amount), row.payment_method, row.comment]),
      dateColumns: [6],
      amountColumns: [5, 7],
    },
    {
      name: 'RECETTES',
      headers: ['Mois', 'Date', 'Catégorie', 'Désignation', 'Montant (FCFA)', 'Source', 'Mode de paiement', 'Commentaire'],
      rows: incomes.map((row) => [row.month_key, row.income_date, row.category, row.title, Number(row.amount), row.source, row.payment_method, row.comment]),
      dateColumns: [1],
      amountColumns: [4],
    },
    {
      name: 'ÉPARGNE',
      headers: ['Date', 'Compte', 'Type', 'Montant (FCFA)', 'Dépense liée', 'Commentaire'],
      rows: savings.map((row) => [
        row.transaction_date,
        row.account,
        row.related_expense_title ? 'Dépense protégée' : (row.transaction_type === 'deposit' ? 'Versement' : 'Retrait'),
        Number(row.amount),
        row.related_expense_title,
        row.comment,
      ]),
      dateColumns: [0],
      amountColumns: [3],
    },
    {
      name: 'SOLDES ÉPARGNE',
      headers: ['Mois', 'Compte', 'Solde observé (FCFA)', 'Fichier source', 'Feuille source', 'Ligne source'],
      rows: savingsSnapshots.map((row) => [
        row.month_key, row.account, Number(row.amount), row.source_file_name, row.source_sheet, row.source_row,
      ]),
      amountColumns: [2],
    },
    {
      name: 'DÉPENSES DÉTAILLÉES',
      headers: ['Date', 'Mois', 'Catégorie', 'Désignation', 'Montant (FCFA)', 'Source de financement', 'Bénéficiaire', 'Payé par', 'Mode de paiement', 'Référence', 'Commentaire'],
      rows: expenses.map((row) => [
        row.expense_date, row.month_key, row.category, row.title, Number(row.amount),
        row.funding_source === 'protected_savings' ? 'Épargne protégée' : 'Compte courant',
        row.beneficiary, row.paid_by, row.payment_method, row.reference, row.description,
      ]),
      dateColumns: [0],
      amountColumns: [4],
    },
    {
      name: 'DÉPENSES ÉPARGNE PROTÉGÉE',
      headers: ['Date', 'Mois', 'Catégorie', 'Désignation', 'Montant (FCFA)', 'Source de financement', 'Bénéficiaire', 'Payé par', 'Mode de paiement', 'Référence', 'Commentaire'],
      rows: expenses.filter((row) => row.funding_source === 'protected_savings').map((row) => [
        row.expense_date, row.month_key, row.category, row.title, Number(row.amount),
        'Épargne protégée', row.beneficiary, row.paid_by, row.payment_method, row.reference, row.description,
      ]),
      dateColumns: [0],
      amountColumns: [4],
    },
    {
      name: 'BILAN FINANCIER',
      headers: ['Mois', 'Solde reporté (FCFA)', 'Recettes (FCFA)', 'Ressources disponibles (FCFA)', 'Loyers attendus (FCFA)', 'Loyers encaissés (FCFA)', 'Impayés (FCFA)', 'Dépenses compte courant (FCFA)', 'Dépenses épargne protégée (FCFA)', 'Versements épargne (FCFA)', 'Retraits épargne (FCFA)', 'Épargne cumulée (FCFA)', 'Solde final (FCFA)'],
      rows: monthlyBalances.map((row) => [
        row.monthKey, row.openingBalance, row.incomeTotal, row.resourcesAvailable,
        row.rentExpected, row.rentReceived, row.rentOutstanding, row.totalExpenses,
        row.protectedSavingsExpenses, row.savingsDeposits, row.savingsWithdrawals,
        row.savingsBalance, row.closingBalance,
      ]),
      amountColumns: Array.from({ length: 12 }, (_, index) => index + 1),
    },
  ];
}

async function getReportedIncomeRows(startMonth, endMonth) {
  const [legacyRows, monthlyRows] = await Promise.all([
    db.prepare(`
      SELECT i.income_date, i.month_key, c.name AS category, i.title, i.amount,
        i.source, i.payment_method, i.comment
      FROM incomes i JOIN income_categories c ON c.id = i.category_id
      WHERE i.month_key BETWEEN ? AND ?
      ORDER BY i.month_key, i.income_date, i.id
    `).all(startMonth, endMonth),
    db.prepare(`
      SELECT * FROM monthly_revenues WHERE month_key BETWEEN ? AND ? ORDER BY month_key
    `).all(startMonth, endMonth),
  ]);
  const monthlyKeys = new Set(monthlyRows.map((row) => row.month_key));
  const reported = legacyRows.filter((income) => (
    !isRentIncomeSource(income.source)
    && (!monthlyKeys.has(income.month_key) || !isPrimaryIncomeSource(income.source))
  ));
  for (const month of monthlyRows) {
    if (Number(month.pape_amount) > 0) {
      reported.push({
        income_date: month.pape_received_date,
        month_key: month.month_key,
        category: 'Pape',
        title: 'Pape',
        amount: Number(month.pape_amount),
        source: 'Pape',
        payment_method: null,
        comment: month.comment,
      });
    }
    if (Number(month.pension_amount) > 0) {
      reported.push({
        income_date: month.pension_received_date,
        month_key: month.month_key,
        category: 'Pension père',
        title: 'Pension du père',
        amount: Number(month.pension_amount),
        source: 'Pension père',
        payment_method: null,
        comment: month.comment,
      });
    }
  }
  return reported.sort((left, right) => (
    left.month_key.localeCompare(right.month_key)
    || String(left.income_date || '').localeCompare(String(right.income_date || ''))
  ));
}

async function exportPeriodRows(monthKey) {
  const [expenses, groceries, incomes, rents, savings, balanceSnapshots] = await Promise.all([
    db.prepare(`
      SELECT e.expense_date, e.month_key, c.name AS category, e.title, e.amount,
        e.beneficiary, e.paid_by, e.payment_method, e.reference, e.description, e.funding_source
      FROM expenses e JOIN categories c ON c.id = e.category_id
      WHERE COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) = ?
      ORDER BY e.expense_date, e.id
    `).all(monthKey),
    db.prepare(`
      SELECT item_name, purchase_date, month_key, quantity, unit, unit_price, total_amount,
        quantity_known AS "quantityKnown", unit_price_known AS "unitPriceKnown", notes
      FROM grocery_items
      WHERE COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM')) = ?
      ORDER BY purchase_date NULLS LAST, id
    `).all(monthKey),
    getReportedIncomeRows(monthKey, monthKey),
    db.prepare(`
      SELECT o.month_key, u.floor, u.label, t.first_name, t.last_name, o.expected_amount,
        p.payment_date, p.amount AS payment_amount, p.payment_method, p.comment
      FROM rent_obligations o
      JOIN tenants t ON t.id = o.tenant_id
      JOIN units u ON u.id = t.unit_id
      LEFT JOIN rent_payments p ON p.obligation_id = o.id
      WHERE o.month_key = ? ORDER BY u.floor, u.label, t.last_name, t.first_name, p.payment_date, p.id
    `).all(monthKey),
    db.prepare(`
      SELECT t.transaction_date, a.name AS account, t.transaction_type, t.amount, t.comment,
        e.title AS related_expense_title
      FROM savings_transactions t JOIN savings_accounts a ON a.id = t.account_id
      LEFT JOIN expenses e ON e.id = t.related_expense_id
      WHERE to_char(t.transaction_date, 'YYYY-MM') = ? ORDER BY t.transaction_date, t.id
    `).all(monthKey),
    db.prepare(`
      SELECT s.month_key, a.name AS account, s.amount, s.source_file_name,
        s.source_sheet, s.source_row
      FROM savings_balance_snapshots s JOIN savings_accounts a ON a.id = s.account_id
      WHERE s.month_key = ? ORDER BY a.name
    `).all(monthKey),
  ]);
  return { expenses, groceries, incomes, rents, savings, balanceSnapshots };
}

function deliverWorkbook(res, workbook, fileName) {
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
}

async function getReferenceMonthData(year, month) {
  const monthDate = new Date(Date.UTC(year, month - 1, 1));
  const monthKey = `${monthDate.getUTCFullYear()}-${String(monthDate.getUTCMonth() + 1).padStart(2, '0')}`;
  const previousDate = new Date(Date.UTC(year, month - 2, 1));
  const previousMonthKey = `${previousDate.getUTCFullYear()}-${String(previousDate.getUTCMonth() + 1).padStart(2, '0')}`;
  const monthLabel = MONTH_NAMES[month - 1];
  const previousMonthLabel = `${MONTH_NAMES[previousDate.getUTCMonth()]}${previousDate.getUTCFullYear() === year ? '' : ` ${previousDate.getUTCFullYear()}`}`;
  const [finance, tenants, incomes, monthlyRevenue, rentPayments, expenses, groceries] = await Promise.all([
    getMonthlyFinance(monthKey),
    db.prepare(`
      SELECT t.id, u.floor, u.label, u.housing_type, t.first_name, t.last_name,
        t.people_count,
        COALESCE(o.expected_amount, CASE
          WHEN t.entry_date IS NULL OR t.entry_date < (? || '-01')::date + INTERVAL '1 month'
          THEN t.monthly_rent ELSE 0
        END) AS expected_amount,
        COALESCE(p.paid_amount, 0) AS paid_amount
      FROM tenants t
      JOIN units u ON u.id = t.unit_id
      LEFT JOIN rent_obligations o ON o.tenant_id = t.id AND o.month_key = ?
      LEFT JOIN LATERAL (
        SELECT SUM(amount) AS paid_amount FROM rent_payments WHERE obligation_id = o.id
      ) p ON TRUE
      WHERE t.is_active = TRUE
      ORDER BY CASE u.floor
        WHEN 'Rez-de-chaussée' THEN 0
        WHEN '1er étage' THEN 1
        WHEN '2e étage' THEN 2
        WHEN '3e étage' THEN 3
        ELSE 4 END, u.label, t.last_name, t.first_name
    `).all(previousMonthKey, previousMonthKey),
    db.prepare(`
      SELECT source, SUM(amount) AS amount
      FROM incomes WHERE month_key = ?
      GROUP BY source ORDER BY source
    `).all(monthKey),
    db.prepare('SELECT * FROM monthly_revenues WHERE month_key = ?').get(monthKey),
    db.prepare(`
      SELECT p.amount
      FROM rent_payments p
      JOIN rent_obligations o ON o.id = p.obligation_id
      WHERE COALESCE(NULLIF(p.received_month_key, ''), to_char(p.payment_date, 'YYYY-MM'), o.month_key) = ?
      ORDER BY p.id
    `).all(monthKey),
    db.prepare(`
      SELECT e.title, e.amount, e.funding_source, c.name AS category
      FROM expenses e
      LEFT JOIN categories c ON c.id = e.category_id
      WHERE COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) = ?
      ORDER BY e.expense_date NULLS LAST, e.id
    `).all(monthKey),
    db.prepare(`
      SELECT item_name AS title, total_amount AS amount
      FROM grocery_items
      WHERE COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM')) = ?
      ORDER BY purchase_date NULLS LAST, id
    `).all(monthKey),
  ]);
  const legacyPension = incomes
    .filter((income) => normalizeIncomeSource(income.source) === 'pension pere')
    .reduce((sum, income) => sum + Number(income.amount), 0);
  const legacyPape = incomes
    .filter((income) => normalizeIncomeSource(income.source) === 'pape')
    .reduce((sum, income) => sum + Number(income.amount), 0);
  const pensionPere = Number(monthlyRevenue?.pension_amount ?? legacyPension);
  const pape = Number(monthlyRevenue?.pape_amount ?? legacyPape);
  const otherIncomes = incomes
    .filter((income) => (
      !['pension pere', 'pape'].includes(normalizeIncomeSource(income.source))
      && !isRentIncomeSource(income.source)
    ))
    .map((income) => ({ source: income.source, amount: Number(income.amount) }));
  const rentReceived = rentPayments.reduce((sum, payment) => sum + Number(payment.amount), 0);
  const protectedExpenseRows = expenses
    .filter((row) => row.funding_source === 'protected_savings')
    .map((row) => ({ title: row.title, amount: Number(row.amount), category: row.category }));
  const expenseRows = [
    ...expenses.filter((row) => row.funding_source !== 'protected_savings')
      .map((row) => ({ title: row.title, amount: Number(row.amount), category: row.category })),
    ...groceries.map((row) => ({
      title: row.title,
      amount: Number(row.amount),
      category: 'Compte boutique',
    })),
  ];
  const tenantsForTemplate = tenants.map((tenant) => {
    const expectedAmount = Number(tenant.expected_amount);
    const paidAmount = Number(tenant.paid_amount);
    const status = expectedAmount <= 0
      ? 'Non dû'
      : paidAmount >= expectedAmount
        ? 'Payé'
        : paidAmount > 0
          ? 'Partiellement payé'
          : 'Non payé';
    return {
      floor: tenant.floor,
      name: `${tenant.first_name} ${tenant.last_name}`,
      housingType: tenant.housing_type,
      peopleCount: Number(tenant.people_count),
      expectedAmount,
      status,
    };
  });
  const sum = (rows) => rows.reduce((total, row) => total + Number(row.amount), 0);
  if (Math.abs(sum(otherIncomes) + pensionPere + pape - Number(finance.ordinaryIncome || 0)) > 0.01
    || Math.abs(rentReceived - Number(finance.rentReceived || 0)) > 0.01
    || Math.abs(sum(expenseRows) - Number(finance.totalExpenses || 0)) > 0.01) {
    throw new Error(`Les totaux de ${monthKey} ne correspondent pas aux opérations de la base.`);
  }

  return {
    monthKey,
    monthLabel,
    previousMonthLabel,
    tenants: tenantsForTemplate,
    totalPeople: tenantsForTemplate.reduce((total, tenant) => total + tenant.peopleCount, 0),
    rentExpected: tenantsForTemplate.reduce((total, tenant) => total + tenant.expectedAmount, 0),
    rentPayments: rentPayments.map((payment) => ({ amount: Number(payment.amount) })),
    rentReceived,
    pensionPere,
    pape,
    otherIncomes,
    incomeTotal: Number(finance.incomeTotal),
    openingBalance: Number(finance.openingBalance),
    resourcesAvailable: Number(finance.resourcesAvailable),
    savingsBalance: Number(finance.protectedSavingsBalance),
    savingsDeposits: Number(finance.savingsDeposits),
    savingsWithdrawals: Number(finance.savingsWithdrawals),
    closingBalance: Number(finance.closingBalance),
    expenses: expenseRows,
    protectedExpenses: protectedExpenseRows,
    totalExpenses: Number(finance.totalExpenses),
  };
}

async function getMonthlyExpenseExportData(year, month, query) {
  const monthKey = getMonthKey(year, month);
  let where = `
    WHERE e.funding_source <> 'protected_savings'
      AND COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) = ?
  `;
  const params = [monthKey];
  if (/^\d+$/.test(String(query.category || ''))) {
    where += ' AND e.category_id = ?';
    params.push(Number(query.category));
  }
  if (query.category_alias) {
    const alias = String(query.category_alias);
    const allCategories = await readCategories();
    const matchingIds = allCategories
      .filter((category) => resolvePriorityCategory(category.name) === alias
        || (alias === 'Compte boutique' && category.type === 'boutique')
        || (resolvePriorityCategory(category.name) === null && category.name === alias))
      .map((category) => Number(category.id));
    if (matchingIds.length) {
      where += ` AND e.category_id IN (${matchingIds.map(() => '?').join(', ')})`;
      params.push(...matchingIds);
    } else {
      where += ' AND 1 = 0';
    }
  }
  const search = String(query.search || '').trim();
  if (search) {
    where += `
      AND (e.title ILIKE ? OR e.description ILIKE ? OR e.reference ILIKE ?
        OR e.beneficiary ILIKE ? OR e.paid_by ILIKE ? OR c.name ILIKE ?)
    `;
    params.push(...Array(6).fill(`%${search}%`));
  }
  const headers = [
    'Date', 'Mois', 'Désignation', 'Catégorie', 'Montant', 'Source de financement',
    'Bénéficiaire', 'Personne ayant payé', 'Mode de paiement', 'Référence', 'Commentaire',
  ];
  const rows = (await db.prepare(`
    SELECT e.expense_date, COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) AS month_key,
      e.title, c.name AS category, e.amount, e.funding_source, e.beneficiary, e.paid_by,
      e.payment_method, e.reference, e.description
    FROM expenses e JOIN categories c ON c.id = e.category_id
    ${where}
    ORDER BY e.expense_date, e.id
  `).all(...params)).map((row) => [
    row.expense_date, row.month_key, row.title, row.category, Number(row.amount),
    row.funding_source === 'protected_savings' ? 'Épargne protégée' : 'Compte courant',
    row.beneficiary, row.paid_by, row.payment_method, row.reference, row.description,
  ]);
  return { monthKey, headers, rows };
}

app.get('/export/csv/month/:year/:month/expenses', async (req, res) => {
  const year = Number(req.params.year);
  const month = Number(req.params.month);
  if (!Number.isInteger(year) || year < 2000 || year > 2100 || !Number.isInteger(month) || month < 1 || month > 12) {
    return res.status(400).send('Période d’exportation invalide.');
  }
  const { monthKey, headers, rows } = await getMonthlyExpenseExportData(year, month, req.query);
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="MAISON_DEPENSES_${monthKey}.csv"`);
  res.send(makeCsv(headers, rows));
});

app.get('/export/xlsx/month/:year/:month/expenses', async (req, res) => {
  const year = Number(req.params.year);
  const month = Number(req.params.month);
  if (!Number.isInteger(year) || year < 2000 || year > 2100 || !Number.isInteger(month) || month < 1 || month > 12) {
    return res.status(400).send('Période d’exportation invalide.');
  }
  const { monthKey, headers, rows } = await getMonthlyExpenseExportData(year, month, req.query);
  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  worksheet['!cols'] = [
    { wch: 14 }, { wch: 12 }, { wch: 34 }, { wch: 27 }, { wch: 17 },
    { wch: 22 }, { wch: 24 }, { wch: 24 }, { wch: 20 }, { wch: 22 }, { wch: 42 },
  ];
  worksheet['!autofilter'] = { ref: `A1:K${Math.max(1, rows.length + 1)}` };
  worksheet['!rows'] = [{ hpt: 24 }];
  headers.forEach((_, columnIndex) => {
    const headerCell = worksheet[XLSX.utils.encode_cell({ r: 0, c: columnIndex })];
    if (headerCell) {
      headerCell.s = {
        fill: { fgColor: { rgb: '147967' } },
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        alignment: { vertical: 'center', wrapText: true },
      };
    }
  });
  rows.forEach((_, rowIndex) => {
    const amountCell = worksheet[XLSX.utils.encode_cell({ r: rowIndex + 1, c: 4 })];
    if (amountCell) amountCell.z = '#,##0 "FCFA"';
  });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Dépenses');
  deliverWorkbook(res, workbook, `MAISON_DEPENSES_${monthKey}.xlsx`);
});

app.get('/export/csv/:year/:kind', async (req, res) => {
  const year = Number(req.params.year);
  if (!Number.isInteger(year) || year < 2000 || year > 2100) return res.status(400).send('Année d’exportation invalide.');
  const startMonth = `${year}-01`;
  const endMonth = `${year}-12`;
  const kind = req.params.kind;
  let headers;
  let rows;
  if (kind === 'expenses') {
    headers = ['Date', 'Mois', 'Désignation', 'Catégorie', 'Montant', 'Source de financement', 'Bénéficiaire', 'Personne ayant payé', 'Mode de paiement', 'Référence', 'Commentaire'];
    rows = (await db.prepare(`
      SELECT e.expense_date, COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) AS month_key,
        e.title, c.name AS category, e.amount, e.funding_source, e.beneficiary, e.paid_by,
        e.payment_method, e.reference, e.description
      FROM expenses e JOIN categories c ON c.id = e.category_id
      WHERE c.type <> 'boutique'
        AND COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) BETWEEN ? AND ?
      ORDER BY month_key, e.expense_date, e.id
    `).all(startMonth, endMonth)).map((row) => [
      row.expense_date, row.month_key, row.title, row.category, Number(row.amount),
      row.funding_source === 'protected_savings' ? 'Épargne protégée' : 'Compte courant',
      row.beneficiary, row.paid_by, row.payment_method, row.reference, row.description,
    ]);
  } else if (kind === 'groceries') {
    headers = ['Date', 'Catégorie', 'Marchandise', 'Quantité', 'Unité', 'Prix unitaire', 'Montant', 'Commentaire'];
    rows = (await db.prepare(`
      SELECT purchase_date, month_key, item_name, quantity, unit, unit_price, total_amount, notes,
        quantity_known AS "quantityKnown", unit_price_known AS "unitPriceKnown"
      FROM grocery_items
      WHERE COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM')) BETWEEN ? AND ?
      ORDER BY month_key, purchase_date, id
    `).all(`${year}-01`, `${year}-12`)).map((row) => [
      row.purchase_date || row.month_key, 'Boutique', row.item_name,
      row.quantityKnown ? Number(row.quantity) : '', row.unit,
      row.unitPriceKnown ? Number(row.unit_price) : '', Number(row.total_amount), row.notes,
    ]);
  } else if (kind === 'incomes') {
    headers = ['Date', 'Mois', 'Désignation', 'Catégorie', 'Montant', 'Source', 'Mode de paiement', 'Commentaire'];
    rows = (await getReportedIncomeRows(startMonth, endMonth)).map((row) => [
      row.income_date,
      row.month_key,
      row.title,
      row.category,
      Number(row.amount),
      row.source,
      row.payment_method,
      row.comment,
    ]);
  } else if (kind === 'rents') {
    headers = ['Mois', 'Étage', 'Logement', 'Locataire', 'Montant attendu', 'Date du paiement', 'Montant payé', 'Mode de paiement', 'Commentaire'];
    rows = (await db.prepare(`
      SELECT o.month_key, u.floor, u.label, t.first_name || ' ' || t.last_name AS tenant,
        o.expected_amount, p.payment_date, p.amount AS payment_amount, p.payment_method, p.comment
      FROM rent_obligations o JOIN tenants t ON t.id = o.tenant_id
      JOIN units u ON u.id = t.unit_id
      LEFT JOIN rent_payments p ON p.obligation_id = o.id
      WHERE o.month_key BETWEEN ? AND ?
      ORDER BY o.month_key, u.floor, u.label, t.last_name, t.first_name, p.payment_date, p.id
    `).all(startMonth, endMonth)).map((row) => [row.month_key, row.floor, row.label, row.tenant, Number(row.expected_amount), row.payment_date, row.payment_amount === null ? '' : Number(row.payment_amount), row.payment_method, row.comment]);
  } else {
    return res.status(404).send('Type d’export CSV inconnu.');
  }
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="MAISON_${kind.toUpperCase()}_${year}.csv"`);
  res.send(makeCsv(headers, rows));
});

app.get('/export/month/:year/:month', async (req, res) => {
  const year = Number(req.params.year);
  const month = Number(req.params.month);
  if (!Number.isInteger(year) || year < 2000 || year > 2100 || !Number.isInteger(month) || month < 1 || month > 12) {
    return res.status(400).send('Période d’exportation invalide.');
  }
  const monthKey = getMonthKey(year, month);
  const data = await getReferenceMonthData(year, month);
  const workbook = await buildReferenceWorkbook([data], {
    includeCarryover: includeCarryoverInExport(req.query.includeCarryover),
  });
  const buffer = await workbook.xlsx.writeBuffer();
  res.setHeader('Content-Disposition', `attachment; filename="DEPENSES_MAISON_${monthKey}.xlsx"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(Buffer.from(buffer));
});

app.get('/export/month', async (req, res) => {
  const year = Number(req.query.year);
  const month = Number(req.query.month);
  if (!Number.isInteger(year) || year < 2000 || year > 2100 || !Number.isInteger(month) || month < 1 || month > 12) {
    return res.status(400).send('Période d’exportation invalide.');
  }
  const monthKey = getMonthKey(year, month);
  const data = await getReferenceMonthData(year, month);
  const workbook = await buildReferenceWorkbook([data], {
    includeCarryover: includeCarryoverInExport(req.query.includeCarryover),
  });
  const buffer = await workbook.xlsx.writeBuffer();
  res.setHeader('Content-Disposition', `attachment; filename="DEPENSES_MAISON_${monthKey}.xlsx"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(Buffer.from(buffer));
});

app.get('/export/months', async (req, res) => {
  const year = Number(req.query.year);
  const rawMonths = Array.isArray(req.query.months) ? req.query.months : [req.query.months];
  const months = rawMonths.map(Number);
  if (!Number.isInteger(year) || year < 2000 || year > 2100 || months.length === 0
    || months.some((month) => !Number.isInteger(month) || month < 1 || month > 12)
    || new Set(months).size !== months.length) {
    return res.status(400).send('Sélection de périodes invalide.');
  }
  const monthData = await Promise.all(months.sort((a, b) => a - b).map((month) => getReferenceMonthData(year, month)));
  const workbook = await buildReferenceWorkbook(monthData, {
    includeCarryover: includeCarryoverInExport(req.query.includeCarryover),
  });
  const buffer = await workbook.xlsx.writeBuffer();
  res.setHeader('Content-Disposition', `attachment; filename="DEPENSES_MAISON_${year}_MOIS.xlsx"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(Buffer.from(buffer));
});

app.get('/export/pdf/month/:year/:month', async (req, res) => {
  const year = Number(req.params.year);
  const month = Number(req.params.month);
  if (!Number.isInteger(year) || year < 2000 || year > 2100 || !Number.isInteger(month) || month < 1 || month > 12) {
    return res.status(400).send('Période d’exportation invalide.');
  }
  const monthKey = getMonthKey(year, month);
  const [finance, period] = await Promise.all([getMonthlyFinance(monthKey), exportPeriodRows(monthKey)]);
  const doc = new PDFDocument({ size: 'A4', margin: 42 });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `attachment; filename="BILAN_MAISON_${monthKey}.pdf"`);
  doc.pipe(res);
  doc.fontSize(18).font('Helvetica-Bold').text(`Bilan financier · ${monthKey}`);
  doc.moveDown(0.6).fontSize(10).font('Helvetica');
  [
    ['Solde reporté', finance.openingBalance],
    ['Recettes encaissées', finance.incomeTotal],
    ['Ressources disponibles', finance.resourcesAvailable],
    ['Loyers attendus', finance.rentExpected],
    ['Loyers encaissés', finance.rentReceived],
    ['Loyers impayés', finance.rentOutstanding],
    ['Dépenses du compte courant', finance.totalExpenses],
    ['Dépenses financées par l’épargne protégée', finance.protectedSavingsExpenses],
    ['Épargne cumulée', finance.savingsBalance],
    ['Solde final', finance.closingBalance],
  ].forEach(([label, amount]) => doc.text(`${label} : ${formatCurrency(amount)}`));
  doc.moveDown().fontSize(13).font('Helvetica-Bold').text('Dépenses du compte courant');
  doc.fontSize(9).font('Helvetica');
  period.expenses.filter((row) => row.funding_source !== 'protected_savings')
    .forEach((row) => doc.text(`${formatDate(row.expense_date)} · ${row.category} · ${row.title} · ${formatCurrency(row.amount)}`));
  period.groceries.forEach((row) => doc.text(`${formatDate(row.purchase_date)} · Boutique · ${row.item_name} · ${formatCurrency(row.total_amount)}`));
  const protectedExpenses = period.expenses.filter((row) => row.funding_source === 'protected_savings');
  if (protectedExpenses.length) {
    doc.moveDown().fontSize(13).font('Helvetica-Bold').text('Dépenses financées par l’épargne protégée');
    doc.fontSize(9).font('Helvetica');
    protectedExpenses.forEach((row) => doc.text(`${formatDate(row.expense_date)} · ${row.category} · ${row.title} · ${formatCurrency(row.amount)}`));
  }
  doc.moveDown().fontSize(13).font('Helvetica-Bold').text('Recettes et loyers encaissés');
  doc.fontSize(9).font('Helvetica');
  period.incomes.forEach((row) => doc.text(`${formatDate(row.income_date)} · ${row.source} · ${row.title} · ${formatCurrency(row.amount)}`));
  period.rents.filter((row) => row.payment_date).forEach((row) => doc.text(`${formatDate(row.payment_date)} · Loyer ${row.first_name} ${row.last_name} · ${formatCurrency(row.payment_amount)}`));
  doc.end();
});

app.get('/export/pdf/year/:year', async (req, res) => {
  const year = Number(req.params.year);
  if (!Number.isInteger(year) || year < 2000 || year > 2100) return res.status(400).send('Année d’exportation invalide.');
  const reports = await Promise.all(Array.from({ length: 12 }, (_, index) => (
    getMonthlyFinance(getMonthKey(year, index + 1))
  )));
  const doc = new PDFDocument({ size: 'A4', margin: 42 });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `attachment; filename="BILAN_ANNUEL_MAISON_${year}.pdf"`);
  doc.pipe(res);
  doc.fontSize(18).font('Helvetica-Bold').text(`Bilan annuel · ${year}`);
  doc.moveDown().fontSize(9).font('Helvetica-Bold');
  doc.text('Mois             Recettes       Loyers attendus       Dépenses courantes       Épargne        Dépenses protégées       Solde final');
  doc.font('Helvetica');
  reports.forEach((row) => doc.text([
    row.monthKey,
    formatCurrency(row.incomeTotal),
    formatCurrency(row.rentExpected),
    formatCurrency(row.totalExpenses),
    formatCurrency(row.savingsBalance),
    formatCurrency(row.protectedSavingsExpenses),
    formatCurrency(row.closingBalance),
  ].join('   ')));
  doc.moveDown().font('Helvetica-Bold').text(`Solde final de l’année : ${formatCurrency(reports.at(-1)?.closingBalance || 0)}`);
  doc.end();
});

app.get('/export/:year', async (req, res) => {
  const year = Number(req.params.year);
  if (!Number.isInteger(year) || year < 2000 || year > 2100) return res.status(400).send('Année d’exportation invalide.');
  const months = await Promise.all(Array.from({ length: 12 }, (_, index) => (
    getReferenceMonthData(year, index + 1)
  )));
  const workbook = await buildReferenceWorkbook(months, {
    annual: true,
    includeCarryover: includeCarryoverInExport(req.query.includeCarryover),
  });
  const buffer = await workbook.xlsx.writeBuffer();
  res.setHeader('Content-Disposition', `attachment; filename="DEPENSES_MAISON_${year}.xlsx"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(Buffer.from(buffer));
});

if (require.main === module) {
  initialization.then(() => {
    app.listen(PORT, () => {
      console.log(`App running on http://localhost:${PORT}`);
    });
  }).catch((error) => {
    console.error('PostgreSQL initialization failed:', error);
    process.exitCode = 1;
  });
}

module.exports = { app, db, initializeDatabase, initialization, formatCurrency, formatDate };
