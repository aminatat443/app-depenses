const test = require('node:test');
const assert = require('node:assert/strict');
const { execFileSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const Database = require('better-sqlite3');

const initialize = (dbPath) => execFileSync(
  process.execPath,
  ['-e', "require('./src/db').initializeDatabase()"],
  { cwd: path.join(__dirname, '..'), env: { ...process.env, DB_PATH: dbPath } },
);

test('initializeDatabase adds the source column to legacy backup tables', () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'app-depenses-db-'));
  const dbPath = path.join(directory, 'app.db');
  const legacyDb = new Database(dbPath);
  legacyDb.exec(`
    CREATE TABLE backup_runs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      file_name TEXT NOT NULL,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP,
      storage TEXT NOT NULL DEFAULT 'local',
      status TEXT NOT NULL DEFAULT 'done',
      metadata TEXT
    );
    INSERT INTO backup_runs (file_name) VALUES ('existing.db');
  `);
  legacyDb.close();

  try {
    initialize(dbPath);
    const migratedDb = new Database(dbPath);
    try {
      const columns = migratedDb.prepare('PRAGMA table_info(backup_runs)').all();
      assert.ok(columns.some((column) => column.name === 'source'));
      assert.deepEqual(
        migratedDb.prepare('SELECT file_name, source FROM backup_runs').get(),
        { file_name: 'existing.db', source: 'manual' },
      );
    } finally {
      migratedDb.close();
    }
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

test('initializeDatabase preserves legacy groceries and marks their detail fields as known', () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'app-depenses-grocery-db-'));
  const dbPath = path.join(directory, 'app.db');
  const legacyDb = new Database(dbPath);
  legacyDb.exec(`
    CREATE TABLE grocery_items (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      item_name TEXT NOT NULL,
      purchase_date TEXT NOT NULL,
      quantity REAL NOT NULL,
      unit TEXT,
      unit_price REAL NOT NULL,
      total_amount REAL NOT NULL,
      notes TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );
    INSERT INTO grocery_items (item_name, purchase_date, quantity, unit_price, total_amount)
    VALUES ('Riz', '2026-01-05', 2, 2500, 5000);
  `);
  legacyDb.close();

  try {
    initialize(dbPath);
    const migratedDb = new Database(dbPath);
    try {
      const columns = migratedDb.prepare('PRAGMA table_info(grocery_items)').all();
      assert.ok(columns.some((column) => column.name === 'quantity_known'));
      assert.ok(columns.some((column) => column.name === 'unit_price_known'));
      assert.deepEqual(
        migratedDb.prepare('SELECT item_name, quantity, unit_price, total_amount, quantity_known, unit_price_known FROM grocery_items').get(),
        {
          item_name: 'Riz',
          quantity: 2,
          unit_price: 2500,
          total_amount: 5000,
          quantity_known: 1,
          unit_price_known: 1,
        },
      );
    } finally {
      migratedDb.close();
    }
  } finally {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});
