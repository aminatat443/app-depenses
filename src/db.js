const path = require('path');
const fs = require('fs');
const Database = require('better-sqlite3');

const dbPath = process.env.DB_PATH || path.join(__dirname, '..', 'data', 'app.db');
const dataDir = path.dirname(dbPath);
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

const db = new Database(dbPath);

db.pragma('journal_mode = WAL');
db.pragma('foreign_keys = ON');

function initializeDatabase() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS categories (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL UNIQUE,
      type TEXT NOT NULL DEFAULT 'ordinary',
      color TEXT DEFAULT '#7c3aed'
    );

    CREATE TABLE IF NOT EXISTS expenses (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      title TEXT NOT NULL,
      category_id INTEGER NOT NULL,
      amount REAL NOT NULL,
      expense_date TEXT NOT NULL,
      description TEXT,
      reference TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY(category_id) REFERENCES categories(id)
    );

    CREATE TABLE IF NOT EXISTS grocery_items (
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

    CREATE TABLE IF NOT EXISTS import_batches (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      file_name TEXT NOT NULL,
      rows_json TEXT NOT NULL,
      status TEXT NOT NULL DEFAULT 'pending',
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS backup_runs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      file_name TEXT NOT NULL,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP,
      storage TEXT NOT NULL DEFAULT 'local',
      status TEXT NOT NULL DEFAULT 'done',
      metadata TEXT,
      source TEXT DEFAULT 'manual'
    );

    CREATE TABLE IF NOT EXISTS settings (
      key TEXT PRIMARY KEY,
      value TEXT
    );
  `);

  const backupColumns = db.prepare('PRAGMA table_info(backup_runs)').all();
  if (!backupColumns.some((column) => column.name === 'source')) {
    db.exec("ALTER TABLE backup_runs ADD COLUMN source TEXT DEFAULT 'manual'");
  }

  const groceryColumns = db.prepare('PRAGMA table_info(grocery_items)').all();
  if (!groceryColumns.some((column) => column.name === 'quantity_known')) {
    db.exec('ALTER TABLE grocery_items ADD COLUMN quantity_known INTEGER NOT NULL DEFAULT 1');
  }
  if (!groceryColumns.some((column) => column.name === 'unit_price_known')) {
    db.exec('ALTER TABLE grocery_items ADD COLUMN unit_price_known INTEGER NOT NULL DEFAULT 1');
  }

  const categoryCount = db.prepare('SELECT COUNT(*) as count FROM categories').get();
  if (!categoryCount || categoryCount.count === 0) {
    const defaults = [
      ['Électricité', 'ordinary', '#22c55e'],
      ['Eau', 'ordinary', '#38bdf8'],
      ['Internet / Wi-Fi', 'ordinary', '#a78bfa'],
      ['Santé / Ordonnances', 'ordinary', '#f97316'],
      ['Transport', 'ordinary', '#facc15'],
      ['Entretien', 'ordinary', '#10b981'],
      ['Réparations', 'ordinary', '#f43f5e'],
      ['Scolarité', 'ordinary', '#fb7185'],
      ['Autres', 'ordinary', '#94a3b8'],
    ];

    const insertCategory = db.prepare('INSERT INTO categories (name, type, color) VALUES (?, ?, ?)');
    const insertMany = db.transaction((rows) => {
      for (const row of rows) insertCategory.run(row[0], row[1], row[2]);
    });
    insertMany(defaults);
  }
}

module.exports = {
  db,
  initializeDatabase,
  dbPath,
};
