require('dotenv').config();
const Database = require('better-sqlite3');
const path = require('node:path');
const { db, initializeDatabase } = require('./db');

const sqlitePath = process.env.DB_PATH || path.join(__dirname, '..', 'data', 'app.db');
const tables = {
  categories: ['id', 'name', 'type', 'color'],
  expenses: ['id', 'title', 'category_id', 'amount', 'expense_date', 'description', 'reference', 'created_at'],
  grocery_items: [
    'id', 'item_name', 'purchase_date', 'quantity', 'unit', 'unit_price', 'total_amount',
    'notes', 'created_at', 'quantity_known', 'unit_price_known',
  ],
  import_batches: ['id', 'file_name', 'rows_json', 'status', 'created_at'],
  backup_runs: ['id', 'file_name', 'created_at', 'storage', 'status', 'metadata', 'source'],
  settings: ['key', 'value'],
};

async function migrate() {
  const source = new Database(sqlitePath, { readonly: true, fileMustExist: true });
  try {
    const integrity = source.pragma('integrity_check', { simple: true });
    const foreignKeyViolations = source.pragma('foreign_key_check');
    if (integrity !== 'ok' || foreignKeyViolations.length) {
      throw new Error('La base SQLite source ne passe pas les contrôles d’intégrité.');
    }

    await initializeDatabase();
    const sourceRows = Object.fromEntries(Object.entries(tables).map(([table, columns]) => {
      const order = table === 'settings' ? 'key' : 'id';
      return [table, source.prepare(`SELECT ${columns.join(', ')} FROM "${table}" ORDER BY "${order}"`).all()];
    }));

    await db.transaction(async () => {
      for (const [table, columns] of Object.entries(tables)) {
        const updates = columns
          .filter((column) => column !== 'id' && column !== 'key')
          .map((column) => `${column} = EXCLUDED.${column}`)
          .join(', ');
        const conflictColumn = table === 'settings' ? 'key' : 'id';
        const conflictAction = updates
          ? `DO UPDATE SET ${updates}`
          : 'DO NOTHING';
        const statement = `INSERT INTO ${table} (${columns.join(', ')}) VALUES (${columns.map(() => '?').join(', ')}) ON CONFLICT (${conflictColumn}) ${conflictAction}`;
        for (const row of sourceRows[table]) {
          await db.prepare(statement).run(...columns.map((column) => row[column]));
        }
      }

      for (const table of Object.keys(tables).filter((name) => name !== 'settings')) {
        await db.query(
          `SELECT setval(pg_get_serial_sequence(?, 'id'), COALESCE(MAX(id), 1), MAX(id) IS NOT NULL) FROM ${table}`,
          [table],
        );
      }
    });

    const results = {};
    for (const table of Object.keys(tables)) {
      const target = await db.prepare(`SELECT COUNT(*) AS count FROM ${table}`).get();
      results[table] = { source: sourceRows[table].length, target: Number(target.count) };
      if (results[table].source !== results[table].target) {
        throw new Error(`Le nombre de lignes de ${table} diffère après migration.`);
      }
    }
    const sourceTotals = {
      expenses: source.prepare('SELECT COALESCE(SUM(amount), 0) AS total FROM expenses').get().total,
      groceries: source.prepare('SELECT COALESCE(SUM(total_amount), 0) AS total FROM grocery_items').get().total,
    };
    const targetTotals = await db.prepare(`
      SELECT
        (SELECT COALESCE(SUM(amount), 0) FROM expenses) AS expenses,
        (SELECT COALESCE(SUM(total_amount), 0) FROM grocery_items) AS groceries
    `).get();
    for (const key of Object.keys(sourceTotals)) {
      if (Number(sourceTotals[key]) !== Number(targetTotals[key])) {
        throw new Error(`Le total de ${key} diffère après migration.`);
      }
    }

    console.log(JSON.stringify({ sqlitePath, rows: results, totals: targetTotals }, null, 2));
  } finally {
    source.close();
    await db.close();
  }
}

migrate().catch((error) => {
  console.error(`Migration SQLite vers PostgreSQL échouée: ${error.message}`);
  process.exitCode = 1;
});
