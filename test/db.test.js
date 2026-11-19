const test = require('node:test');
const assert = require('node:assert/strict');

delete process.env.DATABASE_URL;
process.env.PGHOST = '/var/run/postgresql';
process.env.PGUSER = 'root';
process.env.PGDATABASE = 'app_depenses_db_test';

const { db, initializeDatabase } = require('../src/db');

test('PostgreSQL schema initialization is repeatable and retains grocery detail flags', async () => {
  await initializeDatabase();
  await initializeDatabase();

  const columns = await db.prepare(`
    SELECT column_name FROM information_schema.columns
    WHERE table_schema = 'public' AND table_name = 'grocery_items'
  `).all();
  assert.ok(columns.some((column) => column.column_name === 'quantity_known'));
  assert.ok(columns.some((column) => column.column_name === 'unit_price_known'));
  const rentColumns = await db.prepare(`
    SELECT column_name FROM information_schema.columns
    WHERE table_schema = 'public' AND table_name = 'rent_payments'
  `).all();
  assert.ok(rentColumns.some((column) => column.column_name === 'received_month_key'));
  const revenueColumns = await db.prepare(`
    SELECT column_name FROM information_schema.columns
    WHERE table_schema = 'public' AND table_name = 'monthly_revenues'
  `).all();
  assert.ok(revenueColumns.some((column) => column.column_name === 'pape_received_date'));
  assert.ok(revenueColumns.some((column) => column.column_name === 'pension_received_date'));
  const protectedAccounts = await db.prepare(`
    SELECT id FROM savings_accounts WHERE is_protected = TRUE
  `).all();
  assert.equal(protectedAccounts.length, 1);

  const inserted = await db.prepare(`
    INSERT INTO grocery_items (item_name, purchase_date, quantity, unit_price, total_amount)
    VALUES ('Test postgres', '2026-01-05', 2, 2500, 5000)
  `).run();
  try {
    const grocery = await db.prepare(`
      SELECT item_name, quantity, unit_price, total_amount, quantity_known, unit_price_known
      FROM grocery_items WHERE id = ?
    `).get(inserted.lastInsertRowid);
    assert.deepEqual(grocery, {
      item_name: 'Test postgres',
      quantity: 2,
      unit_price: 2500,
      total_amount: 5000,
      quantity_known: 1,
      unit_price_known: 1,
    });
  } finally {
    await db.prepare('DELETE FROM grocery_items WHERE id = ?').run(inserted.lastInsertRowid);
  }
});

test('category sync merges legacy expenses into the Render category list', async () => {
  await initializeDatabase();
  await db.prepare("DELETE FROM settings WHERE key = 'render_categories_sync_v1'").run();
  const legacyCategory = await db.prepare(`
    INSERT INTO categories (name, type, color)
    VALUES ('Dépenses exceptionnelles ou imprévues', 'ordinary', '#bb7966')
    ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name
    RETURNING id
  `).get();
  const inserted = await db.prepare(`
    INSERT INTO expenses (title, category_id, amount, expense_date)
    VALUES ('Migration de catégorie test', ?, 1250, '2026-11-19')
  `).run(legacyCategory.id);

  try {
    await initializeDatabase();
    const target = await db.prepare(`
      SELECT c.name FROM expenses e JOIN categories c ON c.id = e.category_id
      WHERE e.id = ?
    `).get(inserted.lastInsertRowid);
    assert.equal(target.name, 'Dépenses exceptionnelles');
    assert.equal(await db.prepare(`
      SELECT id FROM categories WHERE name = 'Dépenses exceptionnelles ou imprévues'
    `).get(), undefined);
    const expectedCategories = [
      'achats alimentaires complémentaires',
      'Dépenses alimentaires',
      'Compte boutique',
      'Scolarité et frais éducatifs',
      'Hygiène et beauté',
      'Aumônes et dons',
      'entretien de la maison',
      'Réparations',
      'Santé',
      'Transport',
      'Alimentation des animaux',
      'Autres dépenses',
      'Dépenses exceptionnelles',
      'Depenses imprevues',
      'Factures',
    ];
    const categories = await db.prepare(`
      SELECT name FROM categories WHERE is_system = TRUE ORDER BY name
    `).all();
    assert.deepEqual(categories.map((category) => category.name).sort(), expectedCategories.sort());
  } finally {
    await db.prepare('DELETE FROM expenses WHERE id = ?').run(inserted.lastInsertRowid);
  }
});

test('reserved fallback category is restored even if it was previously excluded from seeding', async () => {
  await initializeDatabase();
  await db.prepare(`
    INSERT INTO category_seed_exclusions (name) VALUES ('Autres dépenses')
    ON CONFLICT (name) DO NOTHING
  `).run();

  await initializeDatabase();

  assert.ok(await db.prepare("SELECT id FROM categories WHERE name = 'Autres dépenses'").get());
  assert.equal(
    await db.prepare("SELECT name FROM category_seed_exclusions WHERE name = 'Autres dépenses'").get(),
    undefined,
  );
});

test('legacy Pape and pension records are backfilled once without deleting their history', async () => {
  await db.prepare("DELETE FROM settings WHERE key = 'monthly_revenues_backfill_v1'").run();
  await db.prepare(`
    INSERT INTO income_categories (name) VALUES ('Pape'), ('Pension père')
    ON CONFLICT (name) DO NOTHING
  `).run();
  const categories = await db.prepare(`
    SELECT id, name FROM income_categories WHERE name IN ('Pape', 'Pension père')
  `).all();
  const papeCategory = categories.find((category) => category.name === 'Pape');
  const pensionCategory = categories.find((category) => category.name === 'Pension père');
  const pape = await db.prepare(`
    INSERT INTO incomes (income_date, title, category_id, amount, source, month_key)
    VALUES ('2026-03-03', 'Pape historique', ?, 350000, 'Pape', '2026-03')
  `).run(papeCategory.id);
  const pension = await db.prepare(`
    INSERT INTO incomes (income_date, title, category_id, amount, source, month_key)
    VALUES ('2026-03-05', 'Pension historique', ?, 180000, 'Pension père', '2026-03')
  `).run(pensionCategory.id);

  try {
    await initializeDatabase();
    const revenue = await db.prepare(`
      SELECT * FROM monthly_revenues WHERE month_key = '2026-03'
    `).get();
    assert.equal(Number(revenue.pape_amount), 350000);
    assert.equal(revenue.pape_received_date, '2026-03-03');
    assert.equal(Number(revenue.pension_amount), 180000);
    assert.equal(revenue.pension_received_date, '2026-03-05');
    const retained = await db.prepare(`
      SELECT COUNT(*) AS count FROM incomes WHERE id IN (?, ?)
    `).get(pape.lastInsertRowid, pension.lastInsertRowid);
    assert.equal(Number(retained.count), 2);
  } finally {
    await db.prepare("DELETE FROM monthly_revenues WHERE month_key = '2026-03'").run();
    await db.prepare('DELETE FROM incomes WHERE id IN (?, ?)').run(pape.lastInsertRowid, pension.lastInsertRowid);
    await db.prepare(`
      UPDATE settings SET value = 'done:0' WHERE key = 'monthly_revenues_backfill_v1'
    `).run();
  }
});

test('imported rents move to the prior obligation month while preserving their cash month', async () => {
  const migrationDb = db;
  const initialize = initializeDatabase;
  await migrationDb.prepare("DELETE FROM settings WHERE key = 'rent_sheet_previous_month_v1'").run();
  const unit = await migrationDb.prepare(`
    INSERT INTO units (floor, label) VALUES ('Rez-de-chaussée', 'MIGRATION-TEST') RETURNING id
  `).get();
  const tenant = await migrationDb.prepare(`
    INSERT INTO tenants (unit_id, first_name, last_name, monthly_rent)
    VALUES (?, 'Migration', 'Test', 15000) RETURNING id
  `).get(unit.id);
  const obligation = await migrationDb.prepare(`
    INSERT INTO rent_obligations (tenant_id, month_key, expected_amount)
    VALUES (?, '2026-02', 15000) RETURNING id
  `).get(tenant.id);
  await migrationDb.prepare(`
    INSERT INTO rent_payments (obligation_id, amount, payment_date, comment)
    VALUES (?, 15000, NULL, 'Statut source : Payé (Février, ligne 3). Date absente du classeur.')
  `).run(obligation.id);

  try {
    await initialize();
    const updatedObligation = await migrationDb.prepare(`
      SELECT month_key FROM rent_obligations WHERE id = ?
    `).get(obligation.id);
    const payment = await migrationDb.prepare(`
      SELECT received_month_key FROM rent_payments WHERE obligation_id = ?
    `).get(obligation.id);
    assert.equal(updatedObligation.month_key, '2026-01');
    assert.equal(payment.received_month_key, '2026-02');
    assert.equal(Number((await migrationDb.prepare(`
      SELECT COUNT(*) AS count FROM audit_log
      WHERE entity_type = 'rent_obligation' AND entity_id = ? AND action = 'period_corrected'
    `).get(String(obligation.id))).count), 1);
  } finally {
    await migrationDb.prepare('DELETE FROM rent_payments WHERE obligation_id = ?').run(obligation.id);
    await migrationDb.prepare('DELETE FROM rent_obligations WHERE id = ?').run(obligation.id);
    await migrationDb.prepare('DELETE FROM tenants WHERE id = ?').run(tenant.id);
    await migrationDb.prepare('DELETE FROM units WHERE id = ?').run(unit.id);
    await migrationDb.prepare(`
      DELETE FROM audit_log WHERE entity_type = 'rent_obligation' AND entity_id = ?
    `).run(String(obligation.id));
    await migrationDb.prepare("UPDATE settings SET value = 'done:1' WHERE key = 'rent_sheet_previous_month_v1'").run();
  }
});

test.after(async () => {
  await db.close();
});
