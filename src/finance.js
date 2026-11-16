const express = require('express');
const { db } = require('./db');
const { calculateMonthlyBalance } = require('./monthly-balance');

const router = express.Router();
const FLOORS = ['Rez-de-chaussée', '1er étage', '2e étage', '3e étage'];
const MONTH_PATTERN = /^\d{4}-(0[1-9]|1[0-2])$/;
const normalizeSource = (value) => String(value || '')
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g, '')
  .toLocaleLowerCase('fr')
  .trim();
const isPrimaryRevenueSource = (value) => ['pape', 'pension pere'].includes(normalizeSource(value));
const isRentRevenueSource = (value) => /loyer|location|locataire/i.test(normalizeSource(value));
const currentMonth = () => new Date().toISOString().slice(0, 7);
const validMonth = (value) => MONTH_PATTERN.test(String(value || ''));
const validDate = (value) => {
  const date = String(value || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return false;
  const parsed = new Date(`${date}T00:00:00.000Z`);
  return !Number.isNaN(parsed.getTime()) && parsed.toISOString().slice(0, 10) === date;
};
const amountValue = (value) => {
  const amount = Number(value);
  return Number.isFinite(amount) && amount >= 0 ? amount : null;
};

async function writeAudit(userId, entityType, entityId, action, beforeData, afterData) {
  await db.prepare(`
    INSERT INTO audit_log (user_id, entity_type, entity_id, action, before_data, after_data)
    VALUES (?, ?, ?, ?, ?::jsonb, ?::jsonb)
  `).run(
    userId,
    entityType,
    String(entityId),
    action,
    beforeData ? JSON.stringify(beforeData) : null,
    afterData ? JSON.stringify(afterData) : null,
  );
}

async function getMonthlyFinance(monthKey) {
  const expenses = await db.prepare(`
    SELECT COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) AS month,
      SUM(CASE WHEN c.type = 'boutique' THEN e.amount ELSE 0 END) AS boutique,
      SUM(CASE WHEN c.type <> 'boutique' AND e.funding_source <> 'protected_savings'
        THEN e.amount ELSE 0 END) AS other,
      SUM(CASE WHEN c.type <> 'boutique' AND e.funding_source = 'protected_savings'
        THEN e.amount ELSE 0 END) AS protected_savings
    FROM expenses e JOIN categories c ON c.id = e.category_id
    WHERE COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) >= '2026-01'
      AND COALESCE(NULLIF(e.month_key, ''), to_char(e.expense_date, 'YYYY-MM')) <= ?
    GROUP BY month
  `).all(monthKey);
  const groceries = await db.prepare(`
    SELECT COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM')) AS month,
      SUM(total_amount) AS total
    FROM grocery_items
    WHERE COALESCE(NULLIF(month_key, ''), to_char(purchase_date, 'YYYY-MM'))
      BETWEEN '2026-01' AND ?
    GROUP BY month
  `).all(monthKey);
  const incomes = await db.prepare(`
    SELECT month_key AS month, source, SUM(amount) AS total
    FROM incomes
    WHERE month_key >= '2026-01' AND month_key <= ?
    GROUP BY month, source
  `).all(monthKey);
  const monthlyRevenues = await db.prepare(`
    SELECT month_key, pape_amount, pension_amount
    FROM monthly_revenues
    WHERE month_key >= '2026-01' AND month_key <= ?
  `).all(monthKey);
  const rentPayments = await db.prepare(`
    SELECT COALESCE(NULLIF(p.received_month_key, ''), to_char(p.payment_date, 'YYYY-MM'), o.month_key) AS month,
      SUM(p.amount) AS total
    FROM rent_payments p
    JOIN rent_obligations o ON o.id = p.obligation_id
    WHERE COALESCE(NULLIF(p.received_month_key, ''), to_char(p.payment_date, 'YYYY-MM'), o.month_key)
      BETWEEN '2026-01' AND ?
    GROUP BY month
  `).all(monthKey);
  const savings = await db.prepare(`
    SELECT to_char(t.transaction_date, 'YYYY-MM') AS month,
      SUM(CASE WHEN transaction_type = 'deposit' THEN amount ELSE 0 END) AS deposits,
      SUM(CASE WHEN transaction_type = 'withdrawal' AND NOT a.is_protected THEN amount ELSE 0 END) AS withdrawals
    FROM savings_transactions t
    JOIN savings_accounts a ON a.id = t.account_id
    WHERE t.transaction_date >= '2026-01-01' AND t.transaction_date < (? || '-01')::date + INTERVAL '1 month'
    GROUP BY month
  `).all(monthKey);
  const overrides = await db.prepare(`
    SELECT month_key, opening_override FROM monthly_balances
    WHERE month_key <= ?
  `).all(monthKey);
  const obligations = await db.prepare(`
    SELECT r.month_key, SUM(r.expected_amount) AS expected,
      SUM(COALESCE(p.paid, 0)) AS received
    FROM rent_obligations r
    LEFT JOIN (
      SELECT obligation_id, SUM(amount) AS paid FROM rent_payments GROUP BY obligation_id
    ) p ON p.obligation_id = r.id
    WHERE r.month_key >= '2026-01' AND r.month_key <= ?
    GROUP BY r.month_key
  `).all(monthKey);

  const mapByMonth = (rows) => new Map(rows.map((row) => [row.month, row]));
  const expenseByMonth = mapByMonth(expenses);
  const groceryByMonth = mapByMonth(groceries);
  const incomeByMonth = new Map();
  for (const income of incomes) {
    const sources = incomeByMonth.get(income.month) || [];
    sources.push(income);
    incomeByMonth.set(income.month, sources);
  }
  const monthlyRevenueByMonth = new Map(monthlyRevenues.map((row) => [row.month_key, row]));
  const rentByMonth = mapByMonth(rentPayments);
  const savingByMonth = mapByMonth(savings);
  const obligationByMonth = new Map(obligations.map((row) => [row.month_key, row]));
  const overrideByMonth = new Map(overrides.map((row) => [row.month_key, row.opening_override]));
  const [year, selectedMonth] = monthKey.split('-').map(Number);
  let openingBalance = 0;
  let selected = null;
  for (let y = 2026; y <= year; y += 1) {
    const lastMonth = y === year ? selectedMonth : 12;
    for (let m = 1; m <= lastMonth; m += 1) {
      const key = `${y}-${String(m).padStart(2, '0')}`;
      const monthlyExpenses = expenseByMonth.get(key) || {};
      const monthlyGrocery = groceryByMonth.get(key) || {};
      const monthlyIncome = incomeByMonth.get(key) || [];
      const monthlyRevenue = monthlyRevenueByMonth.get(key);
      const monthlyRent = rentByMonth.get(key) || {};
      const monthlySavings = savingByMonth.get(key) || {};
      const rentObligation = obligationByMonth.get(key) || {};
      const override = overrideByMonth.get(key);
      if (override !== null && override !== undefined) openingBalance = Number(override);
      const otherExpenses = Number(monthlyExpenses.other || 0);
      const protectedSavingsExpenses = Number(monthlyExpenses.protected_savings || 0);
      const boutiqueExpenses = Number(monthlyExpenses.boutique || 0) + Number(monthlyGrocery.total || 0);
      const legacyPape = monthlyIncome
        .filter((income) => normalizeSource(income.source) === 'pape')
        .reduce((sum, income) => sum + Number(income.total), 0);
      const legacyPension = monthlyIncome
        .filter((income) => normalizeSource(income.source) === 'pension pere')
        .reduce((sum, income) => sum + Number(income.total), 0);
      const otherIncome = monthlyIncome
        .filter((income) => !isPrimaryRevenueSource(income.source) && !isRentRevenueSource(income.source))
        .reduce((sum, income) => sum + Number(income.total), 0);
      const papeAmount = Number(monthlyRevenue?.pape_amount ?? legacyPape);
      const pensionAmount = Number(monthlyRevenue?.pension_amount ?? legacyPension);
      const ordinaryIncome = papeAmount + pensionAmount + otherIncome;
      const rentReceived = Number(monthlyRent.total || 0);
      const savingsDeposits = Number(monthlySavings.deposits || 0);
      const savingsWithdrawals = Number(monthlySavings.withdrawals || 0);
      const incomeTotal = ordinaryIncome + rentReceived;
      const totalExpenses = otherExpenses + boutiqueExpenses;
      const { resourcesAvailable, closingBalance } = calculateMonthlyBalance({
        incomeTotal,
        openingBalance,
        totalExpenses,
      });
      selected = key === monthKey ? {
        monthKey: key,
        openingBalance,
        incomeTotal,
        ordinaryIncome,
        papeAmount,
        pensionAmount,
        otherIncome,
        rentReceived,
        rentExpected: Number(rentObligation.expected || 0),
        rentOutstanding: Math.max(0, Number(rentObligation.expected || 0) - Number(rentObligation.received || 0)),
        otherExpenses,
        protectedSavingsExpenses,
        boutiqueExpenses,
        totalExpenses,
        savingsDeposits,
        savingsWithdrawals,
        savingsBalance: await getSavingsBalance(key),
        protectedSavingsBalance: await getSavingsBalance(key, true),
        closingBalance,
        resourcesAvailable,
      } : null;
      openingBalance = closingBalance;
    }
  }
  return selected || {
    monthKey,
    openingBalance: 0,
    incomeTotal: 0,
    ordinaryIncome: 0,
    papeAmount: 0,
    pensionAmount: 0,
    otherIncome: 0,
    rentReceived: 0,
    rentExpected: 0,
    rentOutstanding: 0,
    otherExpenses: 0,
    protectedSavingsExpenses: 0,
    boutiqueExpenses: 0,
    totalExpenses: 0,
    savingsDeposits: 0,
    savingsWithdrawals: 0,
    savingsBalance: await getSavingsBalance(monthKey),
    protectedSavingsBalance: await getSavingsBalance(monthKey, true),
    closingBalance: 0,
    resourcesAvailable: 0,
  };
}

async function getSavingsBalance(monthKey, protectedOnly = false) {
  const row = await db.prepare(`
    WITH latest AS (
      SELECT DISTINCT ON (account_id) account_id, month_key, amount
      FROM savings_balance_snapshots
      WHERE month_key <= ?
      ORDER BY account_id, month_key DESC
    )
    SELECT COALESCE(SUM(account_balance), 0) AS balance
    FROM (
      SELECT COALESCE(latest.amount, 0)
        + COALESCE(SUM(CASE
          WHEN t.transaction_type = 'deposit' THEN t.amount
          ELSE -t.amount
        END), 0) AS account_balance
      FROM savings_accounts a
      LEFT JOIN latest ON latest.account_id = a.id
      LEFT JOIN savings_transactions t ON t.account_id = a.id
        AND to_char(t.transaction_date, 'YYYY-MM') <= ?
        AND (latest.month_key IS NULL OR to_char(t.transaction_date, 'YYYY-MM') > latest.month_key)
      WHERE (? = FALSE OR a.is_protected = TRUE)
      GROUP BY a.id, latest.amount
    ) balances
  `).get(monthKey, monthKey, protectedOnly);
  return Number(row.balance || 0);
}

router.get('/tenants', async (req, res) => {
  const units = await db.prepare('SELECT * FROM units ORDER BY floor, label').all();
  const tenants = await db.prepare(`
    SELECT t.*, u.floor, u.label AS unit_label, u.housing_type
    FROM tenants t JOIN units u ON u.id = t.unit_id
    ORDER BY t.is_active DESC, u.floor, u.label, t.last_name, t.first_name
  `).all();
  res.render('tenants', { units, tenants, floors: FLOORS });
});

router.post('/units', async (req, res) => {
  const { floor, label, housing_type, notes } = req.body;
  if (!FLOORS.includes(floor) || !String(label || '').trim()) {
    return res.status(400).send('Étage et numéro de logement obligatoires.');
  }
  const result = await db.prepare(`
    INSERT INTO units (floor, label, housing_type, notes) VALUES (?, ?, ?, ?)
  `).run(floor, String(label).trim(), String(housing_type || '').trim() || null, String(notes || '').trim() || null);
  await writeAudit(req.user.id, 'unit', result.lastInsertRowid, 'create', null, { floor, label, housing_type });
  res.redirect('/tenants');
});

router.post('/tenants', async (req, res) => {
  const { unit_id, first_name, last_name, people_count, monthly_rent, entry_date, phone } = req.body;
  const rent = amountValue(monthly_rent);
  const people = Number(people_count);
  if (!Number.isInteger(Number(unit_id)) || !String(first_name || '').trim()
    || !String(last_name || '').trim() || rent === null || !Number.isInteger(people) || people < 1) {
    return res.status(400).send('Vérifiez les informations du locataire et le montant du loyer.');
  }
  if (!await db.prepare('SELECT id FROM units WHERE id = ?').get(unit_id)) {
    return res.status(400).send('Le logement sélectionné est introuvable.');
  }
  const result = await db.prepare(`
    INSERT INTO tenants (unit_id, first_name, last_name, people_count, monthly_rent, entry_date, phone)
    VALUES (?, ?, ?, ?, ?, NULLIF(?, '')::date, ?)
  `).run(unit_id, first_name.trim(), last_name.trim(), people, rent, entry_date || '', phone || null);
  await writeAudit(req.user.id, 'tenant', result.lastInsertRowid, 'create', null, req.body);
  res.redirect('/tenants');
});

router.post('/tenants/:id/update', async (req, res) => {
  const id = Number(req.params.id);
  const { unit_id, first_name, last_name, people_count, monthly_rent, entry_date, phone } = req.body;
  const rent = amountValue(monthly_rent);
  const people = Number(people_count);
  if (!Number.isSafeInteger(id) || !FLOORS.includes(await getUnitFloor(unit_id))
    || !String(first_name || '').trim() || !String(last_name || '').trim()
    || rent === null || !Number.isInteger(people) || people < 1) {
    return res.status(400).send('Vérifiez les informations du locataire.');
  }
  const before = await db.prepare('SELECT * FROM tenants WHERE id = ?').get(id);
  if (!before) return res.status(404).send('Locataire introuvable.');
  await db.prepare(`
    UPDATE tenants SET unit_id = ?, first_name = ?, last_name = ?, people_count = ?,
      monthly_rent = ?, entry_date = NULLIF(?, '')::date, phone = ?, updated_at = CURRENT_TIMESTAMP
    WHERE id = ?
  `).run(unit_id, first_name.trim(), last_name.trim(), people, rent, entry_date || '', phone || null, id);
  const after = await db.prepare('SELECT * FROM tenants WHERE id = ?').get(id);
  await writeAudit(req.user.id, 'tenant', id, 'update', before, after);
  res.redirect('/tenants');
});

router.post('/tenants/:id/deactivate', async (req, res) => {
  const id = Number(req.params.id);
  const before = await db.prepare('SELECT * FROM tenants WHERE id = ?').get(id);
  if (!before) return res.status(404).send('Locataire introuvable.');
  await db.prepare('UPDATE tenants SET is_active = FALSE, updated_at = CURRENT_TIMESTAMP WHERE id = ?').run(id);
  await writeAudit(req.user.id, 'tenant', id, 'deactivate', before, { ...before, is_active: false });
  res.redirect('/tenants');
});

async function getUnitFloor(unitId) {
  const unit = await db.prepare('SELECT floor FROM units WHERE id = ?').get(unitId);
  return unit?.floor;
}

router.get('/rents', async (req, res) => {
  const monthKey = validMonth(req.query.month) ? req.query.month : currentMonth();
  const obligations = await db.prepare(`
    SELECT o.id, o.month_key, o.expected_amount, t.first_name, t.last_name,
      u.floor, u.label AS unit_label, COALESCE(SUM(p.amount), 0) AS amount_paid
    FROM rent_obligations o
    JOIN tenants t ON t.id = o.tenant_id
    JOIN units u ON u.id = t.unit_id
    LEFT JOIN rent_payments p ON p.obligation_id = o.id
    WHERE o.month_key = ?
    GROUP BY o.id, t.first_name, t.last_name, u.floor, u.label
    ORDER BY u.floor, u.label, t.last_name, t.first_name
  `).all(monthKey);
  const paymentDetails = await db.prepare(`
    SELECT p.id, p.amount, p.payment_date, p.payment_method, p.comment, o.month_key,
      o.expected_amount, t.first_name, t.last_name, u.floor, u.label AS unit_label
    FROM rent_payments p
    JOIN rent_obligations o ON o.id = p.obligation_id
    JOIN tenants t ON t.id = o.tenant_id
    JOIN units u ON u.id = t.unit_id
    WHERE o.month_key = ?
    ORDER BY p.payment_date DESC, p.id DESC
  `).all(monthKey);
  const totals = obligations.reduce((result, row) => {
    result.expected += Number(row.expected_amount);
    result.received += Number(row.amount_paid);
    return result;
  }, { expected: 0, received: 0 });
  res.render('rents', {
    monthKey,
    obligations,
    paymentDetails,
    totals,
    generated: req.query.generated,
    marked: req.query.marked,
    paymentMethods: ['Espèces', 'Virement', 'Mobile money', 'Chèque', 'Autre'],
  });
});

router.post('/rents/generate', async (req, res) => {
  const monthKey = String(req.body.month_key || '');
  if (!validMonth(monthKey)) return res.status(400).send('Mois invalide.');
  const generated = await db.transaction(async () => {
    const result = await db.prepare(`
      INSERT INTO rent_obligations (tenant_id, month_key, expected_amount)
      SELECT id, ?, monthly_rent
      FROM tenants
      WHERE is_active = TRUE AND (entry_date IS NULL OR entry_date < (? || '-01')::date + INTERVAL '1 month')
      ON CONFLICT (tenant_id, month_key) DO NOTHING
    `).run(monthKey, monthKey);
    await writeAudit(req.user.id, 'rent_obligations', monthKey, 'generate', null, { inserted: result.changes });
    return result.changes;
  });
  res.redirect(`/rents?month=${encodeURIComponent(monthKey)}&generated=${generated}`);
});

async function markRentObligationsPaid(userId, monthKey, paymentDate, obligationIds, action) {
  const placeholders = obligationIds?.map(() => '?').join(', ');
  const idClause = obligationIds ? ` AND id IN (${placeholders})` : '';
  const obligationIdClause = obligationIds ? ` AND o.id IN (${placeholders})` : '';
  return db.transaction(async () => {
    const locked = await db.prepare(`
      SELECT id FROM rent_obligations WHERE month_key = ?${idClause} ORDER BY id FOR UPDATE
    `).all(monthKey, ...(obligationIds || []));
    if (obligationIds && locked.length !== obligationIds.length) return null;
    const obligations = await db.prepare(`
      SELECT o.id, o.month_key, o.expected_amount, COALESCE(SUM(p.amount), 0) AS amount_paid
      FROM rent_obligations o
      LEFT JOIN rent_payments p ON p.obligation_id = o.id
      WHERE o.month_key = ?${obligationIdClause}
      GROUP BY o.id
      ORDER BY o.id
    `).all(monthKey, ...(obligationIds || []));
    let count = 0;
    let totalAmount = 0;
    for (const obligation of obligations) {
      const remaining = Math.round((Number(obligation.expected_amount) - Number(obligation.amount_paid)) * 100) / 100;
      if (remaining <= 0) continue;
      await db.prepare(`
        INSERT INTO rent_payments (
          obligation_id, amount, payment_date, received_month_key, created_by
        ) VALUES (?, ?, ?, ?, ?)
      `).run(
        obligation.id,
        remaining,
        paymentDate,
        paymentDate ? paymentDate.slice(0, 7) : monthKey,
        userId,
      );
      count += 1;
      totalAmount += remaining;
    }
    if (count) {
      await writeAudit(userId, 'rent_payment', monthKey, action, null, {
        count,
        totalAmount,
        paymentDate,
        obligationIds,
      });
    }
    return { count };
  });
}

router.post('/rents/mark-all-paid', async (req, res) => {
  const monthKey = String(req.body.month_key || '');
  const paymentDateInput = String(req.body.payment_date || '').trim();
  const paymentDate = paymentDateInput || null;
  if (!validMonth(monthKey) || (paymentDate && !validDate(paymentDate))) {
    return res.status(400).send('Vérifiez le mois et la date facultative du paiement.');
  }

  const result = await markRentObligationsPaid(
    req.user.id, monthKey, paymentDate, null, 'mark_all_paid',
  );
  res.redirect(`/rents?month=${encodeURIComponent(monthKey)}&marked=${result.count}`);
});

router.post('/rents/mark-selected-paid', async (req, res) => {
  const monthKey = String(req.body.month_key || '');
  const paymentDateInput = String(req.body.payment_date || '').trim();
  const paymentDate = paymentDateInput || null;
  const submittedIds = Array.isArray(req.body.obligation_ids)
    ? req.body.obligation_ids
    : [req.body.obligation_ids];
  if (!validMonth(monthKey) || !submittedIds.length
    || submittedIds.some((value) => !/^\d+$/.test(String(value || '')))
    || (paymentDate && !validDate(paymentDate))) {
    return res.status(400).send('Vérifiez le mois, les loyers sélectionnés et la date facultative du paiement.');
  }
  const obligationIds = [...new Set(submittedIds.map(Number))];
  if (obligationIds.some((id) => !Number.isSafeInteger(id) || id <= 0)) {
    return res.status(400).send('La sélection des loyers est invalide.');
  }

  const result = await markRentObligationsPaid(
    req.user.id, monthKey, paymentDate, obligationIds, 'mark_selected_paid',
  );
  if (!result) return res.status(400).send('Un loyer sélectionné ne correspond pas au mois affiché.');
  res.redirect(`/rents?month=${encodeURIComponent(monthKey)}&marked=${result.count}`);
});

router.post('/rents/:id/payments', async (req, res) => {
  const obligationId = Number(req.params.id);
  const amount = amountValue(req.body.amount);
  const paymentDateInput = String(req.body.payment_date || '').trim();
  const paymentDate = paymentDateInput || null;
  if (!Number.isSafeInteger(obligationId) || amount === null || amount <= 0
    || (paymentDate && !validDate(paymentDate))) {
    return res.status(400).send('Vérifiez le montant et la date facultative du paiement.');
  }
  const obligation = await db.transaction(async () => {
    const row = await db.prepare('SELECT * FROM rent_obligations WHERE id = ? FOR UPDATE').get(obligationId);
    if (!row) return null;
    const paid = await db.prepare('SELECT COALESCE(SUM(amount), 0) AS total FROM rent_payments WHERE obligation_id = ?')
      .get(obligationId);
    if (Number(paid.total) + amount > Number(row.expected_amount)) {
      throw new Error('Le paiement dépasse le solde de loyer restant.');
    }
    const result = await db.prepare(`
      INSERT INTO rent_payments (
        obligation_id, amount, payment_date, received_month_key, payment_method, comment, created_by
      ) VALUES (?, ?, ?, ?, ?, ?, ?)
    `).run(
      obligationId,
      amount,
      paymentDate,
      paymentDate ? paymentDate.slice(0, 7) : row.month_key,
      req.body.payment_method || null,
      req.body.comment || null,
      req.user.id,
    );
    await writeAudit(req.user.id, 'rent_payment', result.lastInsertRowid, 'create', null, req.body);
    return row;
  }).catch((error) => {
    if (error.message.includes('dépasse le solde')) return { error: error.message };
    throw error;
  });
  if (!obligation) return res.status(404).send('Obligation de loyer introuvable.');
  if (obligation.error) return res.status(400).send(obligation.error);
  res.redirect(`/rents?month=${obligation.month_key}`);
});

router.post('/rents/payments/:id/delete', async (req, res) => {
  const paymentId = Number(req.params.id);
  if (!Number.isSafeInteger(paymentId) || paymentId <= 0) {
    return res.status(400).send('Identifiant de paiement invalide.');
  }

  const deletedPayment = await db.transaction(async () => {
    const before = await db.prepare(`
      SELECT p.*, o.month_key AS obligation_month
      FROM rent_payments p
      JOIN rent_obligations o ON o.id = p.obligation_id
      WHERE p.id = ?
      FOR UPDATE OF p, o
    `).get(paymentId);
    if (!before) return null;
    await db.prepare('DELETE FROM rent_payments WHERE id = ?').run(paymentId);
    await writeAudit(req.user.id, 'rent_payment', paymentId, 'delete', before, null);
    return before;
  });
  if (!deletedPayment) return res.status(404).send('Paiement de loyer introuvable.');

  const receivedMonth = deletedPayment.received_month_key
    || deletedPayment.payment_date?.slice(0, 7)
    || deletedPayment.obligation_month;
  const requestedReturn = String(req.body.return_to || '');
  const returnTo = requestedReturn === '/history'
    || /^\/income\?month=\d{4}-(0[1-9]|1[0-2])$/.test(requestedReturn)
    ? requestedReturn
    : `/income?month=${encodeURIComponent(receivedMonth)}`;
  res.redirect(returnTo);
});

router.get('/income', async (req, res) => {
  const monthKey = validMonth(req.query.month) ? req.query.month : currentMonth();
  const [categories, incomeRows, monthlyRevenue, finance, rentPayments] = await Promise.all([
    db.prepare('SELECT * FROM income_categories ORDER BY name').all(),
    db.prepare(`
      SELECT i.*, c.name AS category_name FROM incomes i
      JOIN income_categories c ON c.id = i.category_id
      WHERE i.month_key = ? ORDER BY i.income_date DESC, i.id DESC
    `).all(monthKey),
    db.prepare('SELECT * FROM monthly_revenues WHERE month_key = ?').get(monthKey),
    getMonthlyFinance(monthKey),
    db.prepare(`
      SELECT p.id, p.payment_date, p.amount, p.payment_method, t.first_name, t.last_name,
        u.floor, u.label AS unit_label
      FROM rent_payments p
      JOIN rent_obligations o ON o.id = p.obligation_id
      JOIN tenants t ON t.id = o.tenant_id
      JOIN units u ON u.id = t.unit_id
      WHERE COALESCE(NULLIF(p.received_month_key, ''), to_char(p.payment_date, 'YYYY-MM'), o.month_key) = ?
      ORDER BY p.payment_date DESC NULLS LAST, p.id DESC
    `).all(monthKey),
  ]);
  const baseRows = incomeRows.filter((income) => isPrimaryRevenueSource(income.source));
  const legacyAmount = (source) => baseRows
    .filter((income) => normalizeSource(income.source) === source)
    .reduce((sum, income) => sum + Number(income.amount), 0);
  const isPapeLegacy = baseRows.some((income) => normalizeSource(income.source) === 'pape');
  const isPensionLegacy = baseRows.some((income) => normalizeSource(income.source) === 'pension pere');
  const revenues = {
    papeAmount: Number(monthlyRevenue?.pape_amount ?? (isPapeLegacy ? legacyAmount('pape') : 0)),
    papeReceivedDate: monthlyRevenue?.pape_received_date || baseRows
      .find((income) => normalizeSource(income.source) === 'pape')?.income_date || '',
    pensionAmount: Number(monthlyRevenue?.pension_amount
      ?? (isPensionLegacy ? legacyAmount('pension pere') : 180000)),
    pensionReceivedDate: monthlyRevenue?.pension_received_date || baseRows
      .find((income) => normalizeSource(income.source) === 'pension pere')?.income_date || '',
    comment: monthlyRevenue?.comment || '',
    isSaved: Boolean(monthlyRevenue),
  };
  const incomes = incomeRows.filter((income) => (
    !isPrimaryRevenueSource(income.source) && !isRentRevenueSource(income.source)
  ));
  const otherCategories = categories.filter((category) => (
    !isPrimaryRevenueSource(category.name) && !isRentRevenueSource(category.name)
  ));
  res.render('income', {
    monthKey,
    categories: otherCategories,
    incomes,
    revenues,
    finance,
    rentPayments,
    notice: req.query.saved === '1' ? 'Le revenu a été enregistré.' : null,
  });
});

router.post('/income/monthly', async (req, res) => {
  const monthKey = String(req.body.month_key || '');
  const papeAmount = amountValue(req.body.pape_amount);
  const pensionAmount = amountValue(req.body.pension_amount);
  const papeDate = String(req.body.pape_received_date || '');
  const pensionDate = String(req.body.pension_received_date || '');
  const comment = String(req.body.comment || '').trim() || null;
  const validReceivedDate = (date) => !date || (
    validDate(date) && date.slice(0, 7) === monthKey
  );
  if (!validMonth(monthKey) || papeAmount === null || pensionAmount === null
    || !validReceivedDate(papeDate) || !validReceivedDate(pensionDate)) {
    return res.status(400).send('Vérifiez les montants et les dates de réception de Pape et de la pension.');
  }

  await db.transaction(async () => {
    const before = await db.prepare('SELECT * FROM monthly_revenues WHERE month_key = ? FOR UPDATE').get(monthKey);
    await db.prepare(`
      INSERT INTO monthly_revenues (
        month_key, pape_amount, pape_received_date, pension_amount, pension_received_date,
        comment, created_by, updated_by
      ) VALUES (?, ?, NULLIF(?, '')::date, ?, NULLIF(?, '')::date, ?, ?, ?)
      ON CONFLICT (month_key) DO UPDATE SET
        pape_amount = EXCLUDED.pape_amount,
        pape_received_date = EXCLUDED.pape_received_date,
        pension_amount = EXCLUDED.pension_amount,
        pension_received_date = EXCLUDED.pension_received_date,
        comment = EXCLUDED.comment,
        updated_by = EXCLUDED.updated_by,
        updated_at = CURRENT_TIMESTAMP
    `).run(
      monthKey,
      papeAmount,
      papeDate,
      pensionAmount,
      pensionDate,
      comment,
      before?.created_by || req.user.id,
      req.user.id,
    );
    const after = await db.prepare('SELECT * FROM monthly_revenues WHERE month_key = ?').get(monthKey);
    await writeAudit(req.user.id, 'monthly_revenue', monthKey, before ? 'update' : 'create', before, after);
  });
  res.redirect(`/income?month=${monthKey}&saved=1`);
});

router.post('/income', async (req, res) => {
  const { income_date, title, category_id, amount, source, payment_method, month_key, comment } = req.body;
  const value = amountValue(amount);
  if (!validDate(income_date) || !String(title || '').trim()
    || !Number.isSafeInteger(Number(category_id)) || value === null || value <= 0
    || !String(source || '').trim() || !validMonth(month_key)) {
    return res.status(400).send('Vérifiez les champs obligatoires de la recette.');
  }
  const category = await db.prepare('SELECT name FROM income_categories WHERE id = ?').get(category_id);
  if (!category) return res.status(400).send('Catégorie de recette introuvable.');
  if ([title, source, category.name].some(isPrimaryRevenueSource)) {
    return res.status(400).send('Pape et la pension du père doivent être saisis dans les revenus mensuels.');
  }
  if (/loyer|locataire|location/i.test(`${title} ${source} ${category.name}`)) {
    return res.status(400).send('Les loyers doivent être saisis dans le module « Loyers » afin d’éviter leur double comptabilisation.');
  }
  const result = await db.prepare(`
    INSERT INTO incomes (income_date, title, category_id, amount, source, payment_method, month_key, comment, created_by)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(income_date, title.trim(), category_id, value, source.trim(), payment_method || null, month_key, comment || null, req.user.id);
  await writeAudit(req.user.id, 'income', result.lastInsertRowid, 'create', null, req.body);
  res.redirect(`/income?month=${month_key}&saved=1`);
});

router.post('/income/:id/update', async (req, res) => {
  const id = Number(req.params.id);
  const { income_date, title, category_id, amount, source, payment_method, month_key, comment } = req.body;
  const value = amountValue(amount);
  if (!Number.isSafeInteger(id) || !validDate(income_date) || !String(title || '').trim()
    || !Number.isSafeInteger(Number(category_id)) || value === null || value <= 0
    || !String(source || '').trim() || !validMonth(month_key)) {
    return res.status(400).send('Vérifiez les champs obligatoires de la recette.');
  }
  if ([title, source].some(isPrimaryRevenueSource)
    || /loyer|locataire|location/i.test(`${title} ${source}`)) {
    return res.status(400).send('Cette source doit être gérée dans son module dédié.');
  }
  const category = await db.prepare('SELECT name FROM income_categories WHERE id = ?').get(category_id);
  if (!category || isPrimaryRevenueSource(category.name) || /loyer|locataire|location/i.test(category.name)) {
    return res.status(400).send('Catégorie de recette introuvable ou réservée aux loyers.');
  }
  const before = await db.prepare('SELECT * FROM incomes WHERE id = ?').get(id);
  if (!before) return res.status(404).send('Recette introuvable.');
  await db.transaction(async () => {
    await db.prepare(`
      UPDATE incomes SET income_date = ?, title = ?, category_id = ?, amount = ?, source = ?,
        payment_method = ?, month_key = ?, comment = ?, updated_at = CURRENT_TIMESTAMP
      WHERE id = ?
    `).run(
      income_date,
      title.trim(),
      category_id,
      value,
      source.trim(),
      payment_method || null,
      month_key,
      comment || null,
      id,
    );
    const after = await db.prepare('SELECT * FROM incomes WHERE id = ?').get(id);
    await writeAudit(req.user.id, 'income', id, 'update', before, after);
  });
  res.redirect(`/income?month=${month_key}&saved=1`);
});

router.post('/income/:id/delete', async (req, res) => {
  const id = Number(req.params.id);
  const before = await db.prepare('SELECT * FROM incomes WHERE id = ?').get(id);
  if (!before) return res.status(404).send('Recette introuvable.');
  if (isPrimaryRevenueSource(before.source)) {
    return res.status(400).send('Pape et la pension du père doivent être modifiés dans les revenus mensuels.');
  }
  await db.transaction(async () => {
    await db.prepare('DELETE FROM incomes WHERE id = ?').run(id);
    await writeAudit(req.user.id, 'income', id, 'delete', before, null);
  });
  res.redirect(`/income?month=${before.month_key}&saved=1`);
});

router.post('/income-categories', async (req, res) => {
  const name = String(req.body.name || '').trim();
  if (!name) return res.status(400).send('Le nom de la catégorie est obligatoire.');
  await db.prepare('INSERT INTO income_categories (name) VALUES (?) ON CONFLICT (name) DO NOTHING').run(name);
  res.redirect('/income');
});

router.get('/savings', async (req, res) => {
  const [accounts, transactions, balanceSnapshots, categories] = await Promise.all([
    db.prepare(`
      SELECT a.id, a.name, a.is_protected,
        COALESCE(snapshot.amount, 0) + COALESCE(SUM(CASE
          WHEN t.transaction_type = 'deposit' THEN t.amount
          ELSE -t.amount
        END), 0) AS balance,
        snapshot.month_key AS snapshot_month,
        first_snapshot.amount AS initial_balance
      FROM savings_accounts a
      LEFT JOIN LATERAL (
        SELECT month_key, amount FROM savings_balance_snapshots
        WHERE account_id = a.id ORDER BY month_key DESC LIMIT 1
      ) snapshot ON TRUE
      LEFT JOIN LATERAL (
        SELECT amount FROM savings_balance_snapshots
        WHERE account_id = a.id ORDER BY month_key, id LIMIT 1
      ) first_snapshot ON TRUE
      LEFT JOIN savings_transactions t ON t.account_id = a.id
        AND (snapshot.month_key IS NULL OR to_char(t.transaction_date, 'YYYY-MM') > snapshot.month_key)
      GROUP BY a.id, snapshot.amount, snapshot.month_key, first_snapshot.amount
      ORDER BY a.name
    `).all(),
    db.prepare(`
      SELECT t.*, a.name AS account_name, a.is_protected, e.title AS related_expense_title
      FROM savings_transactions t
      JOIN savings_accounts a ON a.id = t.account_id
      LEFT JOIN expenses e ON e.id = t.related_expense_id
      ORDER BY transaction_date DESC, t.id DESC LIMIT 100
    `).all(),
    db.prepare(`
      SELECT s.*, a.name AS account_name
      FROM savings_balance_snapshots s JOIN savings_accounts a ON a.id = s.account_id
      ORDER BY s.month_key DESC, s.id DESC LIMIT 100
    `).all(),
    db.prepare("SELECT * FROM categories WHERE type <> 'boutique' ORDER BY name").all(),
  ]);
  const requestedTransactionId = /^savings-details-(\d+)$/.exec(String(req.query.open || ''))?.[1];
  if (requestedTransactionId && !transactions.some((transaction) => Number(transaction.id) === Number(requestedTransactionId))) {
    const requestedTransaction = await db.prepare(`
      SELECT t.*, a.name AS account_name, a.is_protected, e.title AS related_expense_title
      FROM savings_transactions t
      JOIN savings_accounts a ON a.id = t.account_id
      LEFT JOIN expenses e ON e.id = t.related_expense_id
      WHERE t.id = ?
    `).get(Number(requestedTransactionId));
    if (requestedTransaction) transactions.push(requestedTransaction);
  }
  res.render('savings', { accounts, transactions, balanceSnapshots, categories });
});

router.post('/savings/protected-expenses', async (req, res) => {
  const title = String(req.body.title || '').trim();
  const categoryId = Number(req.body.category_id);
  const amount = amountValue(req.body.amount);
  const expenseDate = String(req.body.expense_date || '');
  const monthKey = validMonth(req.body.month_key) ? String(req.body.month_key) : expenseDate.slice(0, 7);
  const description = String(req.body.description || '').trim();
  if (!title || title.length > 250 || !Number.isSafeInteger(categoryId) || !validDate(expenseDate)
    || !validMonth(monthKey) || amount === null || amount <= 0
    || req.body.protected_confirmation !== '1') {
    return res.status(400).send('Vérifiez la désignation, la catégorie, la date, le montant et la confirmation de cette dépense protégée.');
  }
  const category = await db.prepare("SELECT id FROM categories WHERE id = ? AND type <> 'boutique'")
    .get(categoryId);
  if (!category) return res.status(400).send('Catégorie de dépense invalide.');

  try {
    await db.transaction(async () => {
      const account = await db.prepare(`
        SELECT id FROM savings_accounts WHERE is_protected = TRUE ORDER BY id LIMIT 1 FOR UPDATE
      `).get();
      if (!account) throw Object.assign(new Error('Le compte d’épargne protégée est introuvable.'), { status: 404 });

      const snapshot = await db.prepare(`
        SELECT month_key, amount FROM savings_balance_snapshots
        WHERE account_id = ? ORDER BY month_key DESC LIMIT 1
      `).get(account.id);
      if (snapshot && (monthKey <= snapshot.month_key || expenseDate.slice(0, 7) <= snapshot.month_key)) {
        throw Object.assign(new Error(
          `La période et la date de la dépense doivent être postérieures au dernier solde observé (${snapshot.month_key}).`,
        ), { status: 400 });
      }
      const balance = await db.prepare(`
        SELECT COALESCE(?, 0) + COALESCE(SUM(CASE
          WHEN t.transaction_type = 'deposit' THEN t.amount ELSE -t.amount
        END), 0) AS balance
        FROM savings_transactions t
        WHERE t.account_id = ?
          AND (?::varchar IS NULL OR to_char(t.transaction_date, 'YYYY-MM') > ?)
      `).get(snapshot?.amount ?? null, account.id, snapshot?.month_key ?? null, snapshot?.month_key ?? null);
      if (Number(balance.balance) < amount) {
        throw Object.assign(new Error('La dépense dépasse le solde de l’épargne protégée.'), { status: 400 });
      }

      const expense = await db.prepare(`
        INSERT INTO expenses (
          title, category_id, amount, expense_date, description, reference, beneficiary,
          paid_by, payment_method, month_key, funding_source, created_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'protected_savings', ?)
      `).run(
        title,
        categoryId,
        amount,
        expenseDate,
        description,
        String(req.body.reference || '').trim(),
        String(req.body.beneficiary || '').trim(),
        String(req.body.paid_by || '').trim(),
        String(req.body.payment_method || '').trim(),
        monthKey,
        req.user.id,
      );
      const transaction = await db.prepare(`
        INSERT INTO savings_transactions (
          account_id, transaction_date, transaction_type, amount, comment, related_expense_id, created_by
        ) VALUES (?, ?, 'withdrawal', ?, ?, ?, ?)
      `).run(
        account.id,
        expenseDate,
        amount,
        `Dépense financée par l’épargne protégée : ${title}`,
        expense.lastInsertRowid,
        req.user.id,
      );
      const expenseRow = await db.prepare('SELECT * FROM expenses WHERE id = ?').get(expense.lastInsertRowid);
      await writeAudit(req.user.id, 'expense', expense.lastInsertRowid, 'create', null, expenseRow);
      await writeAudit(req.user.id, 'savings_transaction', transaction.lastInsertRowid, 'create', null, {
        account_id: account.id,
        transaction_type: 'withdrawal',
        amount,
        related_expense_id: expense.lastInsertRowid,
      });
    });
  } catch (error) {
    if (error.status) return res.status(error.status).send(error.message);
    throw error;
  }
  res.redirect('/savings');
});

router.post('/savings/accounts', async (req, res) => {
  const name = String(req.body.name || '').trim();
  if (!name) return res.status(400).send('Le nom du compte d’épargne est obligatoire.');
  if (normalizeSource(name) === 'epargne pension pere') {
    return res.status(409).send('Le compte protégé « Épargne pension père » existe déjà.');
  }
  await db.prepare('INSERT INTO savings_accounts (name) VALUES (?)').run(name);
  res.redirect('/savings');
});

router.post('/savings/transactions', async (req, res) => {
  const { account_id, transaction_date, transaction_type, amount, comment } = req.body;
  const value = amountValue(amount);
  if (!Number.isSafeInteger(Number(account_id)) || !validDate(transaction_date)
    || !['deposit', 'withdrawal'].includes(transaction_type) || value === null || value <= 0) {
    return res.status(400).send('Vérifiez le mouvement d’épargne.');
  }
  try {
    await db.transaction(async () => {
      const account = await db.prepare('SELECT id, is_protected FROM savings_accounts WHERE id = ? FOR UPDATE')
        .get(account_id);
      if (!account) throw Object.assign(new Error('Compte d’épargne introuvable.'), { status: 404 });
      if (account.is_protected && req.body.protected_confirmation !== '1') {
        throw Object.assign(new Error('Confirmez explicitement cette opération sur l’épargne pension père protégée.'), { status: 400 });
      }
      if (account.is_protected && transaction_type === 'withdrawal' && !String(comment || '').trim()) {
        throw Object.assign(new Error('Indiquez le motif du retrait de l’épargne pension père.'), { status: 400 });
      }
      if (transaction_type === 'withdrawal') {
        const balance = await db.prepare(`
          SELECT
            COALESCE((
              SELECT amount FROM savings_balance_snapshots
              WHERE account_id = ? ORDER BY month_key DESC LIMIT 1
            ), 0) + COALESCE((
              SELECT SUM(CASE WHEN transaction_type = 'deposit' THEN amount ELSE -amount END)
              FROM savings_transactions t
              WHERE t.account_id = ?
                AND to_char(t.transaction_date, 'YYYY-MM') > COALESCE((
                  SELECT month_key FROM savings_balance_snapshots
                  WHERE account_id = ? ORDER BY month_key DESC LIMIT 1
                ), '0000-00')
            ), 0) AS balance
        `).get(account.id, account.id, account.id);
        if (Number(balance.balance) < value) {
          throw Object.assign(new Error('Le retrait dépasse le solde de cette épargne.'), { status: 400 });
        }
      }
      const result = await db.prepare(`
        INSERT INTO savings_transactions (account_id, transaction_date, transaction_type, amount, comment, created_by)
        VALUES (?, ?, ?, ?, ?, ?)
      `).run(account_id, transaction_date, transaction_type, value, comment || null, req.user.id);
      await writeAudit(req.user.id, 'savings_transaction', result.lastInsertRowid, 'create', null, {
        ...req.body,
        protected_account: account.is_protected,
      });
    });
  } catch (error) {
    if (error.status) return res.status(error.status).send(error.message);
    throw error;
  }
  res.redirect('/savings');
});

router.get('/finance/monthly-data', async (req, res) => {
  const monthKey = validMonth(req.query.month) ? req.query.month : currentMonth();
  res.json(await getMonthlyFinance(monthKey));
});

router.post('/balances/adjustment', async (req, res) => {
  const monthKey = String(req.body.month_key || '');
  const opening = Number(req.body.opening_balance);
  const reason = String(req.body.reason || '').trim();
  if (!validMonth(monthKey) || !Number.isFinite(opening) || !reason) {
    return res.status(400).send('Mois, solde initial et motif de correction obligatoires.');
  }
  await db.transaction(async () => {
    const previous = await db.prepare('SELECT opening_override FROM monthly_balances WHERE month_key = ?')
      .get(monthKey);
    await db.prepare(`
      INSERT INTO balance_adjustments (month_key, previous_opening, new_opening, reason, changed_by)
      VALUES (?, ?, ?, ?, ?)
    `).run(monthKey, previous?.opening_override ?? null, opening, reason, req.user.id);
    await db.prepare(`
      INSERT INTO monthly_balances (month_key, opening_override, correction_reason, updated_by)
      VALUES (?, ?, ?, ?)
      ON CONFLICT (month_key) DO UPDATE SET
        opening_override = EXCLUDED.opening_override,
        correction_reason = EXCLUDED.correction_reason,
        updated_by = EXCLUDED.updated_by,
        updated_at = CURRENT_TIMESTAMP
    `).run(monthKey, opening, reason, req.user.id);
    await writeAudit(req.user.id, 'monthly_balance', monthKey, 'adjust', previous, { opening, reason });
  });
  res.redirect(`/monthly?year=${monthKey.slice(0, 4)}&month=${Number(monthKey.slice(5, 7))}`);
});

router.post('/categories/:id/update', async (req, res) => {
  const id = Number(req.params.id);
  const name = String(req.body.name || '').trim();
  const color = /^#[0-9a-f]{6}$/i.test(String(req.body.color || '')) ? req.body.color : '#929d99';
  if (!Number.isSafeInteger(id) || !name) return res.status(400).send('Nom de catégorie invalide.');
  const before = await db.prepare('SELECT * FROM categories WHERE id = ?').get(id);
  if (!before) return res.status(404).send('Catégorie introuvable.');
  if (before.type === 'boutique') return res.status(403).send('La catégorie du compte boutique ne peut pas être renommée.');
  await db.prepare('UPDATE categories SET name = ?, color = ? WHERE id = ?').run(name, color, id);
  const after = await db.prepare('SELECT * FROM categories WHERE id = ?').get(id);
  await writeAudit(req.user.id, 'category', id, 'update', before, after);
  res.redirect('/categories');
});

router.post('/categories/:id/delete', async (req, res, next) => {
  const id = Number(req.params.id);
  try {
    await db.transaction(async () => {
      const category = await db.prepare('SELECT * FROM categories WHERE id = ? FOR UPDATE').get(id);
      if (!category) throw Object.assign(new Error('Catégorie introuvable.'), { status: 404 });
      if (category.type === 'boutique' || category.name === 'Autres dépenses') {
        throw Object.assign(new Error('Cette catégorie est réservée et ne peut pas être supprimée.'), { status: 403 });
      }
      const fallback = await db.prepare("SELECT id FROM categories WHERE name = 'Autres dépenses'").get();
      if (!fallback || Number(fallback.id) === id) throw new Error('Catégorie de remplacement indisponible.');
      await db.prepare('UPDATE expenses SET category_id = ?, updated_at = CURRENT_TIMESTAMP WHERE category_id = ?')
        .run(fallback.id, id);
      await db.prepare('DELETE FROM categories WHERE id = ?').run(id);
      await writeAudit(req.user.id, 'category', id, 'delete', category, null);
    });
    res.redirect('/categories');
  } catch (error) {
    next(error);
  }
});

module.exports = { router, getMonthlyFinance, getSavingsBalance, writeAudit };
