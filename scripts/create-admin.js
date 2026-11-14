require('dotenv').config();
const { db, initializeDatabase } = require('../src/db');
const { hashPassword } = require('../src/auth');

async function createAdmin() {
  const email = String(process.env.ADMIN_EMAIL || '').trim().toLowerCase();
  const password = String(process.env.ADMIN_PASSWORD || '');
  if (!email || !email.includes('@') || password.length < 12) {
    throw new Error('Définissez ADMIN_EMAIL et un ADMIN_PASSWORD d’au moins 12 caractères.');
  }

  await initializeDatabase();
  const existing = await db.prepare('SELECT email FROM users LIMIT 1').get();
  if (existing) throw new Error('Un compte administrateur unique existe déjà.');

  await db.prepare('INSERT INTO users (email, password_hash) VALUES (?, ?)')
    .run(email, await hashPassword(password));
  console.log(`Compte administrateur créé : ${email}`);
}

createAdmin()
  .catch((error) => {
    console.error(`Création du compte impossible : ${error.message}`);
    process.exitCode = 1;
  })
  .finally(() => db.close());
