# Gestion des dépenses familiales

Application web Node.js pour gérer les dépenses quotidiennes, le compte boutique, les bilans mensuels et les exports Excel annuels.

## Pré-requis

- Node.js 18+
- npm
- PostgreSQL 14+

## Installation

```bash
npm install
cp .env.example .env
npm run create-admin
npm start
```

Le site est ensuite disponible sur http://localhost:3000.
Avant `npm run create-admin`, renseignez `ADMIN_EMAIL` et un mot de passe d’au
moins 12 caractères dans `.env`. Le compte administrateur doit être créé avant
la première connexion.

La connexion utilise `DATABASE_URL` ou les variables `PGHOST`, `PGPORT`,
`PGDATABASE`, `PGUSER` et `PGPASSWORD`. Pour un serveur local avec
authentification par socket, `PGHOST=/var/run/postgresql` et un rôle PostgreSQL
correspondant à l’utilisateur du système suffisent.

## Fonctionnalités

- Dashboard avec synthèse mensuelle
- Gestion des dépenses ordinaires
- Gestion des logements, locataires, loyers et paiements partiels
- Suivi des recettes, de l’épargne et des soldes mensuels
- Gestion autonome du compte boutique
- Catégories personnalisables
- Bilans mensuels
- Export Excel annuel
- Sauvegardes PostgreSQL locales au format custom (`pg_dump`)
- Import initial de fichiers Excel avec prévisualisation
- Détection de doublons potentiels dans les anciens imports
- Sauvegarde optionnelle vers Cloudflare R2
- Restauration de sauvegardes locales depuis l’interface

## Sauvegarde

La base PostgreSQL est sauvegardée dans `data/backups/` via la route `/backup`.
Les sauvegardes sont vérifiées avec `pg_restore --list` avant d’être proposées.
Si les variables R2 sont configurées, elles sont également envoyées vers Cloudflare R2.

## Migration initiale depuis SQLite

Après avoir configuré et créé la base PostgreSQL cible, sauvegardez le fichier
SQLite existant puis lancez une seule fois :

```bash
npm run migrate:sqlite
```

Le script conserve les identifiants, importe toutes les tables métier et vérifie
les nombres de lignes ainsi que les totaux des dépenses et des marchandises.
Le fichier SQLite n’est pas supprimé.

## Structure

- `src/db.js` : connexion PostgreSQL et schéma initial
- `src/migrate-sqlite-to-postgres.js` : transfert ponctuel depuis SQLite
- `src/app.js` : routes Express et logique métier
- `views/` : templates EJS
- `public/` : feuilles de style
- `data/` : fichiers d’import et sauvegardes PostgreSQL
# app-depenses
