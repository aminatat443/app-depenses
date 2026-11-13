# Gestion des dépenses familiales

Application web Node.js pour gérer les dépenses quotidiennes, le compte boutique, les bilans mensuels et les exports Excel annuels.

## Pré-requis

- Node.js 18+
- npm

## Installation

```bash
npm install
cp .env.example .env
npm start
```

Le site est ensuite disponible sur http://localhost:3000.

## Fonctionnalités

- Dashboard avec synthèse mensuelle
- Gestion des dépenses ordinaires
- Gestion autonome du compte boutique
- Catégories personnalisables
- Bilans mensuels
- Export Excel annuel
- Sauvegarde locale de la base SQLite
- Import initial de fichiers Excel avec prévisualisation
- Détection de doublons potentiels dans les anciens imports
- Sauvegarde optionnelle vers Cloudflare R2
- Restauration de sauvegardes locales depuis l’interface

## Sauvegarde

La base SQLite est sauvegardée dans `data/backups/` via la route `/backup`. Les sauvegardes sont conservées localement pour limiter les risques de perte de données si la base Render expire. Si les variables R2 sont configurées, les sauvegardes sont également poussées vers Cloudflare R2.

## Structure

- `src/db.js` : base SQLite et schéma initial
- `src/app.js` : routes Express et logique métier
- `views/` : templates EJS
- `public/` : feuilles de style
- `data/` : base SQLite et sauvegardes
# app-depenses
