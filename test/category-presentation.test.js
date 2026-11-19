const test = require('node:test');
const assert = require('node:assert/strict');
const {
  PRIORITY_CATEGORIES,
  resolvePriorityCategory,
  sortCategoriesForPresentation,
  buildCategoryPresentation,
  filterCategoryPresentation,
} = require('../src/category-presentation');

test('priority categories keep the mandated order and recognize existing aliases', () => {
  assert.deepEqual(PRIORITY_CATEGORIES, [
    'Alimentaire',
    'Eau',
    'Électricité',
    'Wifi',
    'Compte boutique',
    'Scolarité et frais éducatifs',
    'Hygiène et beauté',
    'Aumône et dons',
    'Entretien de la maison',
    'Santé',
    'Transport',
  ]);
  assert.equal(resolvePriorityCategory('Dépenses alimentaires'), 'Alimentaire');
  assert.equal(resolvePriorityCategory('Pain et achats alimentaires complémentaires'), 'Alimentaire');
  assert.equal(resolvePriorityCategory('Facture d’eau'), 'Eau');
  assert.equal(resolvePriorityCategory('Facture d’électricité'), 'Électricité');
  assert.equal(resolvePriorityCategory('Facture Wi-Fi'), 'Wifi');
  assert.equal(resolvePriorityCategory('Aumônes et dons'), 'Aumône et dons');
  assert.equal(resolvePriorityCategory('Réparations et entretien de la maison'), 'Entretien de la maison');
  assert.equal(resolvePriorityCategory('Santé, consultations et médicaments'), 'Santé');
});

test('presentation always includes missing and empty priority categories without DB creation', () => {
  const presentation = buildCategoryPresentation({
    monthKey: '2026-09',
    categories: [
      { id: 1, name: 'Dépenses alimentaires' },
      { id: 2, name: 'Transport' },
      { id: 99, name: 'Loisirs' },
    ],
    expenses: [
      { category_id: 1, category_name: 'Dépenses alimentaires', title: 'Courses', amount: 12500 },
      { category_id: 99, category_name: 'Loisirs', title: 'Sortie', amount: 3000 },
    ],
  });

  assert.equal(presentation.priority.length, 11);
  assert.deepEqual(presentation.priority.map((group) => group.heading), PRIORITY_CATEGORIES);
  assert.equal(presentation.priority[0].total, 12500);
  assert.equal(presentation.priority[0].count, 1);
  assert.equal(presentation.priority[1].total, 0);
  assert.equal(presentation.priority[1].count, 0);
  assert.equal(presentation.priority[10].heading, 'Transport');
  assert.equal(presentation.priority[10].total, 0);
  assert.deepEqual(presentation.extras.map((group) => group.heading), ['Loisirs']);
});

test('groceries are included once in Compte boutique and extra categories stay alphabetical', () => {
  const presentation = buildCategoryPresentation({
    monthKey: '2026-10',
    categories: [
      { id: 3, name: 'Animaux' },
      { id: 4, name: 'Voyages' },
    ],
    expenses: [
      { category_id: 3, category_name: 'Animaux', title: 'Croquettes', amount: 4500 },
    ],
    groceryItems: [
      { id: 8, item_name: 'Riz', total_amount: 12000 },
      { id: 9, item_name: 'Huile', total_amount: 7000 },
    ],
  });

  assert.equal(presentation.priority.find((group) => group.heading === 'Compte boutique').total, 19000);
  assert.equal(presentation.priority.find((group) => group.heading === 'Compte boutique').count, 2);
  assert.deepEqual(presentation.extras.map((group) => group.heading), ['Animaux', 'Voyages']);
});

test('categories created later are sorted into priority or extra positions', () => {
  const sorted = sortCategoriesForPresentation([
    { id: 1, name: 'Zèbres' },
    { id: 2, name: 'Transport' },
    { id: 3, name: 'Alimentation des animaux' },
    { id: 4, name: 'Santé, consultations et médicaments' },
  ]);
  assert.deepEqual(sorted.map((category) => category.name), [
    'Santé, consultations et médicaments',
    'Transport',
    'Alimentation des animaux',
    'Zèbres',
  ]);
});

test('expense category display matches the configured Render category list and omits missing defaults', () => {
  const categories = [
    { id: 1, name: 'Dépenses alimentaires' },
    { id: 2, name: 'Eau' },
    { id: 3, name: 'Compte boutique', type: 'boutique' },
    { id: 4, name: 'Autres dépenses' },
    { id: 5, name: 'Alimentation des animaux' },
    { id: 6, name: 'Loisirs' },
  ];
  const presentation = filterCategoryPresentation(buildCategoryPresentation({
    monthKey: '2026-11',
    categories,
    expenses: [
      { category_id: 1, category_name: 'Dépenses alimentaires', title: 'Courses', amount: 1000 },
      { category_id: 2, category_name: 'Eau', title: 'Facture', amount: 500 },
      { category_id: 6, category_name: 'Loisirs', title: 'Sortie', amount: 200 },
    ],
    groceryItems: [{ id: 7, item_name: 'Riz', total_amount: 1200 }],
  }), categories);

  assert.deepEqual(presentation.priority.map((group) => group.heading), ['Alimentaire', 'Compte boutique']);
  assert.equal(presentation.priority[0].total, 1000);
  assert.equal(presentation.priority[1].total, 1200);
  assert.deepEqual(presentation.extras.map((group) => group.heading), [
    'Alimentation des animaux',
    'Autres dépenses',
  ]);
});
