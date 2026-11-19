const test = require('node:test');
const assert = require('node:assert/strict');
const {
  CATEGORY_ORDER,
  OTHER_CATEGORY_NAMES,
  PRIORITY_CATEGORIES,
  resolvePriorityCategory,
  sortCategoriesForPresentation,
  buildCategoryPresentation,
  filterCategoryPresentation,
} = require('../src/category-presentation');

test('category order matches the Render list and legacy aliases remain recognized', () => {
  assert.deepEqual(CATEGORY_ORDER, [
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
  ]);
  assert.equal(OTHER_CATEGORY_NAMES.size, 3);
  assert.equal(PRIORITY_CATEGORIES.length, 11);
  assert.equal(resolvePriorityCategory('Dépenses alimentaires'), 'Alimentaire');
  assert.equal(resolvePriorityCategory('Facture d’eau'), 'Eau');
  assert.equal(resolvePriorityCategory('Réparations et entretien de la maison'), 'Entretien de la maison');
});

test('every category has its own icon and only the three requested categories are extras', () => {
  const categories = [
    ...CATEGORY_ORDER.map((name, index) => ({
      id: index + 1,
      name,
      type: name === 'Compte boutique' ? 'boutique' : 'ordinary',
    })),
    { id: 16, name: 'Vacances en famille', type: 'ordinary' },
  ];
  const presentation = buildCategoryPresentation({
    monthKey: '2026-11',
    categories,
    expenses: [
      { category_id: 11, category_name: 'Alimentation des animaux', title: 'Croquettes', amount: 4500 },
      { category_id: 13, category_name: 'Dépenses exceptionnelles', title: 'Urgence', amount: 2000 },
      { category_id: 14, category_name: 'Depenses imprevues', title: 'Imprévu', amount: 1500 },
      { category_id: 16, category_name: 'Vacances en famille', title: 'Billet', amount: 8000 },
    ],
  });

  assert.deepEqual(
    presentation.extras.map((group) => group.heading),
    ['Alimentation des animaux', 'Dépenses exceptionnelles', 'Depenses imprevues'],
  );
  assert.equal(presentation.priority.length, CATEGORY_ORDER.length - OTHER_CATEGORY_NAMES.size + 1);
  assert.ok(presentation.priority.some((group) => group.heading === 'Vacances en famille'));
  const allGroups = [...presentation.priority, ...presentation.extras];
  assert.equal(new Set(allGroups.map((group) => group.icon)).size, allGroups.length);
  assert.equal(presentation.extras[0].total, 4500);
  assert.equal(presentation.extras[1].total, 2000);
  assert.equal(presentation.extras[2].total, 1500);
});

test('shop purchases are included in the shop category card', () => {
  const presentation = buildCategoryPresentation({
    monthKey: '2026-10',
    categories: [{ id: 3, name: 'Compte boutique', type: 'boutique' }],
    groceryItems: [
      { id: 8, item_name: 'Riz', total_amount: 12000 },
      { id: 9, item_name: 'Huile', total_amount: 7000 },
    ],
  });
  const shop = presentation.priority.find((group) => group.heading === 'Compte boutique');
  assert.equal(shop.total, 19000);
  assert.equal(shop.count, 2);
  assert.equal(shop.icon, '🛒');
});

test('categories follow the Render order and new categories sort alphabetically after them', () => {
  const categories = [
    { id: 1, name: 'Zèbres' },
    { id: 2, name: 'Transport' },
    { id: 3, name: 'Dépenses alimentaires' },
    { id: 4, name: 'Vacances' },
  ];
  assert.deepEqual(sortCategoriesForPresentation(categories).map((category) => category.name), [
    'Dépenses alimentaires',
    'Transport',
    'Vacances',
    'Zèbres',
  ]);
});

test('expense category display omits categories no longer present', () => {
  const categories = [
    { id: 1, name: 'Dépenses alimentaires' },
    { id: 3, name: 'Compte boutique', type: 'boutique' },
    { id: 12, name: 'Autres dépenses' },
    { id: 16, name: 'Vacances' },
    { id: 14, name: 'Depenses imprevues' },
  ];
  const presentation = filterCategoryPresentation(buildCategoryPresentation({
    monthKey: '2026-11',
    categories: [
      ...categories,
      { id: 2, name: 'Factures' },
      { id: 11, name: 'Alimentation des animaux' },
    ],
    expenses: [
      { category_id: 1, category_name: 'Dépenses alimentaires', title: 'Courses', amount: 1000 },
      { category_id: 2, category_name: 'Factures', title: 'Eau', amount: 500 },
      { category_id: 16, category_name: 'Vacances', title: 'Sortie', amount: 200 },
      { category_id: 11, category_name: 'Alimentation des animaux', title: 'Croquettes', amount: 300 },
    ],
    groceryItems: [{ id: 7, item_name: 'Riz', total_amount: 1200 }],
  }), categories);

  assert.deepEqual(presentation.priority.map((group) => group.heading), [
    'Dépenses alimentaires',
    'Compte boutique',
    'Autres dépenses',
    'Vacances',
  ]);
  assert.deepEqual(presentation.extras.map((group) => group.heading), ['Depenses imprevues']);
  assert.equal(presentation.priority.find((group) => group.heading === 'Dépenses alimentaires').total, 1000);
  assert.equal(presentation.priority.find((group) => group.heading === 'Compte boutique').total, 1200);
  assert.equal(presentation.extras[0].count, 0);
});
