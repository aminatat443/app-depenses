const CATEGORY_ORDER = [
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

const PRIORITY_CATEGORIES = [
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
];

const OTHER_CATEGORY_NAMES = new Set([
  'depenses exceptionnelles',
  'depenses imprevues',
  'alimentation des animaux',
]);

const ICONS = {
  'achats alimentaires complementaires': '🥖',
  'depenses alimentaires': '🍞',
  'compte boutique': '🛒',
  'scolarite et frais educatifs': '🎓',
  'hygiene et beaute': '💄',
  'aumones et dons': '🤝',
  'entretien de la maison': '🏡',
  reparations: '🔧',
  sante: '🩺',
  transport: '🚌',
  'alimentation des animaux': '🐾',
  'autres depenses': '📦',
  'depenses exceptionnelles': '⚠️',
  'depenses imprevues': '🌧️',
  factures: '🧾',
};

const ICON_HINTS = [
  [/animal|veterinaire|croquette/, '🐾'],
  [/aliment|nourriture|repas|cuisine/, '🍽️'],
  [/ecole|scolar|etude|education/, '🎒'],
  [/sante|medic|pharm|soin/, '🩺'],
  [/transport|voiture|carburant|taxi/, '🚗'],
  [/facture|eau|electric|internet|wifi/, '🧾'],
  [/maison|logement|entretien|reparation/, '🏠'],
  [/vetement|chaussure|mode/, '👕'],
  [/loisir|sport|voyage|vacance/, '🎭'],
  [/don|aumone|charite|aide/, '🎁'],
  [/enfant|bebe|poupon/, '🧸'],
  [/jardin|plante|fleur/, '🪴'],
];

const DYNAMIC_ICONS = [
  '🎯', '💡', '🧰', '🎨', '📚', '🧹', '🪙', '🛠️', '🎁', '📌',
  '🌟', '🧺', '🍳', '🧴', '🧯', '🪑', '🪄', '📦', '🔑', '🎵',
  '🧵', '🚲', '🪴', '🎮', '💼', '🧃', '🧊', '🧭', '🛍️',
];

function normalizeCategoryName(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLocaleLowerCase('fr')
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function resolvePriorityCategory(rawName) {
  const normalized = normalizeCategoryName(rawName);
  if (!normalized) return null;

  if (/^(depenses alimentaires|alimentation|aliments|nourriture|courses|alimentaire)$/.test(normalized)
    || normalized.includes('achats alimentaires')
    || normalized.includes('pain')
    || normalized.includes('ravitaillement alimentaire')) return 'Alimentaire';
  if (/facture.*eau|^eau$/.test(normalized)) return 'Eau';
  if (/facture.*electricite|^electricite$/.test(normalized)) return 'Électricité';
  if (/^(facture )?(wi ?fi|wifi|internet)(.*)?$/.test(normalized)
    || normalized.includes('internet') && normalized.includes('wi')) return 'Wifi';
  if (/^compte boutique$/.test(normalized)) return 'Compte boutique';
  if (/^(scolarite|scolarite et frais educatifs)/.test(normalized)
    || normalized.includes('frais educatifs')) return 'Scolarité et frais éducatifs';
  if (/hygiene|beaute|cosmetique/.test(normalized)) return 'Hygiène et beauté';
  if (/aumone|aumones|don(s)?|charite/.test(normalized)) return 'Aumône et dons';
  if (/reparation|reparations|entretien.*maison|maintenance.*maison/.test(normalized)) return 'Entretien de la maison';
  if (/^sante|sante|consultation|medicament|ordonnance|clinique|hopital/.test(normalized)) return 'Santé';
  if (/^transport/.test(normalized)) return 'Transport';
  return null;
}

function sortCategoriesForPresentation(categories) {
  const categoryOrder = new Map(CATEGORY_ORDER.map((name, index) => [normalizeCategoryName(name), index]));
  return [...(categories || [])].sort((left, right) => {
    const leftIndex = categoryOrder.get(normalizeCategoryName(left.name)) ?? CATEGORY_ORDER.length;
    const rightIndex = categoryOrder.get(normalizeCategoryName(right.name)) ?? CATEGORY_ORDER.length;
    if (leftIndex !== rightIndex) return leftIndex - rightIndex;
    return String(left.name || '').localeCompare(String(right.name || ''), 'fr', { sensitivity: 'base' });
  });
}

function hashCategoryName(name) {
  let hash = 0;
  for (const character of normalizeCategoryName(name)) {
    hash = (hash * 31 + character.codePointAt(0)) >>> 0;
  }
  return hash;
}

function chooseCategoryIcon(name, usedIcons) {
  const normalized = normalizeCategoryName(name);
  const mappedIcon = ICONS[normalized]
    || ICON_HINTS.find(([pattern]) => pattern.test(normalized))?.[1];
  if (mappedIcon && !usedIcons.has(mappedIcon)) return mappedIcon;

  const start = hashCategoryName(name) % DYNAMIC_ICONS.length;
  for (let offset = 0; offset < DYNAMIC_ICONS.length; offset += 1) {
    const icon = DYNAMIC_ICONS[(start + offset) % DYNAMIC_ICONS.length];
    if (!usedIcons.has(icon)) return icon;
  }
  return DYNAMIC_ICONS[start];
}

function createCategoryGroups(categories) {
  const groups = [];
  const byId = new Map();
  const byName = new Map();
  const usedIcons = new Set();
  const sortedCategories = sortCategoriesForPresentation(categories);

  for (const category of sortedCategories) {
    const heading = String(category.name || '').trim();
    if (!heading) continue;
    const icon = category.type === 'boutique' ? '🛒' : chooseCategoryIcon(heading, usedIcons);
    usedIcons.add(icon);
    const group = {
      heading,
      icon,
      total: 0,
      count: 0,
      details: [],
      categoryIds: category.id === undefined || category.id === null ? [] : [category.id],
      isBoutique: category.type === 'boutique' || normalizeCategoryName(heading) === 'compte boutique',
    };
    groups.push(group);
    byName.set(normalizeCategoryName(heading), group);
    if (category.id !== undefined && category.id !== null) byId.set(String(category.id), group);
  }
  return { groups, byId, byName, usedIcons };
}

function buildCategoryPresentation({ monthKey = '', categories = [], expenses = [], groceryItems = [] }) {
  const { groups, byId, byName, usedIcons } = createCategoryGroups(categories);
  const addExpense = (group, expense) => {
    group.total += Number(expense.amount || 0);
    group.count += 1;
    group.details.push(expense);
  };

  for (const expense of expenses) {
    const name = String(expense.category_name || expense.category || '').trim();
    if (!name) continue;
    let group = expense.category_id === undefined || expense.category_id === null
      ? null
      : byId.get(String(expense.category_id));
    group ||= byName.get(normalizeCategoryName(name));
    if (!group) {
      const icon = chooseCategoryIcon(name, usedIcons);
      usedIcons.add(icon);
      group = {
        heading: name,
        icon,
        total: 0,
        count: 0,
        details: [],
        categoryIds: expense.category_id === undefined || expense.category_id === null ? [] : [expense.category_id],
        isBoutique: normalizeCategoryName(name) === 'compte boutique',
      };
      groups.push(group);
      byName.set(normalizeCategoryName(name), group);
      if (expense.category_id !== undefined && expense.category_id !== null) {
        byId.set(String(expense.category_id), group);
      }
    }
    addExpense(group, expense);
  }

  const boutiqueGroup = groups.find((group) => group.isBoutique);
  if (boutiqueGroup) {
    for (const item of groceryItems) {
      addExpense(boutiqueGroup, {
        ...item,
        title: item.item_name || item.title || 'Ravitaillement',
        amount: item.amount ?? item.total_amount ?? 0,
        category_name: boutiqueGroup.heading,
        grocery: true,
      });
    }
  }

  return {
    monthKey,
    priority: groups.filter((group) => !OTHER_CATEGORY_NAMES.has(normalizeCategoryName(group.heading))),
    extras: groups.filter((group) => OTHER_CATEGORY_NAMES.has(normalizeCategoryName(group.heading))),
  };
}

function filterCategoryPresentation(presentation, categories) {
  const displayedCategories = categories || [];
  const displayedIds = new Set(displayedCategories
    .filter((category) => category.id !== undefined && category.id !== null)
    .map((category) => String(category.id)));
  const displayedNames = new Set(displayedCategories.map((category) => normalizeCategoryName(category.name)));
  const categoryExists = (id, name) => (
    (id !== undefined && id !== null && displayedIds.has(String(id)))
    || displayedNames.has(normalizeCategoryName(name))
  );
  const filteredGroups = (groups) => groups
    .map((group) => {
      const categoryIds = group.categoryIds.filter((id) => displayedIds.has(String(id)));
      const details = group.details.filter((item) => item.grocery
        ? displayedCategories.some((category) => category.type === 'boutique'
          || normalizeCategoryName(category.name) === 'compte boutique')
        : categoryExists(item.category_id, item.category_name || item.category));
      return {
        ...group,
        categoryIds,
        details,
        total: details.reduce((sum, item) => sum + Number(item.amount ?? item.total_amount ?? 0), 0),
        count: details.length,
      };
    })
    .filter((group) => group.categoryIds.length || displayedNames.has(normalizeCategoryName(group.heading)));

  return {
    ...presentation,
    priority: filteredGroups(presentation.priority),
    extras: filteredGroups(presentation.extras),
  };
}

module.exports = {
  CATEGORY_ORDER,
  PRIORITY_CATEGORIES,
  OTHER_CATEGORY_NAMES,
  ICONS,
  normalizeCategoryName,
  resolvePriorityCategory,
  sortCategoriesForPresentation,
  buildCategoryPresentation,
  filterCategoryPresentation,
};
