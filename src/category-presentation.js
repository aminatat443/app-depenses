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

const ICONS = {
  'Alimentaire': '🍞',
  'Eau': '💧',
  'Électricité': '⚡',
  'Wifi': '📶',
  'Compte boutique': '🛒',
  'Scolarité et frais éducatifs': '🎓',
  'Hygiène et beauté': '✨',
  'Aumône et dons': '🤝',
  'Entretien de la maison': '🏠',
  'Santé': '🩺',
  'Transport': '🚌',
};

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
  return [...(categories || [])].sort((left, right) => {
    const leftPriority = PRIORITY_CATEGORIES.indexOf(resolvePriorityCategory(left.name));
    const rightPriority = PRIORITY_CATEGORIES.indexOf(resolvePriorityCategory(right.name));
    const leftIndex = leftPriority >= 0 ? leftPriority : PRIORITY_CATEGORIES.length;
    const rightIndex = rightPriority >= 0 ? rightPriority : PRIORITY_CATEGORIES.length;
    if (leftIndex !== rightIndex) return leftIndex - rightIndex;
    if (leftIndex === PRIORITY_CATEGORIES.length) {
      return String(left.name || '').localeCompare(String(right.name || ''), 'fr', { sensitivity: 'base' });
    }
    return String(left.name || '').localeCompare(String(right.name || ''), 'fr', { sensitivity: 'base' });
  });
}

function buildCategoryPresentation({ monthKey = '', categories = [], expenses = [], groceryItems = [] }) {
  const byPriority = new Map(PRIORITY_CATEGORIES.map((name) => [name, {
    heading: name,
    icon: ICONS[name],
    total: 0,
    count: 0,
    details: [],
    categoryIds: [],
  }]));
  const extraNames = new Map();

  for (const category of categories) {
    const rawName = String(category.name || '').trim();
    if (!rawName || rawName === 'Compte boutique') continue;
    const priority = resolvePriorityCategory(rawName);
    if (!priority) {
      if (!extraNames.has(rawName)) extraNames.set(rawName, {
        heading: rawName,
        icon: '📦',
        total: 0,
        count: 0,
        details: [],
        categoryIds: [],
      });
      extraNames.get(rawName).categoryIds.push(category.id);
    } else {
      if (!byPriority.get(priority).categoryIds.includes(category.id)) {
        byPriority.get(priority).categoryIds.push(category.id);
      }
    }
  }

  for (const expense of expenses) {
    const rawName = String(expense.category_name || expense.category || '').trim();
    const priority = resolvePriorityCategory(rawName);
    if (priority) {
      const group = byPriority.get(priority);
      group.total += Number(expense.amount || 0);
      group.count += 1;
      group.details.push(expense);
      if (expense.category_id && !group.categoryIds.includes(expense.category_id)) group.categoryIds.push(expense.category_id);
    } else if (rawName) {
      const group = extraNames.get(rawName) || {
        heading: rawName,
        icon: '📦',
        total: 0,
        count: 0,
        details: [],
        categoryIds: [],
      };
      group.total += Number(expense.amount || 0);
      group.count += 1;
      group.details.push(expense);
      if (expense.category_id && !group.categoryIds.includes(expense.category_id)) group.categoryIds.push(expense.category_id);
      extraNames.set(rawName, group);
    }
  }

  for (const item of groceryItems) {
    const group = byPriority.get('Compte boutique');
    const amount = Number(item.total_amount || item.amount || 0);
    group.total += amount;
    group.count += 1;
    group.details.push({
      ...item,
      title: item.item_name || item.title || 'Ravitaillement',
      category_name: 'Compte boutique',
      grocery: true,
    });
  }

  const extras = [...extraNames.values()].sort((left, right) =>
    left.heading.localeCompare(right.heading, 'fr', { sensitivity: 'base' }));

  return {
    monthKey,
    priority: PRIORITY_CATEGORIES.map((name) => byPriority.get(name)),
    extras,
  };
}

function filterCategoryPresentation(presentation, categories) {
  const displayedCategories = categories || [];
  const categoryIsAvailable = (name) => displayedCategories.some((category) =>
    normalizeCategoryName(category.name) === normalizeCategoryName(name));
  const displayedDetails = (details) => details.filter((item) => categoryIsAvailable(
    item.grocery ? 'Compte boutique' : (item.category_name || item.category),
  ));
  const totalFor = (details) => details.reduce((sum, item) =>
    sum + Number(item.amount ?? item.total_amount ?? 0), 0);

  const priority = presentation.priority
    .map((group) => {
      const visibleCategories = displayedCategories.filter((category) => (
        group.heading === 'Compte boutique'
          ? normalizeCategoryName(category.name) === normalizeCategoryName(group.heading)
          : resolvePriorityCategory(category.name) === group.heading
      ));
      const details = displayedDetails(group.details);
      return {
        ...group,
        categoryIds: group.categoryIds.filter((id) =>
          visibleCategories.some((category) => Number(category.id) === Number(id))),
        details,
        total: totalFor(details),
        count: details.length,
        isDisplayed: visibleCategories.length > 0,
      };
    })
    .filter((group) => group.isDisplayed)
    .map(({ isDisplayed, ...group }) => group);

  const extras = presentation.extras
    .map((group) => {
      const details = displayedDetails(group.details);
      return {
        ...group,
        details,
        total: totalFor(details),
        count: details.length,
      };
    })
    .filter((group) => categoryIsAvailable(group.heading));

  return { ...presentation, priority, extras };
}

module.exports = {
  PRIORITY_CATEGORIES,
  ICONS,
  normalizeCategoryName,
  resolvePriorityCategory,
  sortCategoriesForPresentation,
  buildCategoryPresentation,
  filterCategoryPresentation,
};
