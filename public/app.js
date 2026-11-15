(() => {
  const syncThemeToggles = () => {
    const isDark = document.documentElement.dataset.theme === 'dark';
    document.querySelectorAll('[data-theme-toggle]').forEach((button) => {
      button.setAttribute('aria-pressed', String(isDark));
      button.setAttribute('aria-label', `Activer le mode ${isDark ? 'clair' : 'sombre'}`);
      const label = button.querySelector('[data-theme-label]');
      if (label) label.textContent = `Mode ${isDark ? 'clair' : 'sombre'}`;
    });
    document.querySelector('meta[name="theme-color"]')?.setAttribute('content', isDark ? '#111318' : '#fafafa');
  };
  syncThemeToggles();
  document.addEventListener('click', (event) => {
    if (!(event.target instanceof Element)) return;
    const toggle = event.target.closest('[data-theme-toggle]');
    if (!toggle) return;
    const nextTheme = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
    if (nextTheme === 'dark') document.documentElement.dataset.theme = 'dark';
    else delete document.documentElement.dataset.theme;
    localStorage.setItem('maison-theme', nextTheme);
    syncThemeToggles();
  });

  const loginForm = document.querySelector('.login-form');
  loginForm?.addEventListener('submit', () => {
    const button = loginForm.querySelector('.login-submit');
    const label = loginForm.querySelector('[data-login-label]');
    if (!button || !label) return;
    button.disabled = true;
    button.classList.add('is-loading');
    button.setAttribute('aria-busy', 'true');
    label.textContent = 'Connexion en cours…';
  });

  document.querySelector('[data-toggle-password]')?.addEventListener('click', (event) => {
    const button = event.currentTarget;
    const password = document.getElementById('login-password');
    if (!password) return;
    const visible = password.type === 'password';
    password.type = visible ? 'text' : 'password';
    button.textContent = visible ? 'Masquer' : 'Afficher';
    button.setAttribute('aria-label', visible ? 'Masquer le mot de passe' : 'Afficher le mot de passe');
    button.setAttribute('aria-pressed', String(visible));
  });

  const globalAlerts = document.querySelectorAll(
    '.login-notice, .login-error, .income-notice, .account-feedback',
  );
  if (globalAlerts.length) {
    const alertStack = document.createElement('div');
    alertStack.className = 'floating-alert-stack';
    alertStack.setAttribute('aria-label', 'Notifications');
    globalAlerts.forEach((alert) => {
      const isError = alert.classList.contains('login-error')
        || alert.classList.contains('account-feedback-error');
      alert.classList.add('floating-alert', isError ? 'floating-alert-error' : 'floating-alert-success');
      const closeButton = document.createElement('button');
      closeButton.type = 'button';
      closeButton.className = 'floating-alert-close';
      closeButton.setAttribute('aria-label', 'Fermer la notification');
      closeButton.textContent = '×';
      closeButton.addEventListener('click', () => alert.remove());
      alert.append(closeButton);
      alertStack.append(alert);
    });
    document.body.append(alertStack);
  }

  const isZeroNumber = (input) => /^[-+]?0(?:[.,]0*)?$/.test(input.value.trim());
  const clearZeroBeforeEntry = (input) => {
    if (input instanceof HTMLInputElement && input.type === 'number' && isZeroNumber(input)) {
      input.value = '';
    }
  };
  document.addEventListener('keydown', (event) => {
    const input = event.target;
    if (!(input instanceof HTMLInputElement) || input.type !== 'number' || event.ctrlKey || event.metaKey || event.altKey) return;
    if (/^[0-9.,+-]$/.test(event.key) || event.key === 'Backspace' || event.key === 'Delete') {
      clearZeroBeforeEntry(input);
    }
  });
  document.addEventListener('beforeinput', (event) => {
    const input = event.target;
    if (event.inputType.startsWith('insert')) clearZeroBeforeEntry(input);
  });
  document.addEventListener('paste', (event) => clearZeroBeforeEntry(event.target));

  window.addEventListener('pageshow', (event) => {
    if (event.persisted && document.body.dataset.privatePage === 'true') window.location.reload();
  });

  let lockedPageScrollY = 0;
  let savedBodyScrollStyles = null;
  const syncPageScrollLock = () => {
    const shouldLock = document.body.classList.contains('modal-open')
      || document.body.classList.contains('mobile-nav-open');
    if (shouldLock && !savedBodyScrollStyles) {
      lockedPageScrollY = window.scrollY;
      savedBodyScrollStyles = {
        position: document.body.style.position,
        top: document.body.style.top,
        left: document.body.style.left,
        right: document.body.style.right,
        width: document.body.style.width,
        overflow: document.body.style.overflow,
      };
      document.body.style.position = 'fixed';
      document.body.style.top = `-${lockedPageScrollY}px`;
      document.body.style.left = '0';
      document.body.style.right = '0';
      document.body.style.width = '100%';
      document.body.style.overflow = 'hidden';
    } else if (!shouldLock && savedBodyScrollStyles) {
      Object.assign(document.body.style, savedBodyScrollStyles);
      savedBodyScrollStyles = null;
      window.scrollTo(0, lockedPageScrollY);
    }
  };

  const enhanceTables = (root = document) => {
    root.querySelectorAll('.transactions-table').forEach((table) => {
      const headers = Array.from(table.querySelectorAll('thead th'), (cell) => cell.textContent.trim());
      const rows = table.querySelectorAll('tbody tr');
      if (headers.length) {
        table.dataset.responsiveCards = 'true';
        rows.forEach((row) => {
          Array.from(row.cells).forEach((cell, index) => {
            if (cell.tagName !== 'TH' && cell.colSpan === 1 && headers[index]) {
              cell.dataset.label ||= headers[index];
            }
          });
        });
        return;
      }

      rows.forEach((row) => {
        const rowHeading = row.querySelector('th')?.textContent.trim();
        if (!rowHeading) return;
        table.dataset.responsiveCards = 'true';
        Array.from(row.cells).forEach((cell) => {
          if (cell.tagName !== 'TH' && cell.colSpan === 1) cell.dataset.label ||= rowHeading;
        });
      });
    });
  };
  enhanceTables();

  const sidebar = document.querySelector('.sidebar');
  const mobileMenu = document.getElementById('mobile-more-menu');
  const mobileMenuTrigger = document.querySelector('.mobile-more-trigger');
  const mobileMenuOverlay = document.querySelector('.mobile-nav-overlay');
  let mobileMenuReturnFocus = null;

  const setMobileMenu = (open) => {
    if (!sidebar || !mobileMenu || !mobileMenuTrigger || !mobileMenuOverlay) return;
    sidebar.classList.toggle('mobile-menu-open', open);
    mobileMenuTrigger.setAttribute('aria-expanded', String(open));
    mobileMenuTrigger.setAttribute('aria-label', open ? 'Fermer les autres sections' : 'Ouvrir les autres sections');
    mobileMenu.setAttribute('aria-hidden', String(!open));
    mobileMenu.inert = !open;
    mobileMenuOverlay.hidden = !open;
    document.body.classList.toggle('mobile-nav-open', open);
    syncPageScrollLock();
    if (open) {
      mobileMenuReturnFocus = mobileMenuTrigger;
      mobileMenu.querySelector('.mobile-menu-close')?.focus();
    } else {
      if (mobileMenuReturnFocus instanceof HTMLElement && mobileMenuReturnFocus.isConnected) {
        mobileMenuReturnFocus.focus();
      }
      mobileMenuReturnFocus = null;
    }
  };

  mobileMenuTrigger?.addEventListener('click', () => {
    setMobileMenu(mobileMenuTrigger.getAttribute('aria-expanded') !== 'true');
  });
  mobileMenuOverlay?.addEventListener('click', () => setMobileMenu(false));
  mobileMenu?.addEventListener('click', (event) => {
    if (event.target.closest('[data-mobile-nav-close], a, form button')) {
      setMobileMenu(false);
    }
  });

  const modal = document.getElementById('app-modal');
  const modalContent = document.getElementById('modal-content');
  const modalPanel = modal?.querySelector('.modal-panel');
  let modalReturnFocus = null;
  let unsavedDialog = null;
  let unsavedReturnFocus = null;

  const focusableElements = (container) => Array.from(container.querySelectorAll(
    'a[href], button:not([disabled]), input:not([disabled]):not([type="hidden"]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])',
  )).filter((element) => element.getClientRects().length > 0 && !element.closest('[inert]'));

  const dirtyForm = () => modalContent?.querySelector('form[data-modal-dirty="true"]');

  const configureGroceryEntry = (form) => {
    const unitPrice = form.querySelector('input[name="unit_price"]');
    const totalAmount = form.querySelector('input[name="total_amount"]');
    if (!unitPrice || !totalAmount) return;
    const updateRequiredFields = () => {
      unitPrice.required = totalAmount.value.trim() === '';
      totalAmount.required = unitPrice.value.trim() === '';
    };
    unitPrice.addEventListener('input', updateRequiredFields);
    totalAmount.addEventListener('input', updateRequiredFields);
    updateRequiredFields();
  };

  const closeUnsavedDialog = () => {
    unsavedDialog?.remove();
    unsavedDialog = null;
    const target = unsavedReturnFocus;
    unsavedReturnFocus = null;
    (target instanceof HTMLElement && target.isConnected
      ? target
      : modalContent?.querySelector('input:not([type="hidden"]), select, textarea, button'))?.focus();
  };

  const closeModal = (force = false) => {
    if (!modal || !modalContent) return;
    if (unsavedDialog) {
      if (force) {
        unsavedDialog.remove();
        unsavedDialog = null;
      } else {
        closeUnsavedDialog();
        return;
      }
    }
    if (!force && dirtyForm()) {
      const layer = document.createElement('div');
      layer.className = 'modal-unsaved-backdrop';
      const dialog = document.createElement('section');
      dialog.className = 'modal-unsaved-panel';
      dialog.setAttribute('role', 'alertdialog');
      dialog.setAttribute('aria-modal', 'true');
      dialog.setAttribute('aria-labelledby', 'modal-unsaved-title');
      dialog.innerHTML = '<h2 id="modal-unsaved-title">Modifications non enregistrées</h2><p>Vos changements seront perdus si vous fermez cette fenêtre.</p>';
      const actions = document.createElement('div');
      actions.className = 'modal-unsaved-actions';
      const keepEditing = document.createElement('button');
      keepEditing.type = 'button';
      keepEditing.textContent = 'Continuer la saisie';
      keepEditing.dataset.keepEditing = '';
      const discard = document.createElement('button');
      discard.type = 'button';
      discard.textContent = 'Fermer sans enregistrer';
      discard.dataset.discardChanges = '';
      actions.append(keepEditing, discard);
      dialog.append(actions);
      layer.append(dialog);
      unsavedReturnFocus = document.activeElement;
      modal.append(layer);
      unsavedDialog = layer;
      keepEditing.focus();
      return;
    }

    modal.classList.remove('open');
    modal.setAttribute('aria-hidden', 'true');
    document.body.classList.remove('modal-open');
    modalContent.replaceChildren();
    modalPanel?.removeAttribute('aria-labelledby');
    syncPageScrollLock();
    if (modalReturnFocus instanceof HTMLElement && modalReturnFocus.isConnected) modalReturnFocus.focus();
    modalReturnFocus = null;
  };

  const openModal = (content, returnFocus = modalReturnFocus || document.activeElement) => {
    if (!modal || !modalContent) return;
    unsavedDialog?.remove();
    unsavedDialog = null;
    unsavedReturnFocus = null;
    modalReturnFocus = returnFocus;
    modalContent.replaceChildren(content);
    const groceryForm = modalContent.querySelector('form[action="/groceries"]');
    if (groceryForm) configureGroceryEntry(groceryForm);
    const title = modalContent.querySelector('h1, h2');
    if (title) {
      title.id ||= 'modal-title';
      modalPanel?.setAttribute('aria-labelledby', title.id);
    }
    modal.classList.add('open');
    modal.setAttribute('aria-hidden', 'false');
    document.body.classList.add('modal-open');
    syncPageScrollLock();
    const initialFocus = modalContent.querySelector(
      'input:not([type="hidden"]):not([disabled]), select:not([disabled]), textarea:not([disabled]), button:not([disabled])',
    ) || modal.querySelector('.close-modal');
    initialFocus?.focus();
  };

  const showError = (message) => {
    if (!modal || !modalContent) return;
    const error = document.createElement('div');
    error.className = 'modal-error';
    error.setAttribute('role', 'alert');
    error.tabIndex = -1;
    error.textContent = message;
    modalContent.querySelector('.modal-error')?.remove();
    modalContent.prepend(error);
    modal.classList.add('open');
    modal.setAttribute('aria-hidden', 'false');
    document.body.classList.add('modal-open');
    syncPageScrollLock();
    error.focus?.();
  };

  const configureIncomeEntry = (form) => {
    const type = form.querySelector('[data-income-type]')?.value || 'other';
    const previousType = form.dataset.configuredIncomeType;
    const otherFields = form.querySelector('[data-other-income-fields]');
    const isOther = type === 'other';
    form.action = isOther ? '/income' : '/income/monthly';
    if (otherFields) {
      otherFields.hidden = !isOther;
      otherFields.querySelectorAll('input, select').forEach((field) => {
        field.disabled = !isOther;
      });
    }

    const amount = form.elements.namedItem('amount');
    const date = form.elements.namedItem('income_date');
    if (amount) amount.min = isOther ? '1' : '0';
    if (date) date.required = isOther;
    if (!isOther && amount && previousType !== type) {
      const currentAmount = form.elements.namedItem(type === 'pape' ? 'pape_amount' : 'pension_amount');
      const currentValue = Number(currentAmount?.value || 0);
      amount.value = type === 'pension' && currentValue === 0
        ? form.dataset.defaultPension || '180000'
        : String(currentValue);
      const currentDate = form.elements.namedItem(type === 'pape' ? 'pape_received_date' : 'pension_received_date');
      date.value = currentDate?.value || '';
    }
    form.dataset.configuredIncomeType = type;
  };

  const prepareIncomeSubmission = (form) => {
    if (form.matches('[data-income-entry-form]')) {
      const type = form.querySelector('[data-income-type]')?.value || 'other';
      if (type !== 'other') {
        const amount = form.elements.namedItem('amount').value;
        const date = form.elements.namedItem('income_date').value;
        const amountField = form.elements.namedItem(type === 'pape' ? 'pape_amount' : 'pension_amount');
        const dateField = form.elements.namedItem(type === 'pape' ? 'pape_received_date' : 'pension_received_date');
        amountField.value = amount;
        dateField.value = date;
      }
    }

    if (form.matches('[data-income-monthly-form]')) {
      const source = form.elements.namedItem('income_source').value;
      const isPape = source === 'Pape';
      form.elements.namedItem(isPape ? 'pape_amount' : 'pension_amount').value =
        form.elements.namedItem('amount').value;
      form.elements.namedItem(isPape ? 'pape_received_date' : 'pension_received_date').value =
        form.elements.namedItem('received_date').value;
      form.elements.namedItem('comment').value = form.elements.namedItem('source_comment').value;
    }
  };

  const confirmModal = (trigger) => {
    const wrapper = document.createElement('div');
    wrapper.className = 'modal-confirmation';
    const heading = document.createElement('div');
    heading.className = 'modal-heading';
    const eyebrow = document.createElement('span');
    eyebrow.className = 'eyebrow';
    eyebrow.textContent = 'CONFIRMATION';
    const title = document.createElement('h2');
    title.textContent = trigger.dataset.modalTitle || 'Confirmer cette action';
    const message = document.createElement('p');
    message.textContent = trigger.dataset.modalMessage || 'Voulez-vous continuer ?';
    heading.append(eyebrow, title, message);

    const form = document.createElement('form');
    form.method = 'POST';
    form.action = trigger.dataset.modalConfirm;
    form.className = 'modal-confirm-form';
    if (trigger.dataset.modalReturn) {
      const returnTo = document.createElement('input');
      returnTo.type = 'hidden';
      returnTo.name = 'return_to';
      returnTo.value = trigger.dataset.modalReturn;
      form.append(returnTo);
    }
    if (trigger.dataset.modalFieldName && trigger.dataset.modalFieldValue !== undefined) {
      const field = document.createElement('input');
      field.type = 'hidden';
      field.name = trigger.dataset.modalFieldName;
      field.value = trigger.dataset.modalFieldValue;
      form.append(field);
    }
    const actions = document.createElement('div');
    actions.className = 'modal-actions';
    const cancel = document.createElement('button');
    cancel.type = 'button';
    cancel.className = 'secondary-btn';
    cancel.dataset.closeModal = '';
    cancel.textContent = 'Annuler';
    const confirm = document.createElement('button');
    confirm.type = 'submit';
    confirm.className = trigger.dataset.modalDanger === 'true' ? 'danger' : 'primary-btn';
    confirm.textContent = trigger.dataset.modalSubmit || 'Confirmer';
    actions.append(cancel, confirm);
    form.append(actions);
    wrapper.append(heading, form);
    openModal(wrapper, trigger);
  };

  const setFormBusy = (form, busy) => {
    form.querySelectorAll('button[type="submit"], input[type="submit"]').forEach((button) => {
      button.disabled = busy;
      button.setAttribute('aria-busy', String(busy));
      if (busy && !button.querySelector('.submit-spinner')) {
        const spinner = document.createElement('span');
        spinner.className = 'submit-spinner';
        spinner.setAttribute('aria-hidden', 'true');
        button.prepend(spinner);
      } else if (!busy) {
        button.querySelector('.submit-spinner')?.remove();
      }
    });
  };

  const submitModalForm = async (form) => {
    const method = form.method.toUpperCase();
    const formData = new FormData(form);
    const isMultipart = form.enctype === 'multipart/form-data';
    const body = isMultipart ? formData : new URLSearchParams(formData);
    setFormBusy(form, true);
    try {
      const response = await fetch(form.action || window.location.href, {
        method,
        body,
        credentials: 'same-origin',
      });
      if (!response.ok) {
        const responseText = await response.text();
        const errorDocument = new DOMParser().parseFromString(responseText, 'text/html');
        const message = errorDocument.querySelector('.page-alert,[role="alert"]')?.textContent.trim();
        throw new Error(message || `L’action a échoué (${response.status}).`);
      }

      if (response.redirected && new URL(response.url).pathname === '/login') {
        window.location.assign(response.url);
        return;
      }
      if (response.redirected && new URL(response.url).pathname !== window.location.pathname) {
        const html = await response.text();
        const doc = new DOMParser().parseFromString(html, 'text/html');
        const content = doc.querySelector('.page-content') || doc.querySelector('main');
        if (!content) throw new Error('La page demandée est indisponible.');
        openModal(content.cloneNode(true), modalReturnFocus);
        return;
      }
      if (response.redirected && new URL(response.url).href !== window.location.href) {
        window.location.assign(response.url);
        return;
      }

      closeModal(true);
      window.location.reload();
    } catch (error) {
      setFormBusy(form, false);
      showError(error instanceof Error ? error.message : 'Une erreur inattendue est survenue.');
    }
  };

  document.addEventListener('click', (event) => {
    if (!(event.target instanceof Element)) return;
    const row = event.target.closest('tr[data-row-href]');
    if (row && !event.target.closest('a, button, input, select, textarea, form')) {
      window.location.assign(row.dataset.rowHref);
      return;
    }

    const trigger = event.target.closest('[data-modal-template], [data-modal-confirm], [data-close-modal]');
    if (!trigger || !modal || !modalContent) return;

    if (trigger.hasAttribute('data-close-modal')) {
      closeModal();
      return;
    }
    if (trigger.dataset.modalConfirm) {
      confirmModal(trigger);
      return;
    }

    const template = document.getElementById(trigger.dataset.modalTemplate);
    if (!(template instanceof HTMLTemplateElement)) return;
    const content = template.content.cloneNode(true);
    if (template.id === 'income-monthly-edit') {
      const form = content.querySelector('[data-income-monthly-form]');
      const source = trigger.dataset.incomeEditSource || '';
      if (form) {
        form.elements.namedItem('income_source').value = source;
        form.elements.namedItem('amount').value = trigger.dataset.incomeEditAmount || '0';
        form.elements.namedItem('received_date').value = trigger.dataset.incomeEditDate || '';
        form.elements.namedItem('source_comment').value = trigger.dataset.incomeEditComment || '';
        const title = content.querySelector('[data-income-edit-title]');
        if (title) title.textContent = `Modifier ${source}`;
      }
    }
    if (template.id === 'income-create') {
      const form = content.querySelector('[data-income-entry-form]');
      if (form) configureIncomeEntry(form);
    }
    if (template.id === 'export-confirmation') {
      const year = document.getElementById('export-year')?.value || new Date().getFullYear();
      const heading = content.querySelector('h2');
      const download = content.querySelector('a[href^="/export/"]');
      if (heading) heading.textContent = `Exporter le bilan ${year}`;
      if (download) download.href = `/export/${encodeURIComponent(year)}`;
    }
    openModal(content, trigger);
  });

  document.addEventListener('keydown', (event) => {
    if ((event.key !== 'Enter' && event.key !== ' ') || !(event.target instanceof HTMLTableRowElement)) return;
    const row = event.target;
    if (!row.dataset.rowHref && !row.dataset.modalTemplate) return;
    event.preventDefault();
    if (row.dataset.rowHref) {
      window.location.assign(row.dataset.rowHref);
    } else {
      row.click();
    }
  });

  modal?.addEventListener('click', (event) => {
    if (event.target === modal.querySelector('.modal-backdrop')) closeModal();
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (target.closest('[data-keep-editing]')) {
      closeUnsavedDialog();
    } else if (target.closest('[data-discard-changes]')) {
      closeModal(true);
    }
  });

  modalContent?.addEventListener('input', (event) => {
    const form = event.target.closest('form');
    if (form) form.dataset.modalDirty = 'true';
  });
  modalContent?.addEventListener('change', (event) => {
    const target = event.target;
    const form = target.closest('form');
    if (form) form.dataset.modalDirty = 'true';
    if (target.matches('[data-income-type]')) {
      const incomeForm = target.closest('[data-income-entry-form]');
      if (incomeForm) configureIncomeEntry(incomeForm);
    }
  });
  modalContent?.addEventListener('submit', (event) => {
    if (!(event.target instanceof HTMLFormElement) || event.target.method.toUpperCase() === 'GET') return;
    event.preventDefault();
    prepareIncomeSubmission(event.target);
    submitModalForm(event.target);
  });

  document.addEventListener('keydown', (event) => {
    if (modal?.classList.contains('open')) {
      if (event.key === 'Escape') {
        event.preventDefault();
        if (unsavedDialog) closeUnsavedDialog();
        else closeModal();
        return;
      }
      if (event.key === 'Tab') {
        const trap = unsavedDialog?.querySelector('.modal-unsaved-panel') || modalPanel;
        const focusable = trap ? focusableElements(trap) : [];
        if (!focusable.length) {
          event.preventDefault();
          trap?.focus();
          return;
        }
        const first = focusable[0];
        const last = focusable[focusable.length - 1];
        if (!trap?.contains(document.activeElement)) {
          event.preventDefault();
          (event.shiftKey ? last : first).focus();
        } else if (event.shiftKey && document.activeElement === first) {
          event.preventDefault();
          last.focus();
        } else if (!event.shiftKey && document.activeElement === last) {
          event.preventDefault();
          first.focus();
        }
      }
      return;
    }

    if (sidebar?.classList.contains('mobile-menu-open')) {
      if (event.key === 'Escape') {
        event.preventDefault();
        setMobileMenu(false);
      } else if (event.key === 'Tab') {
        const focusable = mobileMenu ? focusableElements(mobileMenu) : [];
        if (!focusable.length) return;
        const first = focusable[0];
        const last = focusable[focusable.length - 1];
        if (!mobileMenu?.contains(document.activeElement)) {
          event.preventDefault();
          (event.shiftKey ? last : first).focus();
        } else if (event.shiftKey && document.activeElement === first) {
          event.preventDefault();
          last.focus();
        } else if (!event.shiftKey && document.activeElement === last) {
          event.preventDefault();
          first.focus();
        }
      }
    }
  });

  const liveSearchControllers = new WeakMap();
  const liveFilterControllers = new WeakMap();
  const liveSearchTimers = new WeakMap();
  let globalSearchController;
  let globalSearchTimer;
  const normalizeSearch = (value) => String(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLocaleLowerCase();
  const typeLabels = {
    expense: 'Dépense',
    grocery: 'Compte boutique',
    income: 'Recette',
    rent: 'Loyer encaissé',
    savings: 'Épargne',
  };

  const setLiveSearchStatus = (form, message, isError = false) => {
    let status = form.nextElementSibling;
    if (!status?.classList.contains('live-search-status')) {
      status = document.createElement('p');
      status.className = 'live-search-status';
      status.setAttribute('aria-live', 'polite');
      form.insertAdjacentElement('afterend', status);
    }
    status.textContent = message;
    status.hidden = !message;
    status.setAttribute('role', isError ? 'alert' : 'status');
  };

  const formUrl = (form) => {
    const params = new URLSearchParams();
    for (const [name, value] of new FormData(form).entries()) {
      if (typeof value === 'string' && value.trim()) params.set(name, value.trim());
    }
    const url = new URL(form.action, window.location.href);
    url.search = params.toString();
    return url;
  };

  const filterCurrentPageItems = (query) => {
    const normalizedQuery = normalizeSearch(query);
    const rows = [...document.querySelectorAll('.main-area tbody tr, .page-content tbody tr')]
      .filter((row) => !row.classList.contains('global-search-empty-row'));
    const cards = [...document.querySelectorAll('.income-entry-list .income-entry')];
    if (!normalizedQuery) {
      rows.forEach((row) => { row.hidden = false; });
      cards.forEach((card) => { card.hidden = false; });
      document.querySelectorAll('.global-search-empty-row').forEach((row) => row.remove());
      const dashboardTable = document.getElementById('dashboard-recent-table');
      const dashboardEmpty = document.getElementById('dashboard-search-empty');
      if (dashboardTable && dashboardEmpty) {
        dashboardTable.hidden = false;
        dashboardEmpty.hidden = true;
      }
      return 0;
    }
    let matchCount = 0;
    [...rows, ...cards].forEach((item) => {
      const text = item.dataset.operationSearch || item.textContent;
      const matches = normalizeSearch(text).includes(normalizedQuery);
      item.hidden = !matches;
      if (matches) matchCount += 1;
    });

    const bodies = new Set(rows.map((row) => row.parentElement).filter(Boolean));
    bodies.forEach((body) => {
      let emptyRow = body.querySelector('.global-search-empty-row');
      const visibleRows = [...body.querySelectorAll('tr:not(.global-search-empty-row)')].filter((row) => !row.hidden);
      if (!emptyRow) {
        emptyRow = document.createElement('tr');
        emptyRow.className = 'global-search-empty-row';
        const cell = document.createElement('td');
        cell.colSpan = Math.max(1, body.closest('table')?.querySelectorAll('thead th').length || 1);
        cell.textContent = 'Aucun résultat sur cette page.';
        emptyRow.append(cell);
        body.append(emptyRow);
      }
      emptyRow.hidden = !normalizedQuery || visibleRows.length > 0;
    });

    const dashboardTable = document.getElementById('dashboard-recent-table');
    const dashboardEmpty = document.getElementById('dashboard-search-empty');
    if (dashboardTable && dashboardEmpty) {
      const dashboardRows = [...dashboardTable.querySelectorAll('[data-dashboard-operation]')];
      const dashboardMatches = dashboardRows.filter((row) => !row.hidden).length;
      dashboardTable.hidden = dashboardMatches === 0;
      dashboardEmpty.hidden = !normalizedQuery || dashboardMatches > 0;
    }
    return matchCount;
  };

  const updatePageSearch = async (form) => {
    liveSearchControllers.get(form)?.abort();
    const controller = new AbortController();
    liveSearchControllers.set(form, controller);
    const url = formUrl(form);
    const currentResults = document.querySelector('.page-table-card');
    currentResults?.setAttribute('aria-busy', 'true');
    setLiveSearchStatus(form, 'Recherche en cours…');
    try {
      const response = await fetch(url, {
        headers: { Accept: 'text/html' },
        signal: controller.signal,
      });
      if (!response.ok) throw new Error(`La recherche a échoué (${response.status}).`);
      const page = new DOMParser().parseFromString(await response.text(), 'text/html');
      const nextResults = page.querySelector('.page-table-card');
      if (!currentResults || !nextResults) throw new Error('Les résultats de recherche sont indisponibles.');
      if (liveSearchControllers.get(form) !== controller) return;
      currentResults.replaceWith(nextResults);
      enhanceTables(nextResults);
      filterCurrentPageItems(document.querySelector('#global-search input')?.value || '');
      window.history.replaceState(window.history.state, '', `${url.pathname}${url.search}${url.hash}`);
      setLiveSearchStatus(form, '');
    } catch (error) {
      if (error.name === 'AbortError' || liveSearchControllers.get(form) !== controller) return;
      console.error(error);
      setLiveSearchStatus(form, 'La recherche est momentanément indisponible. Réessayez.', true);
    } finally {
      currentResults?.removeAttribute('aria-busy');
      document.querySelector('.page-table-card')?.removeAttribute('aria-busy');
    }
  };

  const updatePageFilter = async (form) => {
    liveFilterControllers.get(form)?.abort();
    const controller = new AbortController();
    liveFilterControllers.set(form, controller);
    const selector = form.dataset.liveFilterRoot;
    const currentRoot = selector ? form.closest(selector) : null;
    if (!selector || !currentRoot) {
      setLiveSearchStatus(form, 'Le filtre ne peut pas être appliqué.', true);
      return;
    }
    const globalQuery = document.querySelector('#global-search input')?.value || '';
    const url = formUrl(form);
    currentRoot.setAttribute('aria-busy', 'true');
    setLiveSearchStatus(form, 'Mise à jour des résultats…');
    try {
      const response = await fetch(url, {
        headers: { Accept: 'text/html' },
        signal: controller.signal,
      });
      if (!response.ok) throw new Error(`Le filtre a échoué (${response.status}).`);
      const page = new DOMParser().parseFromString(await response.text(), 'text/html');
      const nextRoot = page.querySelector(selector);
      if (!nextRoot) throw new Error('Les résultats filtrés sont indisponibles.');
      if (liveFilterControllers.get(form) !== controller) return;
      const nextSearchInput = nextRoot.querySelector('#global-search input');
      if (nextSearchInput) nextSearchInput.value = globalQuery;
      currentRoot.replaceWith(nextRoot);
      enhanceTables(nextRoot);
      filterCurrentPageItems(globalQuery);
      if (globalQuery) runGlobalSearch(nextRoot.querySelector('#global-search input'));
      window.history.replaceState(window.history.state, '', `${url.pathname}${url.search}${url.hash}`);
    } catch (error) {
      if (error.name === 'AbortError' || liveFilterControllers.get(form) !== controller) return;
      console.error(error);
      setLiveSearchStatus(form, 'Le filtre est momentanément indisponible. Réessayez.', true);
    } finally {
      currentRoot.removeAttribute('aria-busy');
      document.querySelector(selector)?.removeAttribute('aria-busy');
    }
  };

  const renderGlobalSearchResults = (container, operations, totalCount, query, localMatchCount) => {
    container.replaceChildren();
    if (!operations.length) {
      const message = document.createElement('p');
      message.className = 'global-search-message';
      message.textContent = localMatchCount
        ? 'Les correspondances de cette page sont affichées directement.'
        : 'Aucune opération trouvée.';
      container.append(message);
      return;
    }

    const list = document.createElement('ul');
    list.className = 'global-search-list';
    operations.forEach((operation) => {
      const item = document.createElement('li');
      const link = document.createElement('a');
      link.className = 'global-search-result';
      link.href = operation.href;
      const details = document.createElement('span');
      details.className = 'global-search-result-details';
      const title = document.createElement('strong');
      title.textContent = operation.title;
      const subtitle = document.createElement('small');
      const date = operation.date ? new Date(operation.date) : null;
      const dateLabel = date && Number.isFinite(date.getTime()) ? date.toLocaleDateString('fr-FR') : '';
      subtitle.textContent = [operation.category, typeLabels[operation.type] || 'Opération', dateLabel]
        .filter(Boolean)
        .join(' · ');
      details.append(title, subtitle);
      const amount = document.createElement('strong');
      amount.className = 'global-search-result-amount';
      amount.textContent = `${Number(operation.amount).toLocaleString('fr-FR')} FCFA`;
      link.append(details, amount);
      item.append(link);
      list.append(item);
    });
    container.append(list);
    if (totalCount > operations.length) {
      const more = document.createElement('a');
      more.className = 'global-search-more';
      more.href = `/history?search=${encodeURIComponent(query)}`;
      more.textContent = `Voir les ${totalCount - operations.length} autres résultats dans l’historique`;
      container.append(more);
    }
  };

  const runGlobalSearch = async (input) => {
    const query = input.value.trim();
    const form = input.closest('#global-search');
    const results = document.getElementById('global-search-results');
    if (!form || !results) return;
    globalSearchController?.abort();
    if (!query) {
      results.replaceChildren();
      results.hidden = true;
      form.setAttribute('aria-expanded', 'false');
      filterCurrentPageItems('');
      return;
    }
    const localMatchCount = filterCurrentPageItems(query);
    const controller = new AbortController();
    globalSearchController = controller;
    results.replaceChildren();
    const loading = document.createElement('p');
    loading.className = 'global-search-message';
    loading.textContent = 'Recherche en cours…';
    results.append(loading);
    results.hidden = false;
    form.setAttribute('aria-expanded', 'true');
    try {
      const url = new URL('/history', window.location.origin);
      url.searchParams.set('format', 'json');
      url.searchParams.set('search', query);
      const response = await fetch(url, {
        headers: { Accept: 'application/json' },
        signal: controller.signal,
      });
      if (!response.ok) throw new Error(`La recherche a échoué (${response.status}).`);
      const data = await response.json();
      if (globalSearchController !== controller || input.value.trim() !== query) return;
      renderGlobalSearchResults(results, data.operations, data.totalCount, query, localMatchCount);
    } catch (error) {
      if (error.name === 'AbortError' || globalSearchController !== controller) return;
      console.error(error);
      results.replaceChildren();
      const message = document.createElement('p');
      message.className = 'global-search-message';
      message.setAttribute('role', 'alert');
      message.textContent = 'La recherche est momentanément indisponible. Réessayez.';
      results.append(message);
    }
  };

  document.addEventListener('submit', (event) => {
    if (!(event.target instanceof HTMLFormElement)) return;
    const form = event.target;
    if (form.id === 'global-search') {
      event.preventDefault();
      window.clearTimeout(globalSearchTimer);
      const input = form.querySelector('input[name="search"]');
      if (input) runGlobalSearch(input);
      return;
    }
    if (!form.matches('[data-live-filter], [data-live-search]')) return;
    event.preventDefault();
    const searchInput = form.querySelector('input[type="search"][name="search"]');
    const liveSearchSubmit = form.dataset.liveSearchSubmit === 'true';
    delete form.dataset.liveSearchSubmit;
    if (searchInput && (liveSearchSubmit || (!event.submitter && document.activeElement === searchInput))) {
      updatePageSearch(form);
    } else if (form.matches('[data-live-filter]')) {
      updatePageFilter(form);
    } else {
      updatePageSearch(form);
    }
  });

  document.addEventListener('input', (event) => {
    if (!(event.target instanceof HTMLInputElement)) return;
    const input = event.target;
    if (input.matches('#global-search input[name="search"]')) {
      globalSearchController?.abort();
      window.clearTimeout(globalSearchTimer);
      if (!input.value.trim()) {
        runGlobalSearch(input);
        return;
      }
      filterCurrentPageItems(input.value);
      const results = document.getElementById('global-search-results');
      const form = input.closest('#global-search');
      if (results) {
        results.replaceChildren();
        const loading = document.createElement('p');
        loading.className = 'global-search-message';
        loading.textContent = 'Recherche en cours…';
        results.append(loading);
        results.hidden = false;
        form?.setAttribute('aria-expanded', 'true');
      }
      globalSearchTimer = window.setTimeout(() => runGlobalSearch(input), 250);
      return;
    }
    const form = input.closest('form[data-live-search]');
    if (!form || input.name !== 'search') return;
    liveSearchControllers.get(form)?.abort();
    window.clearTimeout(liveSearchTimers.get(form));
    form.dataset.liveSearchSubmit = 'true';
    liveSearchTimers.set(form, window.setTimeout(() => {
      if (form.isConnected) form.requestSubmit();
    }, 300));
  });

  document.addEventListener('change', (event) => {
    if (!(event.target instanceof Element) || !event.target.matches('select, input[type="month"], input[type="date"], input[type="number"]')) return;
    const form = event.target.closest('form[data-live-filter]');
    if (form) form.requestSubmit();
  });

  document.addEventListener('keydown', (event) => {
    if (!(event.target instanceof HTMLInputElement) || !event.target.matches('#global-search input[name="search"]')) return;
    if (event.key === 'Escape') {
      const form = event.target.closest('#global-search');
      const results = document.getElementById('global-search-results');
      if (results) results.hidden = true;
      form?.setAttribute('aria-expanded', 'false');
    }
  });

  const requestedModalId = new URLSearchParams(window.location.search).get('open');
  const requestedModalTemplate = requestedModalId ? document.getElementById(requestedModalId) : null;
  if (requestedModalTemplate instanceof HTMLTemplateElement) {
    const trigger = [...document.querySelectorAll('[data-modal-template]')]
      .find((element) => element.dataset.modalTemplate === requestedModalId);
    const url = new URL(window.location.href);
    url.searchParams.delete('open');
    window.history.replaceState(window.history.state, '', `${url.pathname}${url.search}${url.hash}`);
    openModal(requestedModalTemplate.content.cloneNode(true), trigger || null);
  }
})();
