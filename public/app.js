(() => {
  const loginForm = document.querySelector('.login-form');
  loginForm?.addEventListener('submit', () => {
    const button = loginForm.querySelector('.login-submit');
    const label = loginForm.querySelector('[data-login-label]');
    if (!button || !label) return;
    button.disabled = true;
    button.classList.add('is-loading');
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

  window.addEventListener('pageshow', (event) => {
    if (event.persisted && document.body.dataset.privatePage === 'true') window.location.reload();
  });

  document.querySelectorAll('.transactions-table').forEach((table) => {
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

  const modal = document.getElementById('app-modal');
  const modalContent = document.getElementById('modal-content');
  if (!modal || !modalContent) return;

  const openModal = (content) => {
    modalContent.replaceChildren(content);
    modal.classList.add('open');
    modal.setAttribute('aria-hidden', 'false');
    document.body.classList.add('modal-open');
    modal.querySelector('.close-modal')?.focus();
  };

  const closeModal = () => {
    modal.classList.remove('open');
    modal.setAttribute('aria-hidden', 'true');
    document.body.classList.remove('modal-open');
    modalContent.replaceChildren();
  };

  const showError = (message) => {
    const error = document.createElement('div');
    error.className = 'modal-error';
    error.setAttribute('role', 'alert');
    error.textContent = message;
    modalContent.querySelector('.modal-error')?.remove();
    modalContent.prepend(error);
    modal.classList.add('open');
    modal.setAttribute('aria-hidden', 'false');
    document.body.classList.add('modal-open');
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
    heading.innerHTML = '<span class="eyebrow">CONFIRMATION</span>';
    const title = document.createElement('h2');
    title.textContent = trigger.dataset.modalTitle || 'Confirmer cette action';
    const message = document.createElement('p');
    message.textContent = trigger.dataset.modalMessage || 'Voulez-vous continuer ?';
    heading.append(title, message);

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
    openModal(wrapper);
  };

  const submitModalForm = async (form) => {
    const method = form.method.toUpperCase();
    const formData = new FormData(form);
    const isMultipart = form.enctype === 'multipart/form-data';
    const body = method === 'GET'
      ? undefined
      : (isMultipart ? formData : new URLSearchParams(formData));
    try {
      const response = await fetch(form.action || window.location.href, {
        method,
        body,
        credentials: 'same-origin',
      });
      if (!response.ok) {
        const responseText = await response.text();
        const errorDocument = new DOMParser().parseFromString(responseText, 'text/html');
        const message = (errorDocument.body.textContent || '').trim();
        throw new Error(message || `L’action a échoué (${response.status}).`);
      }

      if (response.redirected && new URL(response.url).pathname !== window.location.pathname) {
        const html = await response.text();
        const doc = new DOMParser().parseFromString(html, 'text/html');
        const content = doc.querySelector('.page-content') || doc.querySelector('main');
        if (!content) throw new Error('La page demandée est indisponible.');
        openModal(content.cloneNode(true));
        return;
      }

      if (response.redirected && new URL(response.url).href !== window.location.href) {
        window.location.assign(response.url);
        return;
      }

      closeModal();
      window.location.reload();
    } catch (error) {
      showError(error.message);
    }
  };

  document.addEventListener('click', (event) => {
    const trigger = event.target.closest('[data-modal-template], [data-modal-confirm], [data-close-modal]');
    if (!trigger) return;

    if (trigger.hasAttribute('data-close-modal')) {
      closeModal();
      return;
    }
    if (trigger.dataset.modalConfirm) {
      confirmModal(trigger);
      return;
    }

    const template = document.getElementById(trigger.dataset.modalTemplate);
    if (template instanceof HTMLTemplateElement) {
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
      openModal(content);
    }
  });

  modal.addEventListener('click', (event) => {
    if (event.target === modal.querySelector('.modal-backdrop')) closeModal();
  });

  modalContent.addEventListener('submit', (event) => {
    if (!(event.target instanceof HTMLFormElement)) return;
    event.preventDefault();
    prepareIncomeSubmission(event.target);
    submitModalForm(event.target);
  });

  modalContent.addEventListener('change', (event) => {
    if (event.target.matches('[data-income-type]')) {
      const form = event.target.closest('[data-income-entry-form]');
      if (form) configureIncomeEntry(form);
    }
  });

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && modal.classList.contains('open')) closeModal();
  });

  const search = document.getElementById('dashboard-search');
  search?.addEventListener('submit', (event) => {
    event.preventDefault();
    const query = new FormData(search).get('search').toString().trim();
    const url = new URL('/expenses', window.location.origin);
    if (query) url.searchParams.set('search', query);
    window.location.assign(url);
  });

  document.getElementById('month-select')?.addEventListener('change', (event) => {
    event.currentTarget.form.requestSubmit();
  });
})();
