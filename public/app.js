(() => {
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
    openModal(error);
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
      if (!response.ok) throw new Error(`L’action a échoué (${response.status}).`);

      if (response.redirected && new URL(response.url).pathname !== window.location.pathname) {
        const html = await response.text();
        const doc = new DOMParser().parseFromString(html, 'text/html');
        const content = doc.querySelector('.page-content') || doc.querySelector('main');
        if (!content) throw new Error('La page demandée est indisponible.');
        openModal(content.cloneNode(true));
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
    submitModalForm(event.target);
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
