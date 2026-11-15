(() => {
  if (localStorage.getItem('maison-theme') === 'dark') {
    document.documentElement.dataset.theme = 'dark';
  }
})();
