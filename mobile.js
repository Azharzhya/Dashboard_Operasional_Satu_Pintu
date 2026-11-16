/* =====================================================================
   mobile.js — pelengkap tampilan HP (mobile.css)
   Satu tugas: tab aktif di navigasi bawah selalu terlihat, karena
   navigasi bisa digeser bila tabnya lebih banyak dari lebar layar.
   Tidak menyentuh logika aplikasi (dashboard.js).
   ===================================================================== */
(() => {
  'use strict';
  const mq = window.matchMedia('(max-width:760px)');
  const reduce = window.matchMedia('(prefers-reduced-motion:reduce)');
  let raf = 0;

  function reveal() {
    cancelAnimationFrame(raf);
    raf = requestAnimationFrame(() => {
      if (!mq.matches) return;
      const nav = document.querySelector('#app>.tabs');
      const tab = nav && nav.querySelector('.tab[aria-selected="true"]');
      if (!tab || tab.hidden || !tab.offsetParent) return;
      tab.scrollIntoView({ inline: 'center', block: 'nearest', behavior: reduce.matches ? 'auto' : 'smooth' });
    });
  }

  function start() {
    const nav = document.querySelector('#app>.tabs');
    if (!nav) return;
    new MutationObserver(reveal).observe(nav, { subtree: true, attributes: true, attributeFilter: ['aria-selected', 'hidden'] });
    mq.addEventListener('change', reveal);
    window.addEventListener('resize', reveal, { passive: true });
    reveal();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
