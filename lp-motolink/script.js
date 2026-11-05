'use strict';
const toggle = document.querySelector('.menu-toggle');
const navigation = document.querySelector('#navigation');
const mobile = window.matchMedia('(max-width: 760px)');
function setOpen(open) {
  navigation.hidden = !open;
  toggle.setAttribute('aria-expanded', String(open));
  toggle.textContent = open ? '閉じる' : 'メニュー';
}
function syncNavigation() {
  toggle.hidden = !mobile.matches;
  setOpen(!mobile.matches);
}
toggle.addEventListener('click', () => setOpen(toggle.getAttribute('aria-expanded') !== 'true'));
navigation.addEventListener('click', event => {
  if (mobile.matches && event.target.closest('a')) setOpen(false);
});
document.addEventListener('keydown', event => {
  if (event.key === 'Escape' && mobile.matches && !navigation.hidden) {
    setOpen(false);
    toggle.focus();
  }
});
mobile.addEventListener('change', syncNavigation);
syncNavigation();
