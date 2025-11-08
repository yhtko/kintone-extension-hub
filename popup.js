'use strict';

import { initFavorites } from './favorites.js';
import { initSchedule } from './schedule.js';
import { initPins } from './pins.js';
import { initShortcuts } from './shortcuts.js';

const favorites = initFavorites();
const schedule = initSchedule();
const pins = initPins();
const shortcuts = initShortcuts();

function setupTabs() {
  let currentTabName = 'favorites';
  const tabBar = document.querySelector('.tab-bar');
  if (!tabBar) return;

  const buttons = Array.from(tabBar.querySelectorAll('.tab-btn'));
  if (!buttons.length) return;

  const panels = new Map();
  buttons.forEach((btn) => {
    const key = btn.dataset.tab;
    if (!key) return;
    const panel = document.getElementById(`tab-${key}`);
    if (panel) panels.set(key, panel);
  });
  if (!panels.size) return;

  let current = null;

  function activate(name) {
    if (!panels.has(name) || current === name) return;
    current = name;
    currentTabName = name;
    buttons.forEach((btn) => {
      btn.classList.toggle('active', btn.dataset.tab === name);
    });
    panels.forEach((panel, key) => {
      panel.classList.toggle('active', key === name);
    });
    if (name === 'schedule') {
      schedule.refreshSchedule?.();
    } else if (name === 'favorites') {
      favorites.refreshCounts?.();
    } else if (name === 'pins') {
      pins.refreshPins?.();
    }
  }

  buttons.forEach((btn) => {
    btn.addEventListener('click', () => {
      if (btn.disabled) return;
      const key = btn.dataset.tab;
      if (key) activate(key);
    });
  });

  const initial = buttons.find((btn) => btn.classList.contains('active'))?.dataset.tab
    || buttons[0]?.dataset.tab
    || 'favorites';
  if (initial) activate(initial);

  return {
    getActiveTabName() {
      return currentTabName;
    }
  };
}

const tabState = setupTabs() || { getActiveTabName: () => 'favorites' };

const quickAddBtn = document.getElementById('quickAdd');
quickAddBtn?.addEventListener('click', async () => {
  const activeTab = tabState.getActiveTabName();
  if (activeTab === 'pins') {
    await pins.quickAdd?.();
  } else {
    await favorites.quickAdd?.();
  }
});

window.addEventListener('beforeunload', () => {
  favorites.dispose?.();
  schedule.dispose?.();
  pins.dispose?.();
  shortcuts.dispose?.();
});
