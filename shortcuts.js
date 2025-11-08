"use strict";

import {
  loadShortcuts,
  sortShortcuts,
  loadShortcutVisibility,
  saveShortcutVisibility
} from './core.js';

const SHORTCUT_TYPE_LABEL = {
  appTop: 'アプリトップ',
  view: 'ビュー',
  create: 'レコード新規'
};
const MAX_SHORTCUT_INITIAL_LENGTH = 2;

function normalizeCustomInitial(value) {
  if (value == null) return '';
  const str = String(value).trim();
  if (!str) return '';
  return Array.from(str).slice(0, MAX_SHORTCUT_INITIAL_LENGTH).join('');
}

function buildShortcutUrl(entry) {
  const host = (entry.host || '').replace(/\/$/, '');
  const appId = (entry.appId || '').trim();
  if (!host || !appId) return '';
  if (entry.type === 'create') {
    return `${host}/k/${encodeURIComponent(appId)}/edit`;
  }
  const base = `${host}/k/${encodeURIComponent(appId)}/`;
  if (entry.type === 'view') {
    const view = (entry.viewIdOrName || '').trim();
    if (view) {
      return `${base}?view=${encodeURIComponent(view)}`;
    }
  }
  return base;
}

function extractInitial(text) {
  const trimmed = text.trim();
  if (!trimmed) return '';
  const [first] = Array.from(trimmed);
  if (!first) return '';
  const upperCandidate = first.normalize('NFKC');
  if (/^[a-z]$/i.test(upperCandidate)) {
    return upperCandidate.toUpperCase();
  }
  return first;
}

function initialForEntry(entry) {
  const custom = normalizeCustomInitial(entry.initial);
  if (custom) return custom;
  const label = typeof entry.label === 'string' ? entry.label : '';
  const fallback = label.trim()
    || String(entry.appId || '').trim()
    || (entry.host || '').replace(/^https?:\/\//, '').trim();
  return extractInitial(fallback) || 'A';
}

export function initShortcuts(options = {}) {
  const doc = options.document || document;
  const rail = doc.getElementById('shortcutRail');
  const toggleBtn = doc.getElementById('shortcutRailToggle');
  const buttonWrap = doc.getElementById('shortcutButtons');

  if (!rail || !toggleBtn || !buttonWrap) {
    return {
      async refresh() {},
      dispose() {}
    };
  }

  const state = {
    entries: [],
    visible: true,
    storageListener: null,
    warmedHosts: new Set(),
    toastTimer: null,
    toastEl: null
  };

  function applyVisibility(visible) {
    state.visible = visible;
    rail.classList.toggle('collapsed', !visible);
    toggleBtn.textContent = visible ? '×' : '?';
    toggleBtn.setAttribute('aria-label', visible ? 'ショートカットを隠す' : 'ショートカットを表示');
  }

  function clearButtons() {
    while (buttonWrap.firstChild) {
      buttonWrap.removeChild(buttonWrap.firstChild);
    }
  }

  function hideToast(immediate = false) {
    if (state.toastTimer) {
      clearTimeout(state.toastTimer);
      state.toastTimer = null;
    }
    if (state.toastEl) {
      if (immediate) {
        state.toastEl.classList.remove('show');
      } else {
        state.toastEl.classList.remove('show');
      }
    }
  }

  function showToast(message, duration = 4200) {
    if (!doc.body) return;
    if (!state.toastEl) {
      const el = doc.createElement('div');
      el.className = 'shortcut-toast';
      el.setAttribute('role', 'status');
      el.setAttribute('aria-live', 'polite');
      el.addEventListener('click', () => hideToast(true));
      doc.body.appendChild(el);
      state.toastEl = el;
    }
    state.toastEl.textContent = message;
    state.toastEl.classList.add('show');
    if (state.toastTimer) {
      clearTimeout(state.toastTimer);
    }
    state.toastTimer = setTimeout(() => {
      hideToast();
    }, duration);
  }

  function renderButtons() {
    clearButtons();
    if (!state.entries.length) {
      const emptyBtn = doc.createElement('button');
      emptyBtn.type = 'button';
      emptyBtn.className = 'shortcut-button shortcut-button-empty';
      emptyBtn.textContent = '+';
      emptyBtn.title = 'ショートカットはまだありません。設定画面で追加できます。';
      emptyBtn.addEventListener('click', () => {
        if (chrome.runtime.openOptionsPage) {
          chrome.runtime.openOptionsPage();
        }
      });
      buttonWrap.appendChild(emptyBtn);
      return;
    }

    state.entries.forEach((entry) => {
      const btn = doc.createElement('button');
      btn.type = 'button';
      const type = entry.type || 'appTop';
      btn.className = `shortcut-button shortcut-${type}`;
      btn.dataset.type = type;
      const label = entry.label || SHORTCUT_TYPE_LABEL[entry.type] || SHORTCUT_TYPE_LABEL.appTop;
      btn.title = label;
      btn.setAttribute('aria-label', label);
      const targetUrl = buildShortcutUrl(entry);
      const initialChars = Array.from(initialForEntry(entry));
      const initial = initialChars.length
        ? initialChars.slice(0, MAX_SHORTCUT_INITIAL_LENGTH).join('')
        : '?';
      if (!targetUrl) {
        btn.disabled = true;
        btn.classList.add('disabled');
      }
      btn.textContent = initial;
      btn.addEventListener('click', async () => {
        if (!targetUrl) return;
        try {
          if (entry.host) {
            await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host: entry.host });
          }
        } catch (_err) {}
        let opened = false;
        try {
          await chrome.tabs.create({ url: targetUrl, active: true });
          opened = true;
        } catch (_err) {
          // noop
        }
        if (entry.type === 'create') {
          if (opened) {
            showToast('作成をキャンセルする場合はタブを閉じてください');
          } else {
            showToast('レコード作成画面を開けませんでした');
          }
        }
      });
      buttonWrap.appendChild(btn);
    });
  }

  async function warmHosts(entries) {
    for (const entry of entries) {
      const host = entry.host;
      if (!host || state.warmedHosts.has(host)) continue;
      state.warmedHosts.add(host);
      try {
        await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host });
      } catch (_err) {}
    }
  }

  async function refreshEntries() {
    try {
      const [list, visible] = await Promise.all([
        loadShortcuts(),
        loadShortcutVisibility()
      ]);
      state.entries = sortShortcuts(list);
      applyVisibility(visible);
      renderButtons();
      await warmHosts(state.entries);
    } catch (_err) {
      // noop
    }
  }

  async function handleToggleClick() {
    const next = !state.visible;
    applyVisibility(next);
    try {
      await saveShortcutVisibility(next);
    } catch (_err) {}
  }

  const onToggleClick = () => {
    handleToggleClick().catch(() => {});
  };
  toggleBtn.addEventListener('click', onToggleClick);

  if (chrome?.storage?.onChanged) {
    state.storageListener = (changes, area) => {
      if (area !== 'sync') return;
      if (Object.prototype.hasOwnProperty.call(changes, 'kfavShortcuts')) {
        refreshEntries().catch(() => {});
      }
      if (Object.prototype.hasOwnProperty.call(changes, 'kfavShortcutsVisible')) {
        const flag = changes.kfavShortcutsVisible.newValue;
        applyVisibility(typeof flag === 'boolean' ? flag : true);
      }
    };
    chrome.storage.onChanged.addListener(state.storageListener);
  }

  refreshEntries();

  return {
    async refresh() {
      await refreshEntries();
    },
    dispose() {
      toggleBtn.removeEventListener('click', onToggleClick);
      if (state.storageListener) {
        chrome.storage.onChanged.removeListener(state.storageListener);
      }
      hideToast(true);
      if (state.toastEl) {
        state.toastEl.remove();
        state.toastEl = null;
      }
    }
  };
}




