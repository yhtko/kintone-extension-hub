'use strict';

import {
  loadFavorites,
  saveFavorites,
  sortFavorites,
  loadShortcuts,
  sortShortcuts,
  loadShortcutVisibility,
  saveShortcutVisibility,
  getActiveTab,
  parseKintoneUrl,
  isKintoneUrl,
  hasHostPermission,
  sendCountBulk,
  createId
} from './core.js';

const REFRESH_INTERVAL = 30000;
const MAX_SHORTCUT_INITIAL_LENGTH = 2;

const doc = document;

const els = {
  filter: doc.getElementById('filter'),
  clearFilter: doc.getElementById('clearFilter'),
  togglePinsOnly: doc.getElementById('togglePinsOnly'),
  pinList: doc.getElementById('pinList'),
  pinSection: doc.getElementById('pinSection'),
  favSection: doc.getElementById('favSection'),
  favCategories: doc.getElementById('favCategories'),
  notice: doc.getElementById('notice'),
  refreshCounts: doc.getElementById('refreshCounts'),
  quickAdd: doc.getElementById('quickAdd'),
  openOptions: doc.getElementById('openOptions'),
  menuLayer: doc.getElementById('menuLayer'),
  shortcutRow: doc.getElementById('shortcutRow'),
  shortcutPanel: doc.getElementById('shortcutPanel'),
  toggleShortcuts: doc.getElementById('toggleShortcuts')
};

const state = {
  favorites: [],
  filterText: '',
  pinsOnly: false,
  badgeStatus: new Map(),
  selectionIds: [],
  selectionIndex: -1,
  countTimer: null,
  storageListener: null,
  menu: null,
  dragging: null,
  shortcutToggleHandler: null
};

const boundCloseMenuOnBlur = () => closeMenu();
const shortcutState = {
  entries: [],
  visible: true,
  warmedHosts: new Set()
};
const DEFAULT_ICON = 'bookmark';
const DEFAULT_CATEGORY = 'その他';

const ICON_PATHS = {
  clipboard: '<path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><rect x="9" y="2" width="6" height="4" rx="1" ry="1"/>',
  'file-text': '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M10 13h6"/><path d="M10 17h4"/>',
  package: '<path d="M21 16V8a2 2 0 0 0-1-1.73L13 3.27a2 2 0 0 0-2 0L4 6.27A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><path d="M3.3 7.5 12 12.5l8.7-5"/><path d="M12 22v-9"/>',
  box: '<path d="M21 16V8l-9-5-9 5v8l9 5 9-5z"/><path d="M3 8l9 5 9-5"/><path d="M12 13v9"/>',
  truck: '<path d="M3 5h12v10H3z"/><path d="M15 8h4l2 3v4h-6"/><circle cx="7.5" cy="17.5" r="2"/><circle cx="17.5" cy="17.5" r="2"/>',
  factory: '<path d="M2 20h20V10l-5 3V8l-5 3V4L2 8z"/><path d="M6 20v-4"/><path d="M10 20v-4"/><path d="M14 20v-4"/>',
  wrench: '<path d="M21 16.5a5.5 5.5 0 0 1-7.78 0L3.5 6.78a4 4 0 0 1 5.66-5.66l.82.82-3 3 4.24 4.24 3-3 .82.82a4 4 0 0 1 0 5.66z"/>',
  calendar: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 8h18"/><path d="M16 2v4"/><path d="M8 2v4"/>',
  'list-checks': '<path d="M10 6h11"/><path d="M10 12h11"/><path d="M10 18h11"/><path d="m3 6 1.5 1.5 3-3"/><path d="m3 12 1.5 1.5 3-3"/><path d="m3 18 1.5 1.5 3-3"/>',
  search: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/>',
  'chart-bar': '<path d="M3 3v18h18"/><rect x="7" y="13" width="3" height="5" rx="0.5"/><rect x="12" y="9" width="3" height="9" rx="0.5"/><rect x="17" y="5" width="3" height="13" rx="0.5"/>',
  receipt: '<path d="M21 4v16l-3-2-3 2-3-2-3 2-3-2-3 2V4l3 2 3-2 3 2 3-2 3 2z"/><path d="M8 11h8"/><path d="M8 15h5"/>',
  users: '<path d="M16 21v-1a5 5 0 0 0-3-4.58"/><path d="M7 4a4 4 0 1 1 1 7.87A4 4 0 0 1 7 4z"/><path d="M3 21v-1a7 7 0 0 1 11.9-4.9"/><path d="M20 8a3 3 0 1 1-4 2.82"/>',
  settings: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 2l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-2-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09a1.65 1.65 0 0 0-1-1.51 1.65 1.65 0 0 0-2 .33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-2 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09a1.65 1.65 0 0 0 1.51-1 1.65 1.65 0 0 0-.33-2l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 2 .33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 2-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 2V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/>',
  bookmark: '<path d="M6 3h12v18l-6-4-6 4z"/>',
  star: '<path d="M12 3.5 14.8 9l6.2.9-4.5 4.3 1 6.3L12 17l-5.5 3.5 1-6.3-4.5-4.3L9.2 9z"/>'
};
const ICON_OPTIONS = Object.keys(ICON_PATHS);

function renderIconSvg(name) {
  const body = ICON_PATHS[name] || ICON_PATHS[DEFAULT_ICON];
  return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${body}</svg>`;
}

function escapeId(value) {
  if (value == null) return '';
  if (globalThis.CSS?.escape) return CSS.escape(String(value));
  return String(value).replace(/[^a-zA-Z0-9_-]/g, (ch) => `\\${ch}`);
}

function normalizeShortcutInitial(value) {
  if (value == null) return '';
  const str = String(value).trim();
  if (!str) return '';
  return Array.from(str).slice(0, MAX_SHORTCUT_INITIAL_LENGTH).join('');
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

function shortcutInitial(entry) {
  const custom = normalizeShortcutInitial(entry.initial);
  if (custom) return custom;
  const label = typeof entry.label === 'string' ? entry.label : '';
  const fallback = label.trim()
    || String(entry.appId || '').trim()
    || (entry.host || '').replace(/^https?:\/\//, '').trim();
  return extractInitial(fallback) || 'A';
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

function setNotice(message) {
  if (!els.notice) return;
  if (message) {
    els.notice.textContent = message;
    els.notice.classList.remove('hidden');
  } else {
    els.notice.textContent = '';
    els.notice.classList.add('hidden');
  }
}

function normalizeFavorites(list) {
  return sortFavorites(
    (list || []).map((item, index) => ({
      id: item.id || createId(),
      label: item.label || '(no label)',
      url: item.url || '',
      host: item.host || '',
      appId: item.appId || '',
      viewIdOrName: item.viewIdOrName || '',
      query: item.query || '',
      pinned: Boolean(item.pinned),
      order: typeof item.order === 'number' ? item.order : index,
      icon: typeof item.icon === 'string' && item.icon ? item.icon : DEFAULT_ICON,
      category: typeof item.category === 'string' && item.category.trim() ? item.category.trim() : DEFAULT_CATEGORY
    }))
  );
}

function applyFilterText(value) {
  state.filterText = value || '';
  renderLists();
}

function currentFilteredLists() {
  const keyword = state.filterText.trim().toLowerCase();
  const matches = (item) => {
    if (!keyword) return true;
    const fields = [item.label, item.url, item.host, item.viewIdOrName, item.query];
    return fields.some((field) => typeof field === 'string' && field.toLowerCase().includes(keyword));
  };
  const pinned = [];
  const normal = [];
  state.favorites.forEach((item) => {
    if (!matches(item)) return;
    if (item.pinned) pinned.push(item);
    else normal.push(item);
  });
  pinned.sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  normal.sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  return {
    pinned,
    normal: state.pinsOnly ? [] : normal
  };
}

function renderLists() {
  const { pinned, normal } = currentFilteredLists();
  const pinnedOrder = renderPinnedList(pinned);
  const normalOrder = renderCategorySections(normal);
  if (els.pinSection) {
    els.pinSection.classList.toggle('hidden', pinned.length === 0);
  }
  if (els.favSection) {
    els.favSection.classList.remove('hidden');
  }
  const combined = pinnedOrder.concat(normalOrder);
  state.selectionIds = combined.map((item) => item.id);
  if (combined.length === 0) {
    state.selectionIndex = -1;
  } else if (state.selectionIndex < 0 || state.selectionIndex >= combined.length) {
    state.selectionIndex = 0;
  }
  highlightSelection();
}

function renderPinnedList(items) {
  if (!els.pinList) return [];
  els.pinList.textContent = '';
  if (!items.length) {
    const empty = doc.createElement('li');
    empty.className = 'empty-message';
    const hasFilter = state.filterText.trim().length > 0;
    empty.textContent = hasFilter
      ? '検索条件に一致するピンはありません'
      : 'ピンはまだありません';
    els.pinList.appendChild(empty);
    return [];
  }
  items.forEach((item) => {
    els.pinList.appendChild(createEntryElement(item, true));
  });
  return items.slice();
}

function groupByCategory(items) {
  const map = new Map();
  items.forEach((item) => {
    const key = item.category || DEFAULT_CATEGORY;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(item);
  });
  return map;
}

function renderCategorySections(items) {
  const container = els.favCategories;
  if (!container) return [];
  container.textContent = '';
  if (!items.length) {
    const message = doc.createElement('div');
    message.className = 'empty-message';
    const hasFilter = state.filterText.trim().length > 0;
    message.textContent = hasFilter
      ? '検索条件に一致するウォッチリストはありません'
      : 'ウォッチリストはまだ登録されていません';
    container.appendChild(message);
    return [];
  }
  const grouped = groupByCategory(items);
  const order = Array.from(grouped.keys()).sort((a, b) => {
    if (a === DEFAULT_CATEGORY) return 1;
    if (b === DEFAULT_CATEGORY) return -1;
    return a.localeCompare(b, 'ja');
  });
  const flattened = [];
  order.forEach((name) => {
    const block = doc.createElement('section');
    block.className = 'category-block';
    const title = doc.createElement('h3');
    title.className = 'category-title';
    title.textContent = name;
    const list = doc.createElement('ul');
    list.className = 'category-list';
    grouped.get(name).forEach((entry) => {
      list.appendChild(createEntryElement(entry, false));
      flattened.push(entry);
    });
    block.appendChild(title);
    block.appendChild(list);
    container.appendChild(block);
  });
  return flattened;
}

function createEntryElement(entry, pinned) {
  const li = doc.createElement('li');
  li.className = 'entry';
  li.dataset.id = entry.id;
  li.dataset.pinned = String(pinned);
  li.dataset.category = entry.category || DEFAULT_CATEGORY;
  li.draggable = true;
  li.addEventListener('dragstart', handleDragStart);
  li.addEventListener('dragover', handleDragOver);
  li.addEventListener('dragleave', handleDragLeave);
  li.addEventListener('drop', handleDrop);
  li.addEventListener('dragend', handleDragEnd);

  const button = doc.createElement('button');
  button.type = 'button';
  button.className = 'entry-main';
  button.dataset.id = entry.id;
  button.addEventListener('click', () => openEntry(entry));
  button.addEventListener('focus', () => setSelectionById(entry.id));
  button.addEventListener('mouseenter', () => setSelectionById(entry.id));

  const left = doc.createElement('div');
  left.className = 'entry-left';

  const icon = doc.createElement('span');
  icon.className = 'entry-icon';
  icon.innerHTML = renderIconSvg(entry.icon || DEFAULT_ICON);

  const text = doc.createElement('div');
  text.className = 'entry-text';
  const title = doc.createElement('div');
  title.className = 'entry-title';
  title.textContent = entry.label || '(no label)';
  const sub = doc.createElement('div');
  sub.className = 'entry-sub';
  let host = '';
  try {
    host = entry.host ? new URL(entry.host).hostname : new URL(entry.url).hostname;
  } catch (_err) {
    host = entry.host || '';
  }
  const view = entry.viewIdOrName ? `view:${entry.viewIdOrName}` : '';
  sub.textContent = [host, view].filter(Boolean).join(' · ') || entry.url;
  text.appendChild(title);
  text.appendChild(sub);
  left.appendChild(icon);
  left.appendChild(text);

  const badge = doc.createElement('div');
  badge.className = 'entry-badge';
  badge.dataset.id = entry.id;
  syncBadgeToElement(entry.id, badge);

  const right = doc.createElement('div');
  right.className = 'entry-right';
  right.appendChild(badge);

  button.appendChild(left);
  button.appendChild(right);

  const menuBtn = doc.createElement('button');
  menuBtn.type = 'button';
  menuBtn.className = 'entry-menu-btn';
  menuBtn.textContent = '⋯';
  menuBtn.title = 'メニュー';
  menuBtn.addEventListener('click', (event) => {
    event.stopPropagation();
    openMenu(entry, menuBtn);
  });

  li.appendChild(button);
  li.appendChild(menuBtn);
  return li;
}

function highlightSelection() {
  const buttons = doc.querySelectorAll('.entry-main');
  buttons.forEach((btn) => btn.classList.remove('active'));
  if (state.selectionIndex < 0 || state.selectionIndex >= state.selectionIds.length) return;
  const id = state.selectionIds[state.selectionIndex];
  const active = doc.querySelector(`.entry-main[data-id="${escapeId(id)}"]`);
  if (active) {
    active.classList.add('active');
  }
}

function moveSelection(delta) {
  if (!state.selectionIds.length) return;
  const next = (state.selectionIndex + delta + state.selectionIds.length) % state.selectionIds.length;
  state.selectionIndex = next;
  highlightSelection();
  scrollSelectionIntoView();
}

function scrollSelectionIntoView() {
  if (state.selectionIndex < 0 || state.selectionIndex >= state.selectionIds.length) return;
  const id = state.selectionIds[state.selectionIndex];
  const active = doc.querySelector(`.entry-main[data-id="${escapeId(id)}"]`);
  if (active) {
    active.scrollIntoView({ block: 'nearest' });
  }
}

function setSelectionById(id) {
  const index = state.selectionIds.indexOf(id);
  if (index >= 0) {
    state.selectionIndex = index;
    highlightSelection();
  }
}

function openSelection() {
  if (state.selectionIndex < 0 || state.selectionIndex >= state.selectionIds.length) return;
  const id = state.selectionIds[state.selectionIndex];
  const entry = state.favorites.find((item) => item.id === id);
  if (entry) openEntry(entry);
}

async function openEntry(entry) {
  if (!entry?.url) return;
  closeMenu();
  try {
    if (chrome?.tabs?.create) {
      await chrome.tabs.create({ url: entry.url, active: true });
    } else {
      window.open(entry.url, '_blank', 'noopener');
    }
  } catch (error) {
    console.error('Failed to open entry', error);
    window.open(entry.url, '_blank', 'noopener');
  }
}

function syncBadgeToElement(id, el) {
  const stateItem = state.badgeStatus.get(id) || { text: '-', title: '未取得', loading: false };
  el.textContent = stateItem.text ?? '-';
  el.title = stateItem.title || '';
  el.classList.toggle('loading', Boolean(stateItem.loading));
}

function updateBadgeDom(id) {
  const el = doc.querySelector(`.entry-badge[data-id="${escapeId(id)}"]`);
  if (el) syncBadgeToElement(id, el);
}

function setBadgeLoading(id, loading) {
  const prev = state.badgeStatus.get(id) || { text: '-', title: '未取得', loading: false };
  state.badgeStatus.set(id, { ...prev, loading: Boolean(loading) });
  updateBadgeDom(id);
}

function setBadgeValue(id, value, title = '') {
  const text = value == null ? '—' : String(value);
  state.badgeStatus.set(id, { text, title: title || '', loading: false });
  updateBadgeDom(id);
}

async function loadFavoritesAndRender() {
  const list = await loadFavorites();
  state.favorites = normalizeFavorites(list);
  const validIds = new Set(state.favorites.map((item) => item.id));
  for (const key of Array.from(state.badgeStatus.keys())) {
    if (!validIds.has(key)) state.badgeStatus.delete(key);
  }
  renderLists();
  const hosts = new Set(state.favorites.map((item) => item.host).filter(Boolean));
  for (const host of hosts) {
    try {
      await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host });
    } catch (_err) {
      // ignore
    }
  }
}

function groupByHost(items) {
  const map = new Map();
  for (const item of items) {
    if (!item.host) continue;
    if (!map.has(item.host)) map.set(item.host, []);
    map.get(item.host).push(item);
  }
  return map;
}

async function refreshCounts() {
  if (!state.favorites.length) return;
  setNotice('');
  state.favorites.forEach((item) => setBadgeLoading(item.id, true));
  const grouped = groupByHost(state.favorites);
  const missingPerm = new Set();
  const failedHosts = new Set();
  for (const [host, items] of grouped.entries()) {
    try {
      const permitted = await hasHostPermission(host);
      if (!permitted) {
        items.forEach((item) => setBadgeValue(item.id, null, 'ホストの権限が必要です（設定）'));
        missingPerm.add(host);
        continue;
      }
      try {
        await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host });
      } catch (_warmErr) {
        // ignore
      }
      const payloadItems = items.map((item) => ({
        id: item.id,
        appId: item.appId,
        viewIdOrName: item.viewIdOrName || '',
        query: item.query || ''
      }));
      const res = await sendCountBulk(host, payloadItems);
      if (res?.ok) {
        const counts = res.counts || {};
        const errors = res.errors || {};
        for (const item of items) {
          if (errors[item.id]) {
            setBadgeValue(item.id, null, errors[item.id]);
            failedHosts.add(host);
          } else if (Object.prototype.hasOwnProperty.call(counts, item.id)) {
            setBadgeValue(item.id, counts[item.id]);
          } else {
            setBadgeValue(item.id, null, '件数を取得できませんでした');
            failedHosts.add(host);
          }
        }
      } else {
        const message = res?.error || '件数取得に失敗しました';
        items.forEach((item) => setBadgeValue(item.id, null, message));
        failedHosts.add(host);
      }
    } catch (error) {
      const message = String(error?.message || error);
      items.forEach((item) => setBadgeValue(item.id, null, message));
      failedHosts.add(host);
    }
  }
  if (missingPerm.size) {
    setNotice(`⚠️ 次のドメインの権限が必要です: ${Array.from(missingPerm).join(', ')}`);
  } else if (failedHosts.size) {
    setNotice('⚠️ 一部の件数取得に失敗しました。ログイン状態を確認してください。');
  }
}

function attachStorageListener() {
  if (!chrome?.storage?.onChanged) return;
  const listener = (changes, area) => {
    if (area !== 'sync') return;
    if (changes.kintoneFavorites) {
      loadFavoritesAndRender().catch((error) => {
        console.error('Failed to reload favorites', error);
      });
    }
    if (Object.prototype.hasOwnProperty.call(changes, 'kfavShortcuts')) {
      refreshShortcutEntries().catch((error) => {
        console.error('Failed to reload shortcuts', error);
      });
    }
    if (Object.prototype.hasOwnProperty.call(changes, 'kfavShortcutsVisible')) {
      const next = typeof changes.kfavShortcutsVisible.newValue === 'boolean'
        ? changes.kfavShortcutsVisible.newValue
        : true;
      applyShortcutVisibility(next);
    }
  };
  chrome.storage.onChanged.addListener(listener);
  state.storageListener = listener;
}

function detachStorageListener() {
  if (state.storageListener && chrome?.storage?.onChanged) {
    chrome.storage.onChanged.removeListener(state.storageListener);
    state.storageListener = null;
  }
}

function startCountTimer() {
  stopCountTimer();
  state.countTimer = setInterval(() => {
    refreshCounts().catch((error) => console.error('refreshCounts failed', error));
  }, REFRESH_INTERVAL);
}

function stopCountTimer() {
  if (state.countTimer) {
    clearInterval(state.countTimer);
    state.countTimer = null;
  }
}

function closeMenu() {
  if (!els.menuLayer) return;
  els.menuLayer.classList.add('hidden');
  els.menuLayer.classList.remove('active');
  els.menuLayer.textContent = '';
  state.menu = null;
}

function openMenu(entry, anchor) {
  if (!els.menuLayer) return;
  closeMenu();
  const menu = doc.createElement('div');
  menu.className = 'context-menu';
  const actions = [
    { key: 'open', label: '開く', handler: () => openEntry(entry) },
    { key: 'edit', label: '編集', handler: () => editEntry(entry) },
    { key: 'icon', label: 'アイコンを変更', handler: () => changeEntryIcon(entry) },
    { key: 'category', label: 'カテゴリを変更', handler: () => changeEntryCategory(entry) },
    { key: 'pin', label: entry.pinned ? 'ピン解除' : '📌 ピン留め', handler: () => togglePin(entry.id, !entry.pinned) },
    { key: 'delete', label: '削除', handler: () => deleteEntry(entry.id), danger: true }
  ];
  actions.forEach((action) => {
    const btn = doc.createElement('button');
    btn.type = 'button';
    btn.textContent = action.label;
    if (action.danger) btn.classList.add('danger');
    btn.addEventListener('click', () => {
      closeMenu();
      action.handler();
    });
    menu.appendChild(btn);
  });
  els.menuLayer.appendChild(menu);
  els.menuLayer.classList.remove('hidden');
  els.menuLayer.classList.add('active');
  state.menu = { entryId: entry.id, menu };
  positionMenu(menu, anchor);
}

function positionMenu(menu, anchor) {
  const anchorRect = anchor.getBoundingClientRect();
  const menuRect = menu.getBoundingClientRect();
  let left = anchorRect.right - menuRect.width;
  if (left < 12) left = 12;
  let top = anchorRect.bottom + 6;
  if (top + menuRect.height > window.innerHeight - 12) {
    top = anchorRect.top - menuRect.height - 6;
  }
  if (top < 12) top = 12;
  menu.style.left = `${left}px`;
  menu.style.top = `${top}px`;
}

function applyShortcutVisibility(visible) {
  shortcutState.visible = visible;
  if (els.shortcutPanel) {
    els.shortcutPanel.classList.toggle('collapsed', !visible);
  }
  if (els.toggleShortcuts) {
    els.toggleShortcuts.setAttribute('aria-pressed', visible ? 'true' : 'false');
    els.toggleShortcuts.textContent = visible ? '━' : '＋';
    els.toggleShortcuts.title = visible ? 'ショートカットを隠す' : 'ショートカットを表示';
  }
}

function renderShortcuts() {
  if (!els.shortcutRow) return;
  els.shortcutRow.textContent = '';
  const entries = shortcutState.entries;
  if (!entries.length) {
    const empty = doc.createElement('button');
    empty.type = 'button';
    empty.className = 'shortcut-button shortcut-empty';
    empty.textContent = '+';
    empty.title = 'ショートカットを設定';
    empty.addEventListener('click', () => {
      if (chrome?.runtime?.openOptionsPage) {
        chrome.runtime.openOptionsPage();
      } else {
        window.open(chrome.runtime.getURL('options.html'), '_blank');
      }
    });
    els.shortcutRow.appendChild(empty);
    return;
  }
  entries.forEach((entry) => {
    const btn = doc.createElement('button');
    btn.type = 'button';
    const type = entry.type || 'appTop';
    btn.className = `shortcut-button shortcut-${type}`;
    btn.title = entry.label || '';
    btn.setAttribute('aria-label', entry.label || 'ショートカット');
    const url = buildShortcutUrl(entry);
    if (!url) {
      btn.disabled = true;
      btn.classList.add('disabled');
    }
    btn.textContent = Array.from(shortcutInitial(entry)).slice(0, MAX_SHORTCUT_INITIAL_LENGTH).join('') || '?';
    btn.addEventListener('click', () => openShortcut(entry));
    els.shortcutRow.appendChild(btn);
  });
}

async function openShortcut(entry) {
  const url = buildShortcutUrl(entry);
  if (!url) {
    setNotice('ショートカットのURLが設定されていません');
    return;
  }
  closeMenu();
  if (entry.host) {
    try {
      await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host: entry.host });
    } catch (_err) {}
  }
  try {
    await chrome.tabs.create({ url, active: true });
  } catch (_err) {
    window.open(url, '_blank', 'noopener');
  }
}

async function warmShortcutHosts(entries) {
  for (const entry of entries) {
    const host = entry.host;
    if (!host || shortcutState.warmedHosts.has(host)) continue;
    shortcutState.warmedHosts.add(host);
    try {
      await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host });
    } catch (_err) {}
  }
}

async function refreshShortcutEntries() {
  try {
    const [list, visible] = await Promise.all([
      loadShortcuts(),
      loadShortcutVisibility()
    ]);
    shortcutState.entries = sortShortcuts(list);
    applyShortcutVisibility(typeof visible === 'boolean' ? visible : true);
    renderShortcuts();
    await warmShortcutHosts(shortcutState.entries);
  } catch (error) {
    console.error('Failed to refresh shortcuts', error);
  }
}

async function toggleShortcutVisibility() {
  const next = !shortcutState.visible;
  applyShortcutVisibility(next);
  try {
    await saveShortcutVisibility(next);
  } catch (_err) {}
}

async function quickAddFromActiveTab() {
  try {
    const tab = await getActiveTab();
    if (!tab || !isKintoneUrl(tab.url || '')) {
      setNotice('kintoneのタブで実行してください');
      return;
    }
    const parsed = parseKintoneUrl(tab.url);
    if (!parsed.host) {
      setNotice('URLを解析できませんでした');
      return;
    }
    const existing = await loadFavorites();
    if (existing.some((item) => item.url === parsed.url)) {
      setNotice('このビューは既に登録済みです');
      return;
    }
    const label = (tab.title || '').replace(/ - kintone.*/, '') || '(no label)';
    const orderBase = existing.filter((item) => !item.pinned).length;
    const entry = {
      id: createId(),
      label,
      url: parsed.url,
      host: parsed.host,
      appId: parsed.appId,
      viewIdOrName: parsed.viewIdOrName,
      query: '',
      order: orderBase,
      pinned: false
    };
    const next = [...existing, entry];
    await saveFavorites(next);
    state.favorites = normalizeFavorites(next);
    state.badgeStatus.set(entry.id, { text: '-', title: '未取得', loading: false });
    renderLists();
    setNotice('ウォッチリストに追加しました。件数を取得しています…');
    refreshCounts().catch(() => {});
  } catch (error) {
    console.error('Failed to add favorite', error);
    setNotice('追加に失敗しました');
  }
}

async function editEntry(entry) {
  const currentLabel = entry.label || '';
  const nextLabel = window.prompt('表示名を編集', currentLabel);
  if (nextLabel == null) return;
  const currentUrl = entry.url || '';
  const nextUrl = window.prompt('URLを編集', currentUrl);
  if (nextUrl == null) return;
  if (!nextUrl.trim()) {
    setNotice('URLは必須です');
    return;
  }
  if (!isKintoneUrl(nextUrl.trim())) {
    setNotice('kintoneのURLのみ登録できます');
    return;
  }
  const parsed = parseKintoneUrl(nextUrl.trim());
  entry.label = nextLabel.trim() || '(no label)';
  entry.url = parsed.url;
  entry.host = parsed.host;
  entry.appId = parsed.appId;
  entry.viewIdOrName = parsed.viewIdOrName;
  await persistFavorites();
  setNotice('更新しました');
}

async function changeEntryIcon(entry) {
  const current = entry.icon || DEFAULT_ICON;
  const promptText = `アイコン名を入力してください。\n利用可能: ${ICON_OPTIONS.join(', ')}`;
  const input = window.prompt(promptText, current);
  if (input == null) return;
  const value = input.trim();
  entry.icon = ICON_OPTIONS.includes(value) ? value : DEFAULT_ICON;
  await persistFavorites();
  setNotice('アイコンを更新しました');
}

async function changeEntryCategory(entry) {
  const current = entry.category && entry.category !== DEFAULT_CATEGORY ? entry.category : '';
  const input = window.prompt('カテゴリ名を入力してください（空欄でその他）', current);
  if (input == null) return;
  const value = input.trim();
  entry.category = value || DEFAULT_CATEGORY;
  await persistFavorites();
  setNotice('カテゴリを更新しました');
}

async function deleteEntry(id) {
  if (!window.confirm('このショートカットを削除しますか？')) return;
  const next = state.favorites.filter((item) => item.id !== id);
  state.favorites = next;
  reindexOrders();
  await persistFavorites();
  state.badgeStatus.delete(id);
  renderLists();
}

async function togglePin(id, pinned) {
  const entry = state.favorites.find((item) => item.id === id);
  if (!entry) return;
  entry.pinned = pinned;
  if (pinned) {
    const max = Math.max(-1, ...state.favorites.filter((item) => item.pinned).map((item) => item.order ?? 0));
    entry.order = max + 1;
  } else {
    const max = Math.max(-1, ...state.favorites.filter((item) => !item.pinned).map((item) => item.order ?? 0));
    entry.order = max + 1;
  }
  reindexOrders();
  await persistFavorites();
  renderLists();
}

function reindexOrders() {
  const pinned = state.favorites.filter((item) => item.pinned).sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  pinned.forEach((item, index) => { item.order = index; });
  const normal = state.favorites.filter((item) => !item.pinned).sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  normal.forEach((item, index) => { item.order = index; });
}

async function persistFavorites() {
  await saveFavorites(state.favorites);
  state.favorites = normalizeFavorites(state.favorites);
  renderLists();
}

function handleDragStart(event) {
  const li = event.currentTarget;
  const id = li.dataset.id;
  if (!id) return;
  state.dragging = { id, pinned: li.dataset.pinned === 'true' };
  li.classList.add('is-dragging');
  event.dataTransfer.effectAllowed = 'move';
}

function handleDragOver(event) {
  if (!state.dragging) return;
  const li = event.currentTarget;
  if (li.dataset.pinned !== String(state.dragging.pinned)) return;
  event.preventDefault();
  const rect = li.getBoundingClientRect();
  const before = event.clientY < rect.top + rect.height / 2;
  li.classList.toggle('drop-before', before);
  li.classList.toggle('drop-after', !before);
}

function handleDragLeave(event) {
  event.currentTarget.classList.remove('drop-before', 'drop-after');
}

function handleDrop(event) {
  if (!state.dragging) return;
  const target = event.currentTarget;
  if (target.dataset.pinned !== String(state.dragging.pinned)) return;
  event.preventDefault();
  const rect = target.getBoundingClientRect();
  const before = event.clientY < rect.top + rect.height / 2;
  target.classList.remove('drop-before', 'drop-after');
  reorderWithinGroup(state.dragging.id, target.dataset.id, before, state.dragging.pinned);
}

function handleDragEnd(event) {
  event.currentTarget.classList.remove('is-dragging', 'drop-before', 'drop-after');
  state.dragging = null;
}

async function reorderWithinGroup(sourceId, targetId, before, pinned) {
  if (!sourceId || !targetId || sourceId === targetId) return;
  const subset = state.favorites.filter((item) => item.pinned === pinned).sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  const movingIndex = subset.findIndex((item) => item.id === sourceId);
  const targetIndex = subset.findIndex((item) => item.id === targetId);
  if (movingIndex === -1 || targetIndex === -1) return;
  const [moving] = subset.splice(movingIndex, 1);
  const insertIndex = before ? targetIndex : targetIndex + 1;
  subset.splice(insertIndex > subset.length ? subset.length : insertIndex, 0, moving);
  subset.forEach((item, index) => {
    item.order = index;
  });
  await persistFavorites();
}

function handleFilterKeydown(event) {
  if (event.key === 'Enter') {
    event.preventDefault();
    openSelection();
  } else if (event.key === 'ArrowDown') {
    event.preventDefault();
    moveSelection(1);
  } else if (event.key === 'ArrowUp') {
    event.preventDefault();
    moveSelection(-1);
  } else if (event.key === 'Escape') {
    event.preventDefault();
    els.filter.value = '';
    applyFilterText('');
  }
}

function handleMenuLayerClick(event) {
  if (event.target === els.menuLayer) {
    closeMenu();
  }
}

function handleGlobalKeydown(event) {
  if (event.key === 'Escape' && state.menu) {
    closeMenu();
  }
}

function focusFilterSoon() {
  if (!els.filter) return;
  setTimeout(() => {
    try {
      els.filter.focus();
      els.filter.select();
    } catch (_err) {
      // ignore
    }
  }, 50);
}

function wireEvents() {
  els.filter?.addEventListener('input', (event) => applyFilterText(event.target.value));
  els.filter?.addEventListener('keydown', handleFilterKeydown);
  els.clearFilter?.addEventListener('click', () => {
    if (els.filter) {
      els.filter.value = '';
      applyFilterText('');
      els.filter.focus();
    }
  });
  els.togglePinsOnly?.addEventListener('change', (event) => {
    state.pinsOnly = Boolean(event.target.checked);
    renderLists();
  });
  els.refreshCounts?.addEventListener('click', () => refreshCounts().catch(() => {}));
  els.quickAdd?.addEventListener('click', () => quickAddFromActiveTab());
  els.openOptions?.addEventListener('click', () => {
    if (chrome?.runtime?.openOptionsPage) {
      chrome.runtime.openOptionsPage();
    } else {
      window.open(chrome.runtime.getURL('options.html'), '_blank');
    }
  });
  els.menuLayer?.addEventListener('mousedown', handleMenuLayerClick);
  document.addEventListener('keydown', handleGlobalKeydown);
  window.addEventListener('blur', boundCloseMenuOnBlur);
  if (els.toggleShortcuts) {
    const handler = () => {
      toggleShortcutVisibility().catch(() => {});
    };
    els.toggleShortcuts.addEventListener('click', handler);
    state.shortcutToggleHandler = handler;
  }
}

function dispose() {
  stopCountTimer();
  detachStorageListener();
  document.removeEventListener('keydown', handleGlobalKeydown);
  window.removeEventListener('blur', boundCloseMenuOnBlur);
  if (els.toggleShortcuts && state.shortcutToggleHandler) {
    els.toggleShortcuts.removeEventListener('click', state.shortcutToggleHandler);
    state.shortcutToggleHandler = null;
  }
}

async function init() {
  wireEvents();
  attachStorageListener();
  await refreshShortcutEntries();
  await loadFavoritesAndRender();
  refreshCounts().catch(() => {});
  startCountTimer();
  focusFilterSoon();
}

init().catch((error) => {
  console.error('Failed to initialize launcher', error);
  setNotice('初期化に失敗しました');
});

window.addEventListener('beforeunload', dispose);
