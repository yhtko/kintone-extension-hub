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
  createId,
  loadRecordPinVisibility,
  loadRecentRecords,
  saveRecentRecords,
  loadAppNameMap,
  saveAppNameMap
} from './core.js';
import { initPins } from './pins.js';

const REFRESH_INTERVAL = 30000;
const MAX_SHORTCUT_INITIAL_LENGTH = 2;

const doc = document;

const els = {
  filterPanel: doc.getElementById('filterPanel'),
  toggleFilterPanel: doc.getElementById('toggleFilterPanel'),
  filter: doc.getElementById('filter'),
  clearFilter: doc.getElementById('clearFilter'),
  togglePinsOnly: doc.getElementById('togglePinsOnly'),
  togglePinsOnlyIcon: doc.getElementById('togglePinsOnlyIcon'),
  pinnedList: doc.getElementById('pinnedWatchList'),
  pinSection: doc.getElementById('pinSection'),
  recentSection: doc.getElementById('recentSection'),
  recentList: doc.getElementById('recentList'),
  recentClear: doc.getElementById('recentClear'),
  recordPinSection: doc.getElementById('recordPinSection'),
  favSection: doc.getElementById('favSection'),
  favCategories: doc.getElementById('favCategories'),
  notice: doc.getElementById('notice'),
  refreshCounts: doc.getElementById('refreshCounts'),
  quickAdd: doc.getElementById('quickAdd'),
  openOptions: doc.getElementById('openOptions'),
  shortcutRow: doc.getElementById('shortcutRow'),
  shortcutPanel: doc.getElementById('shortcutPanel'),
  toggleShortcuts: doc.getElementById('toggleShortcuts'),
  toggleRecordPins: doc.getElementById('toggleRecordPins'),
  pinList: doc.getElementById('pinList')
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
  dragging: null,
  shortcutToggleHandler: null,
  recentRecords: [],
  recentHydrationSeq: 0,
  filterPanelVisible: false,
  recordPinsCollapsed: false
};

const shortcutState = {
  entries: [],
  visible: true,
  warmedHosts: new Set()
};
const recordPinState = {
  visible: true,
  controller: null
};
const DEFAULT_ICON = 'file-text';
const DEFAULT_CATEGORY = 'その他';
const ICON_OPTIONS = [
  'clipboard', 'file-text', 'package', 'box', 'truck', 'factory', 'wrench', 'calendar',
  'list-checks', 'search', 'chart-bar', 'receipt', 'users', 'settings', 'bookmark', 'star', 'history'
];
const DEFAULT_ICON_COLOR = 'gray';
const ICON_COLOR_OPTIONS = ['gray', 'blue', 'green', 'orange', 'red', 'purple'];

function normalizeIconName(value) {
  const name = typeof value === 'string' ? value.trim() : '';
  return ICON_OPTIONS.includes(name) ? name : DEFAULT_ICON;
}

function normalizeIconColor(value) {
  const color = typeof value === 'string' ? value.trim().toLowerCase() : '';
  return ICON_COLOR_OPTIONS.includes(color) ? color : DEFAULT_ICON_COLOR;
}

function normalizeCategoryName(value) {
  return typeof value === 'string' ? value.trim() : '';
}

function getCategoryLabel(value) {
  const name = normalizeCategoryName(value);
  return name || DEFAULT_CATEGORY;
}

function toPascalIconName(name) {
  return String(name || '')
    .trim()
    .split(/[-_\s]+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
    .join('');
}

function iconToSvg(name, size = 16) {
  const lucide = globalThis.lucide;
  const icons = lucide?.icons;
  if (!icons) return '';
  const kebab = normalizeIconName(name);
  const candidates = [kebab, toPascalIconName(kebab)];
  for (const key of candidates) {
    const iconNode = icons[key];
    if (!iconNode) continue;
    if (typeof iconNode.toSvg === 'function') {
      return iconNode.toSvg({ width: size, height: size });
    }
    if (typeof lucide.createElement === 'function') {
      const node = lucide.createElement(iconNode, { width: size, height: size });
      return node?.outerHTML || '';
    }
  }
  return '';
}

function renderLucideIcons(root = doc) {
  const nodes = root.querySelectorAll('.lc[data-icon]');
  nodes.forEach((el) => {
    const iconName = normalizeIconName(el.dataset.icon);
    el.dataset.icon = iconName;
    const svg = iconToSvg(iconName, 16);
    el.innerHTML = svg;
  });
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

function resolveShortcutIcon(entry) {
  return normalizeIconName(entry?.icon);
}

function resolveShortcutIconColor(entry) {
  return normalizeIconColor(entry?.iconColor);
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

function setFilterPanelVisible(visible, { focus = false } = {}) {
  state.filterPanelVisible = Boolean(visible);
  if (els.filterPanel) {
    els.filterPanel.classList.toggle('hidden', !state.filterPanelVisible);
  }
  if (els.toggleFilterPanel) {
    els.toggleFilterPanel.setAttribute('aria-pressed', state.filterPanelVisible ? 'true' : 'false');
    els.toggleFilterPanel.title = state.filterPanelVisible ? '検索を閉じる' : '検索を表示';
  }
  if (focus && state.filterPanelVisible && els.filter) {
    setTimeout(() => {
      try {
        els.filter.focus();
        els.filter.select();
      } catch (_err) {
        // ignore
      }
    }, 0);
  }
}

function setRecordPinsCollapsed(collapsed) {
  state.recordPinsCollapsed = Boolean(collapsed);
  if (els.pinList) {
    els.pinList.classList.toggle('hidden', state.recordPinsCollapsed);
  }
  if (els.toggleRecordPins) {
    els.toggleRecordPins.textContent = state.recordPinsCollapsed ? '▸' : '▾';
    els.toggleRecordPins.setAttribute('aria-pressed', state.recordPinsCollapsed ? 'true' : 'false');
    els.toggleRecordPins.title = state.recordPinsCollapsed ? 'レコードピンを展開' : 'レコードピンを折りたたむ';
  }
}

function setWatchlistCollapsed(collapsed) {
  state.pinsOnly = Boolean(collapsed);
  if (els.togglePinsOnly) {
    els.togglePinsOnly.checked = state.pinsOnly;
    els.togglePinsOnly.setAttribute('aria-pressed', state.pinsOnly ? 'true' : 'false');
  }
  if (els.favCategories) {
    els.favCategories.classList.toggle('is-collapsed', state.pinsOnly);
  }
  if (els.togglePinsOnlyIcon) {
    els.togglePinsOnlyIcon.textContent = state.pinsOnly ? '▸' : '▾';
    els.togglePinsOnlyIcon.title = state.pinsOnly ? 'ウォッチリストを展開' : 'ウォッチリストを折りたたむ';
  }
}

function formatRecentTime(value) {
  const ts = Number(value);
  if (!Number.isFinite(ts)) return '';
  try {
    const date = new Date(ts);
    return date.toLocaleString([], {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit'
    });
  } catch (_err) {
    return '';
  }
}

function getRecentAppLabel(entry) {
  const appName = String(entry?.appName || '').trim();
  if (appName) return appName;
  const appId = String(entry?.appId || '').trim();
  return appId ? `App ${appId}` : 'App';
}

function normalizeRecentHost(host) {
  const raw = String(host || '').trim();
  if (!raw) return '';
  try {
    return new URL(raw).origin.toLowerCase().replace(/\/$/, '');
  } catch (_err) {
    return raw.toLowerCase().replace(/\/$/, '');
  }
}

function cssEscapeValue(value) {
  const raw = String(value || '');
  if (globalThis.CSS?.escape) return globalThis.CSS.escape(raw);
  return raw.replace(/[^a-zA-Z0-9_-]/g, (ch) => `\\${ch}`);
}

function setRecentAppNameInState(host, appId, appName) {
  const safeHost = normalizeRecentHost(host);
  const safeAppId = String(appId || '').trim();
  const safeName = String(appName || '').trim();
  if (!safeHost || !safeAppId || !safeName) return;
  state.recentRecords.forEach((item) => {
    if (normalizeRecentHost(item.host) === safeHost && String(item.appId || '').trim() === safeAppId) {
      item.appName = safeName;
    }
  });
}

function updateRecentAppNameDom(host, appId, appName) {
  const safeHost = normalizeRecentHost(host);
  const safeAppId = String(appId || '').trim();
  const safeName = String(appName || '').trim();
  if (!safeHost || !safeAppId || !safeName) return 0;
  const selector = `.recent-item[data-host="${cssEscapeValue(safeHost)}"][data-app-id="${cssEscapeValue(safeAppId)}"] .recent-item-app`;
  const nodes = doc.querySelectorAll(selector);
  nodes.forEach((el) => {
    el.textContent = safeName;
  });
  return nodes.length;
}

async function fillRecentAppNamesFromCache() {
  const list = Array.isArray(state.recentRecords) ? state.recentRecords : [];
  if (!list.length) return;
  const hosts = new Set();
  list.forEach((entry) => {
    const host = normalizeRecentHost(entry.host);
    if (host) hosts.add(host);
  });
  for (const host of hosts.values()) {
    try {
      const { map } = await loadAppNameMap(host);
      Object.entries(map).forEach(([appId, appName]) => {
        setRecentAppNameInState(host, appId, appName);
      });
    } catch (_err) {
      // ignore
    }
  }
}

async function hydrateRecentAppNames() {
  const seq = ++state.recentHydrationSeq;
  const list = Array.isArray(state.recentRecords) ? state.recentRecords.slice() : [];
  if (!list.length) return;

  const appIdsByHost = new Map();
  list.forEach((entry) => {
    const host = normalizeRecentHost(entry.host);
    const appId = String(entry.appId || '').trim();
    if (!host || !appId) return;
    if (!appIdsByHost.has(host)) appIdsByHost.set(host, new Set());
    appIdsByHost.get(host).add(appId);
  });

  let updatedCount = 0;
  for (const [host, set] of appIdsByHost.entries()) {
    if (seq !== state.recentHydrationSeq) return;
    const appIds = Array.from(set).filter(Boolean);
    if (!appIds.length) continue;

    let cache = { savedAt: 0, map: {}, fresh: false };
    try {
      cache = await loadAppNameMap(host);
    } catch (_err) {
      // ignore
    }
    if (seq !== state.recentHydrationSeq) return;

    if (cache.fresh) {
      for (const appId of appIds) {
        const appName = String(cache.map[appId] || '').trim();
        if (!appName) continue;
        setRecentAppNameInState(host, appId, appName);
        updatedCount += updateRecentAppNameDom(host, appId, appName);
      }
      continue;
    }

    try {
      const res = await chrome.runtime.sendMessage({ type: 'GET_APPS_MAP', host, appIds });
      if (seq !== state.recentHydrationSeq) return;
      if (!res?.ok || !res?.map || typeof res.map !== 'object') {
        console.debug('[kfav] recent app names hydrate skipped', { host, reason: res?.error || 'no map' });
        continue;
      }
      const hostMap = res.map;
      await saveAppNameMap(host, hostMap);
      for (const appId of appIds) {
        const appName = String(hostMap[appId] || '').trim();
        if (!appName) continue;
        setRecentAppNameInState(host, appId, appName);
        updatedCount += updateRecentAppNameDom(host, appId, appName);
      }
    } catch (_err) {
      console.debug('[kfav] recent app names hydrate failed', { host });
    }
  }
  console.debug('[kfav] recent app names updated', { updatedCount });
}

function renderRecentRecords() {
  if (!els.recentSection || !els.recentList) return;
  const list = Array.isArray(state.recentRecords) ? state.recentRecords : [];
  els.recentList.textContent = '';
  if (els.recentClear) {
    els.recentClear.disabled = list.length === 0;
  }
  if (!list.length) {
    els.recentSection.classList.add('hidden');
    return;
  }
  list.forEach((item) => {
    const li = doc.createElement('li');
    const btn = doc.createElement('button');
    btn.type = 'button';
    btn.className = 'recent-item';
    btn.dataset.host = normalizeRecentHost(item.host);
    btn.dataset.appId = String(item.appId || '').trim();
    btn.title = item.url || '';
    btn.addEventListener('click', () => openRecentRecord(item));

    const icon = doc.createElement('span');
    icon.className = 'recent-item-icon lc';
    icon.dataset.icon = 'history';
    icon.setAttribute('aria-hidden', 'true');

    const text = doc.createElement('div');
    text.className = 'recent-item-text';
    const title = doc.createElement('div');
    title.className = 'recent-title recent-item-title';
    const appLabel = getRecentAppLabel(item);
    const app = doc.createElement('span');
    app.className = 'recent-item-app';
    app.textContent = appLabel;
    title.appendChild(app);
    const sub = doc.createElement('div');
    sub.className = 'recent-sub recent-item-meta';
    sub.textContent = `#${item.recordId}`;
    text.appendChild(title);
    text.appendChild(sub);

    const time = doc.createElement('span');
    time.className = 'recent-item-time';
    time.textContent = formatRecentTime(item.visitedAt || item.lastSeenAt);

    btn.appendChild(icon);
    btn.appendChild(text);
    btn.appendChild(time);
    li.appendChild(btn);
    els.recentList.appendChild(li);
  });
  els.recentSection.classList.remove('hidden');
  renderLucideIcons(els.recentSection);
}

async function loadRecentAndRender() {
  try {
    state.recentRecords = await loadRecentRecords();
    await fillRecentAppNamesFromCache();
  } catch (error) {
    console.error('Failed to load recent records', error);
    state.recentRecords = [];
  }
  renderRecentRecords();
  hydrateRecentAppNames().catch((error) => {
    console.error('Failed to hydrate recent app names', error);
  });
}

async function clearRecentRecords() {
  try {
    state.recentHydrationSeq += 1;
    await saveRecentRecords([]);
    state.recentRecords = [];
    renderRecentRecords();
  } catch (error) {
    console.error('Failed to clear recent records', error);
  }
}

async function openRecentRecord(entry) {
  if (!entry?.url) return;
  try {
    if (chrome?.tabs?.create) {
      await chrome.tabs.create({ url: entry.url, active: true });
    } else {
      window.open(entry.url, '_blank', 'noopener');
    }
  } catch (_err) {
    window.open(entry.url, '_blank', 'noopener');
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
      icon: normalizeIconName(item.icon),
      iconColor: normalizeIconColor(item.iconColor),
      category: normalizeCategoryName(item.category)
    }))
  );
}

function applyFilterText(value) {
  // #filter is local filter only (pinned/watchlist visibility).
  state.filterText = value || '';
  renderLists();
}

function splitFavoritesByPin() {
  const pinned = [];
  const normal = [];
  state.favorites.forEach((item) => {
    if (item.pinned) pinned.push(item);
    else normal.push(item);
  });
  pinned.sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  normal.sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  return { pinned, normal };
}

function applySearchVisibility() {
  const keyword = state.filterText.trim().toLowerCase();
  const allEntries = Array.from(doc.querySelectorAll('.entry[data-id]'));
  allEntries.forEach((entryEl) => {
    const haystack = entryEl.dataset.search || '';
    const matched = !keyword || haystack.includes(keyword);
    entryEl.classList.toggle('entry-filter-hidden', !matched);
  });

  if (els.favCategories) {
    const blocks = els.favCategories.querySelectorAll('.category-block');
    blocks.forEach((block) => {
      const hasVisible = Array.from(block.querySelectorAll('.entry')).some((entryEl) => !entryEl.classList.contains('entry-filter-hidden'));
      block.classList.toggle('hidden-by-filter', !hasVisible);
    });
  }

  if (els.pinSection) {
    const hasPinned = Boolean(els.pinnedList?.querySelector('.entry'));
    const hasPinnedVisible = Boolean(els.pinnedList?.querySelector('.entry:not(.entry-filter-hidden)'));
    const hidePinned = !hasPinned || (keyword.length > 0 && !hasPinnedVisible);
    els.pinSection.classList.toggle('hidden', hidePinned);
  }
}

function isSelectableEntry(entryEl) {
  if (!entryEl?.dataset?.id) return false;
  if (entryEl.classList.contains('entry-filter-hidden')) return false;
  if (entryEl.closest('.category-block.hidden-by-filter')) return false;
  if (entryEl.closest('.hidden')) return false;
  if (entryEl.closest('#favCategories.is-collapsed')) return false;
  return true;
}

function syncSelectionState() {
  const previousId = state.selectionIds[state.selectionIndex] || null;
  const visibleEntries = Array.from(doc.querySelectorAll('.entry[data-id]')).filter(isSelectableEntry);
  state.selectionIds = visibleEntries.map((item) => item.dataset.id);
  if (!state.selectionIds.length) {
    state.selectionIndex = -1;
    return;
  }
  const keepIndex = previousId ? state.selectionIds.indexOf(previousId) : -1;
  state.selectionIndex = keepIndex >= 0 ? keepIndex : 0;
}

function renderLists() {
  const { pinned, normal } = splitFavoritesByPin();
  renderPinnedList(pinned);
  renderCategorySections(normal);
  setWatchlistCollapsed(state.pinsOnly);
  applySearchVisibility();
  renderLucideIcons();
  syncSelectionState();
  highlightSelection();
}

function renderPinnedList(items) {
  if (!els.pinnedList) return [];
  els.pinnedList.textContent = '';
  if (!items.length) {
    const empty = doc.createElement('li');
    empty.className = 'empty-message';
    const hasFilter = state.filterText.trim().length > 0;
    empty.textContent = hasFilter
      ? 'No pinned items matched your search'
      : 'No pinned items yet';
    els.pinnedList.appendChild(empty);
    return [];
  }
  items.forEach((item) => {
    els.pinnedList.appendChild(createEntryElement(item, true));
  });
  return items.slice();
}

function groupByCategory(items) {
  const map = new Map();
  items.forEach((item) => {
    const key = getCategoryLabel(item.category);
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
      ? 'No favorites matched your search'
      : 'No favorites yet';
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
  const iconName = normalizeIconName(entry.icon);
  const categoryLabel = getCategoryLabel(entry.category);
  const li = doc.createElement('li');
  li.className = 'entry';
  li.dataset.id = entry.id;
  li.dataset.pinned = String(pinned);
  li.dataset.category = categoryLabel;
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
  icon.className = 'entry-icon kp-ico lc';
  icon.dataset.icon = iconName;
  icon.dataset.icoColor = normalizeIconColor(entry.iconColor);
  icon.setAttribute('aria-hidden', 'true');

  const text = doc.createElement('div');
  text.className = 'entry-text';
  const title = doc.createElement('div');
  title.className = 'entry-title';
  title.textContent = entry.label || '(no label)';
  const sub = doc.createElement('div');
  sub.className = 'entry-sub';
  const view = entry.viewIdOrName ? `view:${entry.viewIdOrName}` : '';
  sub.textContent = view;
  sub.classList.toggle('hidden', !view);
  li.dataset.search = [
    entry.label || '',
    entry.url || '',
    entry.host || '',
    categoryLabel,
    entry.viewIdOrName || '',
    entry.query || ''
  ].join('\n').toLowerCase();
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

  li.appendChild(button);
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
  const stateItem = state.badgeStatus.get(id) || { text: '-', title: 'Not fetched', loading: false };
  el.textContent = stateItem.text ?? '-';
  el.title = stateItem.title || '';
  el.classList.toggle('loading', Boolean(stateItem.loading));
}

function updateBadgeDom(id) {
  const el = doc.querySelector(`.entry-badge[data-id="${escapeId(id)}"]`);
  if (el) syncBadgeToElement(id, el);
}

function setBadgeLoading(id, loading) {
  const prev = state.badgeStatus.get(id) || { text: '-', title: 'Not fetched', loading: false };
  state.badgeStatus.set(id, { ...prev, loading: Boolean(loading) });
  updateBadgeDom(id);
}

function setBadgeValue(id, value, title = '') {
  const text = value == null ? '-' : String(value);
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
  renderShortcuts();
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
        items.forEach((item) => setBadgeValue(item.id, null, 'Host permission required'));
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
            setBadgeValue(item.id, null, 'Count was not returned for this item');
            failedHosts.add(host);
          }
        }
      } else {
        const message = res?.error || 'Failed to fetch counts';
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
    setNotice(`Permission is missing for hosts: ${Array.from(missingPerm).join(', ')}`);
  } else if (failedHosts.size) {
    setNotice('Some hosts failed to fetch counts. Check the log for details.');
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
    if (Object.prototype.hasOwnProperty.call(changes, 'kfavRecentRecords')) {
      loadRecentAndRender().catch((error) => {
        console.error('Failed to reload recent records', error);
      });
    }
    if (Object.prototype.hasOwnProperty.call(changes, 'kfavPins')) {
      try {
        if (recordPinState.visible) {
          if (recordPinState.controller?.reload) {
            const task = recordPinState.controller.reload();
            if (task?.catch) {
              task.catch((error) => console.error('Failed to reload record pins', error));
            }
          } else {
            disposeRecordPins();
            recordPinState.controller = initPins();
          }
        }
      } catch (error) {
        console.error('Failed to reload record pins', error);
      }
    }
    if (Object.prototype.hasOwnProperty.call(changes, 'kfavShortcutsVisible')) {
      const next = typeof changes.kfavShortcutsVisible.newValue === 'boolean'
        ? changes.kfavShortcutsVisible.newValue
        : true;
      applyShortcutVisibility(next);
    }
    if (Object.prototype.hasOwnProperty.call(changes, 'kfavRecordPinsVisible')) {
      const nextRecordPins = typeof changes.kfavRecordPinsVisible.newValue === 'boolean'
        ? changes.kfavRecordPinsVisible.newValue
        : true;
      applyRecordPinVisibility(nextRecordPins);
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

function applyShortcutVisibility(visible) {
  shortcutState.visible = visible;
  if (els.shortcutPanel) {
    els.shortcutPanel.classList.toggle('collapsed', !visible);
  }
  if (els.toggleShortcuts) {
    els.toggleShortcuts.setAttribute('aria-pressed', visible ? 'true' : 'false');
    els.toggleShortcuts.textContent = visible ? '▾' : '▸';
    els.toggleShortcuts.title = visible ? 'ショートカットを折りたたむ' : 'ショートカットを展開';
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
    empty.title = 'Configure shortcuts';
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
    const label = (entry.label || '').trim();
    const iconColor = resolveShortcutIconColor(entry);
    btn.className = `shortcut-button shortcut-${type}`;
    btn.title = label;
    btn.setAttribute('aria-label', label);
    btn.dataset.tooltip = label;
    btn.dataset.icoColor = iconColor;
    const url = buildShortcutUrl(entry);
    if (!url) {
      btn.disabled = true;
      btn.classList.add('disabled');
    }
    const icon = doc.createElement('span');
    icon.className = 'shortcut-icon lc';
    icon.dataset.icon = resolveShortcutIcon(entry);
    icon.dataset.icoColor = iconColor;
    icon.setAttribute('aria-hidden', 'true');

    btn.appendChild(icon);
    btn.addEventListener('click', () => openShortcut(entry));
    els.shortcutRow.appendChild(btn);
  });
  renderLucideIcons(els.shortcutRow);
}
async function openShortcut(entry) {
  const url = buildShortcutUrl(entry);
  if (!url) {
    setNotice('Shortcut URL is not configured');
    return;
  }
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
      setNotice('Open a kintone tab first');
      return;
    }
    const parsed = parseKintoneUrl(tab.url);
    if (!parsed.host) {
      setNotice('Could not parse the URL');
      return;
    }
    const existing = await loadFavorites();
    if (existing.some((item) => item.url === parsed.url)) {
      setNotice('This view is already registered');
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
    state.badgeStatus.set(entry.id, { text: '-', title: 'Not fetched', loading: false });
    renderLists();
    setNotice('Added to favorites. Fetching counts...');
    refreshCounts().catch(() => {});
  } catch (error) {
    console.error('Failed to add favorite', error);
    setNotice('Failed to add favorite');
  }
}

async function editEntry(entry) {
  const currentLabel = entry.label || '';
  const nextLabel = window.prompt('Edit label', currentLabel);
  if (nextLabel == null) return;
  const currentUrl = entry.url || '';
  const nextUrl = window.prompt('Edit URL', currentUrl);
  if (nextUrl == null) return;
  if (!nextUrl.trim()) {
    setNotice('URL is required');
    return;
  }
  if (!isKintoneUrl(nextUrl.trim())) {
    setNotice('Please enter a kintone URL');
    return;
  }
  const parsed = parseKintoneUrl(nextUrl.trim());
  entry.label = nextLabel.trim() || '(no label)';
  entry.url = parsed.url;
  entry.host = parsed.host;
  entry.appId = parsed.appId;
  entry.viewIdOrName = parsed.viewIdOrName;
  await persistFavorites();
  setNotice('Updated');
}

async function changeEntryIcon(entry) {
  const current = normalizeIconName(entry.icon);
  const promptText = `Enter icon name\nChoices: ${ICON_OPTIONS.join(', ')}`;
  const input = window.prompt(promptText, current);
  if (input == null) return;
  const value = input.trim();
  entry.icon = normalizeIconName(value);
  await persistFavorites();
  setNotice('Icon updated');
}

async function changeEntryCategory(entry) {
  const current = normalizeCategoryName(entry.category);
  const input = window.prompt('Enter category (empty = Other)', current);
  if (input == null) return;
  entry.category = normalizeCategoryName(input);
  await persistFavorites();
  setNotice('Category updated');
}

async function deleteEntry(id) {
  if (!window.confirm('Delete this shortcut?')) return;
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

function disposeRecordPins() {
  if (recordPinState.controller?.dispose) {
    recordPinState.controller.dispose();
  }
  recordPinState.controller = null;
}

function applyRecordPinVisibility(visible) {
  const flag = Boolean(visible);
  recordPinState.visible = flag;
  if (els.recordPinSection) {
    els.recordPinSection.classList.toggle('hidden', !flag);
  }
  if (!els.recordPinSection) return;
  if (flag) {
    if (!recordPinState.controller) {
      recordPinState.controller = initPins();
    }
  } else {
    disposeRecordPins();
  }
}

async function initializeRecordPins() {
  if (!els.recordPinSection) return;
  try {
    const visible = await loadRecordPinVisibility();
    applyRecordPinVisibility(visible);
  } catch (error) {
    console.error('Failed to load record pin visibility', error);
    applyRecordPinVisibility(true);
  }
}

function focusFilterSoon() {
  if (!els.filter || !state.filterPanelVisible) return;
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
  setFilterPanelVisible(false);
  setRecordPinsCollapsed(false);
  setWatchlistCollapsed(Boolean(els.togglePinsOnly?.checked));
  els.toggleFilterPanel?.addEventListener('click', () => {
    const next = !state.filterPanelVisible;
    setFilterPanelVisible(next, { focus: next });
  });
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
    setWatchlistCollapsed(Boolean(event.target.checked));
    syncSelectionState();
    highlightSelection();
  });
  els.refreshCounts?.addEventListener('click', () => refreshCounts().catch(() => {}));
  els.quickAdd?.addEventListener('click', () => quickAddFromActiveTab());
  els.toggleRecordPins?.addEventListener('click', () => {
    setRecordPinsCollapsed(!state.recordPinsCollapsed);
  });
  els.openOptions?.addEventListener('click', () => {
    if (chrome?.runtime?.openOptionsPage) {
      chrome.runtime.openOptionsPage();
    } else {
      window.open(chrome.runtime.getURL('options.html'), '_blank');
    }
  });
  els.recentClear?.addEventListener('click', () => {
    clearRecentRecords().catch(() => {});
  });
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
  disposeRecordPins();
  if (els.toggleShortcuts && state.shortcutToggleHandler) {
    els.toggleShortcuts.removeEventListener('click', state.shortcutToggleHandler);
    state.shortcutToggleHandler = null;
  }
}

async function init() {
  wireEvents();
  attachStorageListener();
  await initializeRecordPins();
  await refreshShortcutEntries();
  await loadRecentAndRender();
  await loadFavoritesAndRender();
  refreshCounts().catch(() => {});
  startCountTimer();
  focusFilterSoon();
}

init().catch((error) => {
  console.error('Failed to initialize launcher', error);
  setNotice('Initialization failed');
});

window.addEventListener('beforeunload', dispose);
