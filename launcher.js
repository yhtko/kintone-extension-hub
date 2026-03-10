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
const SHORTCUT_SEARCH_OPEN_MODE_KEY = 'shortcutSearchOpenMode';
const SHORTCUT_SEARCH_OPEN_MODE_CURRENT_TAB = 'current_tab';
const SHORTCUT_SEARCH_OPEN_MODE_NEW_TAB = 'new_tab';
const COLLAPSED_SECTIONS_KEY = 'kfavLauncherCollapsedSections';
const UI_LANGUAGE_KEY = 'uiLanguage';
const UI_LANGUAGE_VALUES = ['auto', 'ja', 'en'];
const DEFAULT_UI_LANGUAGE = 'auto';
const DEFAULT_LANGUAGE = 'ja';
const DEFAULT_LOCALE = 'ja';

const I18N_MESSAGES = {
  ja: {
    panel_title: 'kintone Base Launcher',
    panel_search_label: '検索',
    panel_search_placeholder: '検索（ラベル / URL / ホスト）',
    panel_shortcuts_toolbar: 'ショートカット一覧',
    panel_section_shortcuts: 'ショートカット',
    panel_section_watchlist_pinned: 'ウォッチリスト（ピン）',
    panel_section_watchlist: 'ウォッチリスト',
    panel_section_record_pins: 'レコードピン',
    panel_section_recent_records: '最近のレコード',
    panel_action_toggle_search_show: '検索を表示',
    panel_action_toggle_search_hide: '検索を隠す',
    panel_action_open_settings: '設定を開く',
    panel_action_clear_search: '検索をクリア',
    panel_action_toggle_shortcuts_collapse: 'ショートカットを折りたたむ',
    panel_action_toggle_shortcuts_expand: 'ショートカットを展開',
    panel_action_add_current_view: '現在のkintoneビューを追加',
    panel_action_refresh_counts: '件数を更新',
    panel_action_toggle_watchlist_collapse: 'ウォッチリストを折りたたむ',
    panel_action_toggle_watchlist_expand: 'ウォッチリストを展開',
    panel_action_add_current_record_pin: '現在のレコードをピン留め',
    panel_action_refresh_pins: 'ピンを更新',
    panel_action_toggle_record_pins_collapse: 'レコードピンを折りたたむ',
    panel_action_toggle_record_pins_expand: 'レコードピンを展開',
    panel_action_toggle_recent_records_collapse: '最近のレコードを折りたたむ',
    panel_action_toggle_recent_records_expand: '最近のレコードを展開',
    panel_action_clear_recent_records: '履歴をクリア',
    panel_action_app_search: 'アプリ検索',
    panel_app_search_placeholder: 'アプリ名で検索',
    panel_app_search_load_failed: 'アプリ一覧を取得できませんでした',
    panel_app_search_no_results: '一致するアプリがありません',
    panel_collapsed_sections_label: '折りたたみ済みセクション',
    panel_collapsed_watchlist: 'ウォッチリスト',
    panel_collapsed_record_pins: 'レコードピン',
    panel_collapsed_recent_records: '最近のレコード',
    panel_recent_app_fallback: 'App',
    panel_recent_app_prefix: 'App',
    panel_empty_pinned_no_match: '検索条件に一致するピン留めはありません',
    panel_empty_pinned: 'ピン留めされた項目はありません',
    panel_empty_favorites_no_match: '検索条件に一致するウォッチリストはありません',
    panel_empty_favorites: 'ウォッチリストはまだありません',
    panel_empty_shortcuts: 'ショートカットを設定',
    panel_entry_no_label: '(no label)',
    panel_badge_not_fetched: '未取得',
    panel_badge_permission_required: 'ホスト権限が必要です',
    panel_badge_missing_result: 'この項目の件数が返されませんでした',
    panel_badge_fetch_failed: '件数取得に失敗しました',
    panel_notice_permission_missing_hosts: '次のホスト権限が不足しています: {hosts}',
    panel_notice_counts_failed: '一部ホストの件数取得に失敗しました。ログを確認してください。',
    panel_notice_shortcut_url_missing: 'ショートカットURLが設定されていません',
    panel_notice_open_kintone_first: '先にkintoneタブを開いてください',
    panel_notice_parse_url_failed: 'URLを解析できませんでした',
    panel_notice_already_registered: 'このビューは既に登録済みです',
    panel_notice_added_fetching: 'ウォッチリストに追加しました。件数を取得しています...',
    panel_notice_add_failed: 'ウォッチリストの追加に失敗しました',
    panel_notice_url_required: 'URLは必須です',
    panel_notice_kintone_url_required: 'kintone URLを入力してください',
    panel_notice_updated: '更新しました',
    panel_notice_icon_updated: 'アイコンを更新しました',
    panel_notice_category_updated: 'カテゴリを更新しました',
    panel_notice_initialization_failed: '初期化に失敗しました',
    panel_prompt_edit_label: 'ラベルを編集',
    panel_prompt_edit_url: 'URLを編集',
    panel_prompt_icon: 'アイコン名を入力\n選択肢: {icons}',
    panel_prompt_category: 'カテゴリを入力（空欄 = その他）',
    panel_confirm_delete_shortcut: 'このショートカットを削除しますか？',
    panel_meta_view_prefix: 'view',
    panel_search_panel_placeholder: 'パネル内を検索'
  },
  en: {
    panel_title: 'kintone Base Launcher',
    panel_search_label: 'Search',
    panel_search_placeholder: 'Search (label / URL / host)',
    panel_shortcuts_toolbar: 'Shortcut list',
    panel_section_shortcuts: 'Shortcuts',
    panel_section_watchlist_pinned: 'Pinned watchlist',
    panel_section_watchlist: 'Watchlist',
    panel_section_record_pins: 'Record pins',
    panel_section_recent_records: 'Recent records',
    panel_action_toggle_search_show: 'Show search',
    panel_action_toggle_search_hide: 'Hide search',
    panel_action_open_settings: 'Open settings',
    panel_action_clear_search: 'Clear search',
    panel_action_toggle_shortcuts_collapse: 'Collapse shortcuts',
    panel_action_toggle_shortcuts_expand: 'Expand shortcuts',
    panel_action_add_current_view: 'Add current kintone view',
    panel_action_refresh_counts: 'Refresh counts',
    panel_action_toggle_watchlist_collapse: 'Collapse watchlist',
    panel_action_toggle_watchlist_expand: 'Expand watchlist',
    panel_action_add_current_record_pin: 'Pin current record',
    panel_action_refresh_pins: 'Refresh pins',
    panel_action_toggle_record_pins_collapse: 'Collapse record pins',
    panel_action_toggle_record_pins_expand: 'Expand record pins',
    panel_action_toggle_recent_records_collapse: 'Collapse recent records',
    panel_action_toggle_recent_records_expand: 'Expand recent records',
    panel_action_clear_recent_records: 'Clear history',
    panel_action_app_search: 'App search',
    panel_app_search_placeholder: 'Search app name',
    panel_app_search_load_failed: 'App list could not be loaded',
    panel_app_search_no_results: 'No apps matched',
    panel_collapsed_sections_label: 'Collapsed sections',
    panel_collapsed_watchlist: 'Watchlist',
    panel_collapsed_record_pins: 'Record pins',
    panel_collapsed_recent_records: 'Recent records',
    panel_recent_app_fallback: 'App',
    panel_recent_app_prefix: 'App',
    panel_empty_pinned_no_match: 'No pinned items matched your search',
    panel_empty_pinned: 'No pinned items yet',
    panel_empty_favorites_no_match: 'No watchlist items matched your search',
    panel_empty_favorites: 'No watchlist items yet',
    panel_empty_shortcuts: 'Configure shortcuts',
    panel_entry_no_label: '(no label)',
    panel_badge_not_fetched: 'Not fetched',
    panel_badge_permission_required: 'Host permission required',
    panel_badge_missing_result: 'Count was not returned for this item',
    panel_badge_fetch_failed: 'Failed to fetch count',
    panel_notice_permission_missing_hosts: 'Permission is missing for hosts: {hosts}',
    panel_notice_counts_failed: 'Some hosts failed to fetch counts. Check the log for details.',
    panel_notice_shortcut_url_missing: 'Shortcut URL is not configured',
    panel_notice_open_kintone_first: 'Open a kintone tab first',
    panel_notice_parse_url_failed: 'Could not parse the URL',
    panel_notice_already_registered: 'This view is already registered',
    panel_notice_added_fetching: 'Added to watchlist. Fetching counts...',
    panel_notice_add_failed: 'Failed to add watchlist item',
    panel_notice_url_required: 'URL is required',
    panel_notice_kintone_url_required: 'Please enter a kintone URL',
    panel_notice_updated: 'Updated',
    panel_notice_icon_updated: 'Icon updated',
    panel_notice_category_updated: 'Category updated',
    panel_notice_initialization_failed: 'Initialization failed',
    panel_prompt_edit_label: 'Edit label',
    panel_prompt_edit_url: 'Edit URL',
    panel_prompt_icon: 'Enter icon name\nChoices: {icons}',
    panel_prompt_category: 'Enter category (empty = Other)',
    panel_confirm_delete_shortcut: 'Delete this shortcut?',
    panel_meta_view_prefix: 'view',
    panel_search_panel_placeholder: 'Search panel'
  }
};

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
  toggleRecentRecords: null,
  appSearchToggle: null,
  appSearchPopover: null,
  appSearchInput: null,
  appSearchResults: null,
  listsScroll: doc.getElementById('listsScroll'),
  collapsedTray: null,
  collapsedTrayRow: null,
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
  recordPinsCollapsed: false,
  recentRecordsCollapsed: false,
  appCatalog: [],
  appCatalogSeq: 0,
  appCatalogError: false,
  appSearchOpen: false,
  appSearchQuery: '',
  appSearchMatches: [],
  appSearchActiveIndex: -1,
  shortcutSearchOpenMode: SHORTCUT_SEARCH_OPEN_MODE_CURRENT_TAB,
  appSearchToggleHandler: null,
  appSearchInputHandler: null,
  appSearchInputKeydownHandler: null,
  appSearchOutsideHandler: null,
  pinListObserver: null,
  uiLanguageSetting: DEFAULT_UI_LANGUAGE,
  currentLang: DEFAULT_LANGUAGE
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
const DEFAULT_CATEGORY = 'Other';
const ICON_OPTIONS = [
  'clipboard', 'file-text', 'package', 'box', 'truck', 'factory', 'wrench', 'calendar',
  'list-checks', 'search', 'filter', 'chart-bar', 'receipt', 'users', 'settings', 'bookmark', 'star', 'history'
];
const DEFAULT_ICON_COLOR = 'gray';
const ICON_COLOR_OPTIONS = ['gray', 'blue', 'green', 'orange', 'red', 'purple'];
const APP_CANDIDATE_LIMIT = 10;
const SHORTCUT_MAX_VISIBLE = 16;

function normalizeUiLanguage(raw) {
  const value = String(raw || '').trim().toLowerCase();
  if (!value) return DEFAULT_LANGUAGE;
  if (value === 'ja' || value.startsWith('ja-')) return 'ja';
  return 'en';
}

function normalizeUiLanguageSetting(raw) {
  const value = String(raw || '').trim().toLowerCase();
  if (UI_LANGUAGE_VALUES.includes(value)) return value;
  return DEFAULT_UI_LANGUAGE;
}

function getBrowserUiLanguage() {
  try {
    const browserLang = navigator.language || (Array.isArray(navigator.languages) ? navigator.languages[0] : '');
    return normalizeUiLanguage(browserLang);
  } catch (_err) {
    return DEFAULT_LANGUAGE;
  }
}

function resolveEffectiveUiLanguage(setting) {
  if (setting === 'ja' || setting === 'en') return setting;
  return getBrowserUiLanguage();
}

function t(key, vars) {
  const base = I18N_MESSAGES[state.currentLang]?.[key]
    ?? I18N_MESSAGES.ja?.[key]
    ?? key;
  if (!vars || typeof vars !== 'object') return base;
  return Object.keys(vars).reduce((text, name) => {
    return text.replaceAll(`{${name}}`, String(vars[name]));
  }, base);
}

function applyI18n(root = doc) {
  if (!root) return;
  root.querySelectorAll('[data-i18n]').forEach((el) => {
    const key = el.getAttribute('data-i18n');
    if (!key) return;
    el.textContent = t(key);
  });
  root.querySelectorAll('[data-i18n-title]').forEach((el) => {
    const key = el.getAttribute('data-i18n-title');
    if (!key) return;
    el.setAttribute('title', t(key));
  });
  root.querySelectorAll('[data-i18n-placeholder]').forEach((el) => {
    const key = el.getAttribute('data-i18n-placeholder');
    if (!key) return;
    el.setAttribute('placeholder', t(key));
  });
  root.querySelectorAll('[data-i18n-aria-label]').forEach((el) => {
    const key = el.getAttribute('data-i18n-aria-label');
    if (!key) return;
    el.setAttribute('aria-label', t(key));
  });
}

function getSortLocale() {
  return state.currentLang === 'en' ? 'en' : DEFAULT_LOCALE;
}

async function initializeI18n() {
  let setting = DEFAULT_UI_LANGUAGE;
  try {
    const stored = await chrome.storage.local.get(UI_LANGUAGE_KEY);
    setting = normalizeUiLanguageSetting(stored?.[UI_LANGUAGE_KEY]);
  } catch (_err) {
    setting = DEFAULT_UI_LANGUAGE;
  }
  await applyUiLanguageSetting(setting, { persist: false, rerender: false });
}

async function applyUiLanguageSetting(settingValue, { persist = false, rerender = true } = {}) {
  const setting = normalizeUiLanguageSetting(settingValue);
  const nextLang = resolveEffectiveUiLanguage(setting);
  state.uiLanguageSetting = setting;
  state.currentLang = nextLang;
  doc.documentElement.lang = nextLang;
  applyI18n(doc);
  if (persist) {
    try {
      await chrome.storage.local.set({ [UI_LANGUAGE_KEY]: setting });
    } catch (_err) {
      // ignore
    }
  }
  if (rerender) {
    rerenderLocalizedUi();
  }
}

function rerenderLocalizedUi() {
  if (els.filter) {
    els.filter.placeholder = t('panel_search_placeholder');
  }
  if (els.collapsedTray) {
    els.collapsedTray.setAttribute('aria-label', t('panel_collapsed_sections_label'));
  }
  if (els.appSearchToggle) {
    els.appSearchToggle.title = t('panel_action_app_search');
    els.appSearchToggle.setAttribute('aria-label', t('panel_action_app_search'));
  }
  if (els.appSearchInput) {
    els.appSearchInput.placeholder = t('panel_app_search_placeholder');
  }
  setFilterPanelVisible(state.filterPanelVisible);
  applyShortcutVisibility(shortcutState.visible);
  setWatchlistCollapsed(state.pinsOnly, { persist: false });
  setRecordPinsCollapsed(state.recordPinsCollapsed, { persist: false });
  setRecentRecordsCollapsed(state.recentRecordsCollapsed, { persist: false });
  renderLists();
  renderShortcuts();
  renderRecentRecords();
  if (state.appSearchOpen) {
    refreshAppSearchMatches();
  } else {
    renderAppSearchResults();
  }
}

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

function setNoticeKey(key, vars) {
  setNotice(t(key, vars));
}

function setFilterPanelVisible(visible, { focus = false } = {}) {
  state.filterPanelVisible = Boolean(visible);
  if (els.filterPanel) {
    els.filterPanel.classList.toggle('hidden', !state.filterPanelVisible);
  }
  if (els.toggleFilterPanel) {
    els.toggleFilterPanel.setAttribute('aria-pressed', state.filterPanelVisible ? 'true' : 'false');
    const label = state.filterPanelVisible ? t('panel_action_toggle_search_hide') : t('panel_action_toggle_search_show');
    els.toggleFilterPanel.title = label;
    els.toggleFilterPanel.setAttribute('aria-label', label);
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

function ensureRecentCollapseControl() {
  if (!els.recentSection) return;
  const head = els.recentSection.querySelector('.section-head');
  if (!head) return;

  let actions = head.querySelector('.section-actions');
  if (!actions) {
    actions = doc.createElement('div');
    actions.className = 'section-actions section-actions-icons';
    head.appendChild(actions);
  } else {
    actions.classList.add('section-actions-icons');
  }

  if (els.recentClear && !actions.contains(els.recentClear)) {
    actions.appendChild(els.recentClear);
  }

  if (!els.toggleRecentRecords) {
    const toggle = doc.createElement('button');
    toggle.id = 'toggleRecentRecords';
    toggle.type = 'button';
    toggle.className = 'icon-btn small collapse-btn';
    toggle.setAttribute('aria-pressed', 'false');
    actions.appendChild(toggle);
    els.toggleRecentRecords = toggle;
  }
}

function ensureCollapsedSectionsTray() {
  if (els.collapsedTray && els.collapsedTrayRow) return;
  if (!els.shortcutPanel) return;

  const tray = doc.createElement('section');
  tray.id = 'collapsedSectionsTray';
  tray.className = 'collapsed-sections-tray hidden';
  tray.setAttribute('aria-label', t('panel_collapsed_sections_label'));

  const row = doc.createElement('div');
  row.id = 'collapsedSectionsRow';
  row.className = 'collapsed-sections-row';
  tray.appendChild(row);

  if (els.notice?.parentElement) {
    els.notice.parentElement.insertBefore(tray, els.notice);
  } else if (els.shortcutPanel.parentElement) {
    els.shortcutPanel.parentElement.insertBefore(tray, els.shortcutPanel.nextSibling);
  }

  els.collapsedTray = tray;
  els.collapsedTrayRow = row;
}

function openCollapsedTrayTooltip(btn) {
  if (!btn) return;
  const tooltip = btn.querySelector('.collapsed-tray-tooltip');
  if (!tooltip || !tooltip.textContent) return;
  btn.classList.add('tray-tooltip-open');
  btn.classList.remove('tray-tooltip-align-left', 'tray-tooltip-align-right');
  const panelRect = (els.collapsedTray || els.shortcutPanel || doc.body).getBoundingClientRect();
  const tooltipRect = tooltip.getBoundingClientRect();
  if (tooltipRect.left < panelRect.left + 4) {
    btn.classList.add('tray-tooltip-align-left');
    return;
  }
  if (tooltipRect.right > panelRect.right - 4) {
    btn.classList.add('tray-tooltip-align-right');
  }
}

function closeCollapsedTrayTooltip(btn) {
  if (!btn) return;
  btn.classList.remove('tray-tooltip-open', 'tray-tooltip-align-left', 'tray-tooltip-align-right');
}

function createCollapsedTrayItem({ id, icon, label, count, onClick }) {
  const btn = doc.createElement('button');
  btn.type = 'button';
  btn.className = 'collapsed-tray-item';
  btn.dataset.section = id;
  btn.setAttribute('aria-label', label);
  btn.title = label;
  btn.addEventListener('click', onClick);
  btn.addEventListener('mouseenter', () => openCollapsedTrayTooltip(btn));
  btn.addEventListener('focus', () => openCollapsedTrayTooltip(btn));
  btn.addEventListener('mouseleave', () => closeCollapsedTrayTooltip(btn));
  btn.addEventListener('blur', () => closeCollapsedTrayTooltip(btn));

  const iconEl = doc.createElement('span');
  iconEl.className = 'collapsed-tray-icon lc';
  iconEl.dataset.icon = icon;
  iconEl.setAttribute('aria-hidden', 'true');

  const tooltip = doc.createElement('span');
  tooltip.className = 'collapsed-tray-tooltip';
  tooltip.textContent = label;
  tooltip.setAttribute('aria-hidden', 'true');

  btn.appendChild(iconEl);
  if (Number(count) > 0) {
    const badge = doc.createElement('span');
    badge.className = 'collapsed-tray-badge';
    badge.textContent = String(count);
    badge.setAttribute('aria-hidden', 'true');
    btn.appendChild(badge);
  }
  btn.appendChild(tooltip);
  return btn;
}

function renderCollapsedSectionsTray() {
  ensureCollapsedSectionsTray();
  if (!els.collapsedTray || !els.collapsedTrayRow) return;

  const items = [];
  if (state.pinsOnly) {
    items.push({
      id: 'watchlist',
      icon: 'bookmark',
      label: t('panel_collapsed_watchlist'),
      count: state.favorites.length,
      onClick: () => setWatchlistCollapsed(false)
    });
  }
  if (state.recordPinsCollapsed && recordPinState.visible) {
    const count = els.pinList ? els.pinList.querySelectorAll('li').length : 0;
    items.push({
      id: 'record-pins',
      icon: 'pin',
      label: t('panel_collapsed_record_pins'),
      count,
      onClick: () => setRecordPinsCollapsed(false)
    });
  }
  if (state.recentRecordsCollapsed && state.recentRecords.length > 0) {
    items.push({
      id: 'recent-records',
      icon: 'history',
      label: t('panel_collapsed_recent_records'),
      count: state.recentRecords.length,
      onClick: () => setRecentRecordsCollapsed(false)
    });
  }

  els.collapsedTrayRow.textContent = '';
  if (!items.length) {
    els.collapsedTray.classList.add('hidden');
    return;
  }
  items.forEach((item) => {
    els.collapsedTrayRow.appendChild(createCollapsedTrayItem(item));
  });
  els.collapsedTray.classList.remove('hidden');
  renderLucideIcons(els.collapsedTray);
}

function setRecordPinsCollapsed(collapsed, { persist = true } = {}) {
  state.recordPinsCollapsed = Boolean(collapsed);
  if (els.pinList) {
    els.pinList.classList.toggle('hidden', state.recordPinsCollapsed);
  }
  if (els.recordPinSection) {
    els.recordPinSection.classList.toggle('tray-collapsed', state.recordPinsCollapsed);
  }
  if (els.toggleRecordPins) {
    const isCollapsed = state.recordPinsCollapsed;
    els.toggleRecordPins.textContent = isCollapsed ? '\u25B6' : '\u25BC';
    els.toggleRecordPins.setAttribute('aria-pressed', isCollapsed ? 'true' : 'false');
    const label = isCollapsed ? t('panel_action_toggle_record_pins_expand') : t('panel_action_toggle_record_pins_collapse');
    els.toggleRecordPins.title = label;
    els.toggleRecordPins.setAttribute('aria-label', label);
  }
  if (persist) {
    persistCollapsedSectionsState().catch(() => {});
  }
  renderCollapsedSectionsTray();
}

function setWatchlistCollapsed(collapsed, { persist = true } = {}) {
  state.pinsOnly = Boolean(collapsed);
  if (els.togglePinsOnly) {
    els.togglePinsOnly.checked = state.pinsOnly;
    els.togglePinsOnly.setAttribute('aria-pressed', state.pinsOnly ? 'true' : 'false');
  }
  if (els.favCategories) {
    els.favCategories.classList.toggle('is-collapsed', state.pinsOnly);
  }
  if (els.favSection) {
    els.favSection.classList.toggle('tray-collapsed', state.pinsOnly);
  }
  if (els.togglePinsOnlyIcon) {
    const isCollapsed = state.pinsOnly;
    els.togglePinsOnlyIcon.textContent = isCollapsed ? '\u25B6' : '\u25BC';
    const label = isCollapsed ? t('panel_action_toggle_watchlist_expand') : t('panel_action_toggle_watchlist_collapse');
    els.togglePinsOnlyIcon.title = label;
    els.togglePinsOnlyIcon.setAttribute('aria-label', label);
  }
  if (persist) {
    persistCollapsedSectionsState().catch(() => {});
  }
  renderCollapsedSectionsTray();
}

function setRecentRecordsCollapsed(collapsed, { persist = true } = {}) {
  state.recentRecordsCollapsed = Boolean(collapsed);
  if (els.recentSection) {
    els.recentSection.classList.toggle('tray-collapsed', state.recentRecordsCollapsed);
  }
  if (els.toggleRecentRecords) {
    const isCollapsed = state.recentRecordsCollapsed;
    els.toggleRecentRecords.textContent = isCollapsed ? '\u25B6' : '\u25BC';
    els.toggleRecentRecords.setAttribute('aria-pressed', isCollapsed ? 'true' : 'false');
    const label = isCollapsed ? t('panel_action_toggle_recent_records_expand') : t('panel_action_toggle_recent_records_collapse');
    els.toggleRecentRecords.title = label;
    els.toggleRecentRecords.setAttribute('aria-label', label);
  }
  if (persist) {
    persistCollapsedSectionsState().catch(() => {});
  }
  renderCollapsedSectionsTray();
}

function normalizeShortcutSearchOpenMode(value) {
  if (value === SHORTCUT_SEARCH_OPEN_MODE_NEW_TAB) return SHORTCUT_SEARCH_OPEN_MODE_NEW_TAB;
  return SHORTCUT_SEARCH_OPEN_MODE_CURRENT_TAB;
}

function normalizeCollapsedSections(value) {
  const source = value && typeof value === 'object' ? value : {};
  return {
    watchlist: Boolean(source.watchlist),
    recordPins: Boolean(source.recordPins),
    recentRecords: Boolean(source.recentRecords)
  };
}

async function loadCollapsedSectionsState() {
  try {
    const stored = await chrome.storage.local.get(COLLAPSED_SECTIONS_KEY);
    const normalized = normalizeCollapsedSections(stored?.[COLLAPSED_SECTIONS_KEY]);
    state.pinsOnly = normalized.watchlist;
    state.recordPinsCollapsed = normalized.recordPins;
    state.recentRecordsCollapsed = normalized.recentRecords;
  } catch (_err) {
    state.pinsOnly = false;
    state.recordPinsCollapsed = false;
    state.recentRecordsCollapsed = false;
  }
}

async function persistCollapsedSectionsState() {
  const payload = {
    watchlist: Boolean(state.pinsOnly),
    recordPins: Boolean(state.recordPinsCollapsed),
    recentRecords: Boolean(state.recentRecordsCollapsed)
  };
  try {
    await chrome.storage.local.set({ [COLLAPSED_SECTIONS_KEY]: payload });
  } catch (_err) {
    // ignore persistence errors
  }
}

async function loadShortcutSearchOpenMode() {
  try {
    const stored = await chrome.storage.sync.get(SHORTCUT_SEARCH_OPEN_MODE_KEY);
    state.shortcutSearchOpenMode = normalizeShortcutSearchOpenMode(stored?.[SHORTCUT_SEARCH_OPEN_MODE_KEY]);
  } catch (_err) {
    state.shortcutSearchOpenMode = SHORTCUT_SEARCH_OPEN_MODE_CURRENT_TAB;
  }
}

async function openUrlByMode(url, mode) {
  if (!url) return;
  const normalizedMode = normalizeShortcutSearchOpenMode(mode);
  if (normalizedMode === SHORTCUT_SEARCH_OPEN_MODE_NEW_TAB) {
    await chrome.tabs.create({ url, active: true });
    return;
  }
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  const activeTab = tabs && tabs[0] ? tabs[0] : null;
  if (activeTab?.id) {
    await chrome.tabs.update(activeTab.id, { url });
    return;
  }
  await chrome.tabs.create({ url, active: true });
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
  return appId ? `${t('panel_recent_app_prefix')} ${appId}` : t('panel_recent_app_fallback');
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

function normalizeHostOrigin(host) {
  const raw = String(host || '').trim();
  if (!raw) return '';
  try {
    return new URL(raw).origin.replace(/\/$/, '');
  } catch (_err) {
    return raw.replace(/\/$/, '');
  }
}

function buildFavoriteAppKey(host, appId) {
  const safeHost = normalizeHostOrigin(host);
  const safeAppId = String(appId || '').trim();
  if (!safeHost || !safeAppId) return '';
  return `${safeHost}|${safeAppId}`;
}

function appSearchText(value) {
  return String(value || '').trim().toLowerCase();
}

function collectFavoriteHostAppIds(host) {
  const safeHost = normalizeHostOrigin(host);
  return state.favorites
    .filter((item) => normalizeHostOrigin(item.host) === safeHost)
    .map((item) => String(item.appId || '').trim())
    .filter((appId) => /^\d+$/.test(appId));
}

function buildFavoriteAppSet() {
  const set = new Set();
  state.favorites.forEach((item) => {
    const key = buildFavoriteAppKey(item.host, item.appId);
    if (key) set.add(key);
  });
  return set;
}

async function loadHostAppNameMap(host) {
  const safeHost = normalizeHostOrigin(host);
  if (!safeHost) return { map: {}, ok: false };
  let cached = { map: {}, fresh: false };
  try {
    cached = await loadAppNameMap(safeHost);
  } catch (_err) {
    cached = { map: {}, fresh: false };
  }
  const cachedMap = cached?.map && typeof cached.map === 'object' ? cached.map : {};
  if (cached.fresh && Object.keys(cachedMap).length) {
    return { map: cachedMap, ok: true };
  }
  try {
    const res = await chrome.runtime.sendMessage({
      type: 'GET_APPS_MAP',
      host: safeHost,
      appIds: collectFavoriteHostAppIds(safeHost)
    });
    if (res?.ok && res?.map && typeof res.map === 'object') {
      return { map: res.map, ok: true };
    }
  } catch (_err) {
    // ignore
  }
  return { map: cachedMap, ok: false };
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
    renderCollapsedSectionsTray();
    applySearchVisibility();
    return;
  }
  list.forEach((item) => {
    const li = doc.createElement('li');
    li.className = 'recent-list-item';
    const btn = doc.createElement('button');
    btn.type = 'button';
    btn.className = 'recent-item';
    const safeHost = normalizeRecentHost(item.host);
    const safeAppId = String(item.appId || '').trim();
    const safeRecordId = String(item.recordId || '').trim();
    const appLabel = getRecentAppLabel(item);
    btn.dataset.host = safeHost;
    btn.dataset.appId = safeAppId;
    btn.dataset.recordId = safeRecordId;
    btn.dataset.search = [
      appLabel,
      safeRecordId,
      item.url || '',
      safeHost
    ].join('\n').toLowerCase();
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
  setRecentRecordsCollapsed(state.recentRecordsCollapsed, { persist: false });
  renderLucideIcons(els.recentSection);
  applySearchVisibility();
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

function searchAppsByName(query, limit = APP_CANDIDATE_LIMIT) {
  const keyword = appSearchText(query);
  if (!keyword) return [];
  const prefix = [];
  const partial = [];
  state.appCatalog.forEach((item) => {
    const nameText = appSearchText(item.name);
    if (!nameText) return;
    if (nameText.startsWith(keyword)) {
      prefix.push(item);
      return;
    }
    if (nameText.includes(keyword)) {
      partial.push(item);
    }
  });
  const sorter = (a, b) => {
    const nameA = String(a.name || '');
    const nameB = String(b.name || '');
    return nameA.localeCompare(nameB, getSortLocale());
  };
  return [...prefix.sort(sorter), ...partial.sort(sorter)].slice(0, Math.max(1, Number(limit) || APP_CANDIDATE_LIMIT));
}

function normalizeAppSearchHost(host) {
  return String(host || '')
    .trim()
    .replace(/^https?:\/\//i, '')
    .replace(/\/+$/, '');
}

function buildAppSearchUrl(appId, hostHint = '') {
  const safeAppId = String(appId || '').trim();
  if (!/^\d+$/.test(safeAppId)) return '';

  let host = normalizeAppSearchHost(hostHint);
  if (!host && state.appSearchMatches.length) {
    host = normalizeAppSearchHost(state.appSearchMatches[0]?.host || '');
  }
  if (!host && state.appCatalog.length) {
    host = normalizeAppSearchHost(state.appCatalog[0]?.host || '');
  }
  if (!host) {
    const activeHost = normalizeAppSearchHost(state.favorites[0]?.host || '');
    host = activeHost || '';
  }
  const normalizedHost = host;
  if (!normalizedHost) return '';
  return `https://${normalizedHost}/k/${encodeURIComponent(safeAppId)}/`;
}

async function openAppSearchCandidate(appId, hostHint = '') {
  try {
    const safeAppId = String(appId || '').trim();
    if (!/^\d+$/.test(safeAppId)) return;

    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    const activeTab = tabs && tabs[0] ? tabs[0] : null;

    let normalizedHost = normalizeAppSearchHost(hostHint);
    const activeUrl = String(activeTab?.url || '').trim();
    if (!normalizedHost && activeUrl && isKintoneUrl(activeUrl)) {
      const parsed = parseKintoneUrl(activeUrl);
      normalizedHost = normalizeAppSearchHost(parsed?.host || '');
    }
    if (!normalizedHost) {
      normalizedHost = normalizeAppSearchHost(location.host);
    }
    if (!normalizedHost) {
      console.warn('No host resolved for app search navigation');
      return;
    }

    const url = buildAppSearchUrl(safeAppId, normalizedHost);
    if (!url) return;

    await openUrlByMode(url, state.shortcutSearchOpenMode);
  } catch (err) {
    console.error('Failed to open app search candidate', err);
  }
}

function renderAppSearchResults() {
  if (!els.appSearchResults) return;
  els.appSearchResults.textContent = '';
  const query = appSearchText(state.appSearchQuery);
  if (!query) return;

  if (!state.appSearchMatches.length) {
    const empty = doc.createElement('li');
    empty.className = 'kp-app-search-empty';
    empty.textContent = state.appCatalogError && !state.appCatalog.length
      ? t('panel_app_search_load_failed')
      : t('panel_app_search_no_results');
    els.appSearchResults.appendChild(empty);
    return;
  }

  state.appSearchMatches.forEach((item, index) => {
    const li = doc.createElement('li');
    li.className = 'kp-app-search-row';

    const btn = doc.createElement('button');
    btn.type = 'button';
    btn.className = 'kp-app-search-item';
    if (index === state.appSearchActiveIndex) {
      btn.classList.add('is-active');
    }
    btn.title = item.url || '';
    btn.dataset.index = String(index);
    btn.dataset.appId = String(item.appId || '');
    btn.dataset.host = normalizeAppSearchHost(item.host || '');
    btn.addEventListener('mouseenter', () => {
      state.appSearchActiveIndex = index;
      updateAppSearchActiveDom();
    });
    btn.addEventListener('click', async () => {
      const appId = btn.dataset.appId || item.appId;
      const host = btn.dataset.host || item.host || '';
      await openAppSearchCandidate(appId, host);
      setAppSearchPopoverVisible(false);
    });

    const name = doc.createElement('div');
    name.className = 'kp-app-search-name';
    name.textContent = item.name;

    const meta = doc.createElement('div');
    meta.className = 'kp-app-search-meta';
    meta.textContent = `app:${item.appId}`;

    btn.appendChild(name);
    btn.appendChild(meta);
    li.appendChild(btn);
    els.appSearchResults.appendChild(li);
  });
}

function updateAppSearchActiveDom() {
  if (!els.appSearchResults) return;
  const items = els.appSearchResults.querySelectorAll('.kp-app-search-item');
  items.forEach((node, i) => {
    node.classList.toggle('is-active', i === state.appSearchActiveIndex);
  });
}

function refreshAppSearchMatches() {
  const query = appSearchText(state.appSearchQuery);
  if (!query) {
    state.appSearchMatches = [];
    state.appSearchActiveIndex = -1;
    renderAppSearchResults();
    return;
  }
  state.appSearchMatches = searchAppsByName(query, 8);
  state.appSearchActiveIndex = state.appSearchMatches.length ? 0 : -1;
  renderAppSearchResults();
}

function setAppSearchPopoverVisible(visible) {
  state.appSearchOpen = Boolean(visible);
  if (els.appSearchPopover) {
    els.appSearchPopover.classList.toggle('hidden', !state.appSearchOpen);
  }
  if (els.appSearchToggle) {
    els.appSearchToggle.setAttribute('aria-pressed', state.appSearchOpen ? 'true' : 'false');
  }
  if (!state.appSearchOpen) {
    state.appSearchQuery = '';
    state.appSearchMatches = [];
    state.appSearchActiveIndex = -1;
    if (els.appSearchInput) els.appSearchInput.value = '';
    renderAppSearchResults();
    return;
  }
  refreshAppCatalog().catch(() => {});
  refreshAppSearchMatches();
  if (els.appSearchInput) {
    setTimeout(() => {
      try {
        els.appSearchInput.focus();
      } catch (_err) {
        // ignore
      }
    }, 0);
  }
}

function moveAppSearchSelection(delta) {
  const len = state.appSearchMatches.length;
  if (!len) return;
  const base = state.appSearchActiveIndex >= 0 ? state.appSearchActiveIndex : 0;
  const next = (base + delta + len) % len;
  state.appSearchActiveIndex = next;
  updateAppSearchActiveDom();
  const active = els.appSearchResults?.querySelector('.kp-app-search-item.is-active');
  if (active) {
    active.scrollIntoView({ block: 'nearest' });
  }
}

function openActiveOrFirstAppSearchResult() {
  if (!state.appSearchMatches.length) return;
  const index = state.appSearchActiveIndex >= 0 ? state.appSearchActiveIndex : 0;
  const item = state.appSearchMatches[index];
  if (!item) return;
  openAppSearchCandidate(item.appId, item.host).finally(() => {
    setAppSearchPopoverVisible(false);
  });
}

function handleAppSearchInputKeydown(event) {
  if (event.key === 'Enter' && isImeComposing(event)) {
    return;
  }
  if (event.key === 'ArrowDown') {
    event.preventDefault();
    moveAppSearchSelection(1);
    return;
  }
  if (event.key === 'ArrowUp') {
    event.preventDefault();
    moveAppSearchSelection(-1);
    return;
  }
  if (event.key === 'Enter') {
    event.preventDefault();
    openActiveOrFirstAppSearchResult();
    return;
  }
  if (event.key === 'Escape') {
    event.preventDefault();
    setAppSearchPopoverVisible(false);
  }
}

function isEditableTarget(target) {
  if (!target) return false;
  if (target instanceof HTMLTextAreaElement) return true;
  if (target instanceof HTMLInputElement) return true;
  if (target.isContentEditable) return true;
  if (typeof target.closest === 'function') {
    return Boolean(target.closest('textarea, input, [contenteditable="true"]'));
  }
  return false;
}

function isImeComposing(event) {
  return Boolean(event?.isComposing) || Number(event?.keyCode) === 229;
}

function ensureAppSearchUi() {
  const actions = els.shortcutPanel?.querySelector('.shortcut-head .section-actions');
  if (!actions || !els.shortcutPanel) return;

  if (!els.appSearchToggle) {
    const toggle = doc.createElement('button');
    toggle.id = 'kp-app-search-toggle';
    toggle.type = 'button';
    toggle.className = 'icon-btn small';
    toggle.setAttribute('aria-label', t('panel_action_app_search'));
    toggle.setAttribute('aria-pressed', 'false');
    toggle.title = t('panel_action_app_search');

    const icon = doc.createElement('span');
    icon.className = 'lc';
    icon.dataset.icon = 'search';
    icon.setAttribute('aria-hidden', 'true');
    toggle.appendChild(icon);

    const collapseButton = actions.querySelector('#toggleShortcuts');
    if (collapseButton) {
      actions.insertBefore(toggle, collapseButton);
    } else {
      actions.appendChild(toggle);
    }
    els.appSearchToggle = toggle;
    renderLucideIcons(toggle);
  }

  if (!els.appSearchPopover) {
    const popover = doc.createElement('div');
    popover.id = 'kp-app-search-popover';
    popover.className = 'kp-app-search-popover hidden';

    const head = doc.createElement('div');
    head.className = 'kp-app-search-head';

    const input = doc.createElement('input');
    input.id = 'kp-app-search-input';
    input.type = 'text';
    input.placeholder = t('panel_app_search_placeholder');
    input.autocomplete = 'off';

    const results = doc.createElement('ul');
    results.id = 'kp-app-search-results';
    results.className = 'kp-app-search-results';
    results.setAttribute('aria-live', 'polite');

    head.appendChild(input);
    popover.appendChild(head);
    popover.appendChild(results);
    els.shortcutPanel.appendChild(popover);

    els.appSearchPopover = popover;
    els.appSearchInput = input;
    els.appSearchResults = results;
  }

  if (!state.appSearchToggleHandler && els.appSearchToggle) {
    state.appSearchToggleHandler = () => {
      setAppSearchPopoverVisible(!state.appSearchOpen);
    };
    els.appSearchToggle.addEventListener('click', state.appSearchToggleHandler);
  }
  if (!state.appSearchInputHandler && els.appSearchInput) {
    state.appSearchInputHandler = (event) => {
      state.appSearchQuery = String(event.target?.value || '');
      refreshAppSearchMatches();
    };
    els.appSearchInput.addEventListener('input', state.appSearchInputHandler);
  }
  if (!state.appSearchInputKeydownHandler && els.appSearchInput) {
    state.appSearchInputKeydownHandler = handleAppSearchInputKeydown;
    els.appSearchInput.addEventListener('keydown', state.appSearchInputKeydownHandler);
  }
  if (!state.appSearchOutsideHandler) {
    state.appSearchOutsideHandler = (event) => {
      if (!state.appSearchOpen) return;
      const target = event.target;
      if (els.appSearchPopover?.contains(target) || els.appSearchToggle?.contains(target)) return;
      setAppSearchPopoverVisible(false);
    };
    document.addEventListener('mousedown', state.appSearchOutsideHandler);
  }
}

async function refreshAppCatalog() {
  const seq = ++state.appCatalogSeq;
  const hostSet = new Set(
    state.favorites
      .map((item) => normalizeHostOrigin(item.host))
      .filter(Boolean)
  );
  try {
    const tab = await getActiveTab();
    const tabUrl = String(tab?.url || '');
    if (tabUrl && isKintoneUrl(tabUrl)) {
      const parsed = parseKintoneUrl(tabUrl);
      const host = normalizeHostOrigin(parsed?.host || '');
      if (host) hostSet.add(host);
    }
  } catch (_err) {
    // ignore
  }
  const hostList = Array.from(hostSet);
  if (!hostList.length) {
    state.appCatalog = [];
    state.appCatalogError = false;
    if (state.appSearchOpen) {
      refreshAppSearchMatches();
    }
    return;
  }

  const favoriteSet = buildFavoriteAppSet();
  const nextCatalog = [];
  let hasFailure = false;

  for (const host of hostList) {
    if (seq !== state.appCatalogSeq) return;
    const { map, ok } = await loadHostAppNameMap(host);
    if (!ok) hasFailure = true;
    Object.entries(map || {}).forEach(([appIdRaw, nameRaw]) => {
      const appId = String(appIdRaw || '').trim();
      const name = String(nameRaw || '').trim();
      if (!appId || !name) return;
      const key = buildFavoriteAppKey(host, appId);
      nextCatalog.push({
        host,
        appId,
        name,
        nameLower: appSearchText(name),
        url: `${host}/k/${encodeURIComponent(appId)}/`,
        isFavorite: key ? favoriteSet.has(key) : false
      });
    });
  }

  if (seq !== state.appCatalogSeq) return;

  nextCatalog.sort((a, b) => {
    const byName = String(a.name || '').localeCompare(String(b.name || ''), getSortLocale());
    if (byName !== 0) return byName;
    return String(a.appId || '').localeCompare(String(b.appId || ''), getSortLocale());
  });

  state.appCatalog = nextCatalog;
  state.appCatalogError = hasFailure && nextCatalog.length === 0;
  if (state.appSearchOpen) {
    refreshAppSearchMatches();
  }
}

function normalizeFavorites(list) {
  return sortFavorites(
    (list || []).map((item, index) => ({
      id: item.id || createId(),
      label: item.label || t('panel_entry_no_label'),
      url: item.url || '',
      host: item.host || '',
      appId: item.appId || '',
      appName: item.appName || '',
      viewIdOrName: item.viewIdOrName || '',
      viewName: item.viewName || '',
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
  // #filter is local filter only for sidepanel items.
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
    els.pinSection.classList.toggle('hidden', !hasPinned);
  }

  const pinCards = Array.from(doc.querySelectorAll('#pinList .pin-card'));
  pinCards.forEach((cardEl) => {
    const title = cardEl.querySelector('.pin-title-input')?.value || '';
    const note = cardEl.querySelector('.pin-note-input')?.value || '';
    const recordTitle = cardEl.querySelector('.pin-record-title')?.textContent || '';
    const meta = cardEl.querySelector('.pin-meta')?.textContent || '';
    const url = cardEl.querySelector('.record-pin-open')?.getAttribute('href') || '';
    let host = '';
    try {
      host = new URL(url).origin;
    } catch (_err) {
      host = '';
    }
    const haystack = [title, note, recordTitle, meta, url, host].join('\n').toLowerCase();
    const matched = !keyword || haystack.includes(keyword);
    cardEl.classList.toggle('pin-filter-hidden', !matched);
  });

  const recentItems = Array.from(doc.querySelectorAll('#recentList .recent-item'));
  recentItems.forEach((itemEl) => {
    const haystack = itemEl.dataset.search || itemEl.textContent?.toLowerCase() || '';
    const matched = !keyword || haystack.includes(keyword);
    const row = itemEl.closest('.recent-list-item');
    if (row) {
      row.classList.toggle('recent-filter-hidden', !matched);
    } else {
      itemEl.classList.toggle('recent-filter-hidden', !matched);
    }
  });
}

function isSelectableEntry(entryEl) {
  if (!entryEl?.dataset?.id) return false;
  if (entryEl.classList.contains('entry-filter-hidden')) return false;
  if (entryEl.closest('.category-block.hidden-by-filter')) return false;
  if (entryEl.closest('.hidden')) return false;
  if (entryEl.closest('.tray-collapsed')) return false;
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
  setWatchlistCollapsed(state.pinsOnly, { persist: false });
  applySearchVisibility();
  renderCollapsedSectionsTray();
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
      ? t('panel_empty_pinned_no_match')
      : t('panel_empty_pinned');
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
      ? t('panel_empty_favorites_no_match')
      : t('panel_empty_favorites');
    container.appendChild(message);
    return [];
  }
  const grouped = groupByCategory(items);
  const order = Array.from(grouped.keys()).sort((a, b) => {
    if (a === DEFAULT_CATEGORY) return 1;
    if (b === DEFAULT_CATEGORY) return -1;
    return a.localeCompare(b, getSortLocale());
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
  button.addEventListener('click', () => {
    const useSearchMode = Boolean(state.filterText.trim());
    openEntry(entry, { useSearchMode });
  });
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
  title.textContent = entry.label || t('panel_entry_no_label');
  const sub = doc.createElement('div');
  sub.className = 'entry-sub';
  const view = entry.viewIdOrName ? `${t('panel_meta_view_prefix')}:${entry.viewIdOrName}` : '';
  sub.textContent = view;
  sub.classList.toggle('hidden', !view);
  li.dataset.search = [
    entry.label || '',
    entry.url || '',
    entry.host || '',
    entry.appId || '',
    entry.appName || '',
    categoryLabel,
    entry.viewIdOrName || '',
    entry.viewName || '',
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
  if (entry) openEntry(entry, { useSearchMode: true });
}

async function openEntry(entry, options = {}) {
  if (!entry?.url) return;
  const useSearchMode = Boolean(options.useSearchMode);
  const openMode = useSearchMode ? state.shortcutSearchOpenMode : SHORTCUT_SEARCH_OPEN_MODE_NEW_TAB;
  try {
    await openUrlByMode(entry.url, openMode);
  } catch (error) {
    console.error('Failed to open entry', error);
    window.open(entry.url, '_blank', 'noopener,noreferrer');
  }
}

function syncBadgeToElement(id, el) {
  const stateItem = state.badgeStatus.get(id) || { text: '-', title: t('panel_badge_not_fetched'), loading: false };
  el.textContent = stateItem.text ?? '-';
  el.title = stateItem.title || '';
  el.classList.toggle('loading', Boolean(stateItem.loading));
}

function updateBadgeDom(id) {
  const el = doc.querySelector(`.entry-badge[data-id="${escapeId(id)}"]`);
  if (el) syncBadgeToElement(id, el);
}

function setBadgeLoading(id, loading) {
  const prev = state.badgeStatus.get(id) || { text: '-', title: t('panel_badge_not_fetched'), loading: false };
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
  await refreshAppCatalog();
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
        items.forEach((item) => setBadgeValue(item.id, null, t('panel_badge_permission_required')));
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
            setBadgeValue(item.id, null, t('panel_badge_missing_result'));
            failedHosts.add(host);
          }
        }
      } else {
        const message = res?.error || t('panel_badge_fetch_failed');
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
    setNoticeKey('panel_notice_permission_missing_hosts', { hosts: Array.from(missingPerm).join(', ') });
  } else if (failedHosts.size) {
    setNoticeKey('panel_notice_counts_failed');
  }
}

function attachStorageListener() {
  if (!chrome?.storage?.onChanged) return;
  const listener = (changes, area) => {
    if (area === 'local' && Object.prototype.hasOwnProperty.call(changes, UI_LANGUAGE_KEY)) {
      const nextSetting = normalizeUiLanguageSetting(changes[UI_LANGUAGE_KEY]?.newValue);
      applyUiLanguageSetting(nextSetting, { persist: false, rerender: true }).catch((error) => {
        console.error('Failed to apply launcher language', error);
      });
      return;
    }
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
    // kfavPins is handled by pins.js own storage listener to avoid
    // input-time focus loss caused by duplicated reloads.
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
    if (Object.prototype.hasOwnProperty.call(changes, SHORTCUT_SEARCH_OPEN_MODE_KEY)) {
      state.shortcutSearchOpenMode = normalizeShortcutSearchOpenMode(
        changes[SHORTCUT_SEARCH_OPEN_MODE_KEY]?.newValue
      );
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
    els.toggleShortcuts.textContent = visible ? '▼' : '▶';
    const label = visible ? t('panel_action_toggle_shortcuts_collapse') : t('panel_action_toggle_shortcuts_expand');
    els.toggleShortcuts.title = label;
    els.toggleShortcuts.setAttribute('aria-label', label);
  }
}

function openShortcutTooltip(btn) {
  if (!btn) return;
  const tooltip = btn.querySelector('.shortcut-tooltip');
  if (!tooltip || !tooltip.textContent) return;
  btn.classList.add('tooltip-open');
  btn.classList.remove('tooltip-align-left', 'tooltip-align-right');
  const panelRect = (els.shortcutPanel || els.shortcutRow || doc.body).getBoundingClientRect();
  const tooltipRect = tooltip.getBoundingClientRect();
  if (tooltipRect.left < panelRect.left + 4) {
    btn.classList.add('tooltip-align-left');
    return;
  }
  if (tooltipRect.right > panelRect.right - 4) {
    btn.classList.add('tooltip-align-right');
  }
}

function closeShortcutTooltip(btn) {
  if (!btn) return;
  btn.classList.remove('tooltip-open', 'tooltip-align-left', 'tooltip-align-right');
}

function renderShortcuts() {
  if (!els.shortcutRow) return;
  els.shortcutRow.textContent = '';
  const entries = shortcutState.entries.slice(0, SHORTCUT_MAX_VISIBLE);
  if (!entries.length) {
    const empty = doc.createElement('button');
    empty.type = 'button';
    empty.className = 'shortcut-button shortcut-empty';
    empty.textContent = '+';
    empty.title = t('panel_empty_shortcuts');
    empty.setAttribute('aria-label', t('panel_empty_shortcuts'));
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
    const displayLabel = label || t('panel_entry_no_label');
    const iconColor = resolveShortcutIconColor(entry);
    btn.className = `shortcut-button shortcut-${type}`;
    btn.setAttribute('aria-label', displayLabel);
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

    const tooltip = doc.createElement('span');
    tooltip.className = 'shortcut-tooltip';
    tooltip.textContent = displayLabel;
    tooltip.setAttribute('aria-hidden', 'true');

    btn.appendChild(icon);
    btn.appendChild(tooltip);
    btn.addEventListener('mouseenter', () => openShortcutTooltip(btn));
    btn.addEventListener('focus', () => openShortcutTooltip(btn));
    btn.addEventListener('mouseleave', () => closeShortcutTooltip(btn));
    btn.addEventListener('blur', () => closeShortcutTooltip(btn));
    btn.addEventListener('click', () => openShortcut(entry));
    els.shortcutRow.appendChild(btn);
  });
  renderLucideIcons(els.shortcutRow);
}
async function openShortcut(entry) {
  const url = buildShortcutUrl(entry);
  if (!url) {
    setNoticeKey('panel_notice_shortcut_url_missing');
    return;
  }
  if (entry.host) {
    try {
      await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host: entry.host });
    } catch (_err) {}
  }
  try {
    await openUrlByMode(url, state.shortcutSearchOpenMode);
  } catch (error) {
    console.error('Failed to open shortcut', error);
    window.open(url, '_blank', 'noopener,noreferrer');
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
      setNoticeKey('panel_notice_open_kintone_first');
      return;
    }
    const parsed = parseKintoneUrl(tab.url);
    if (!parsed.host) {
      setNoticeKey('panel_notice_parse_url_failed');
      return;
    }
    const existing = await loadFavorites();
    if (existing.some((item) => item.url === parsed.url)) {
      setNoticeKey('panel_notice_already_registered');
      return;
    }
    const label = (tab.title || '').replace(/ - kintone.*/, '') || t('panel_entry_no_label');
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
    state.badgeStatus.set(entry.id, { text: '-', title: t('panel_badge_not_fetched'), loading: false });
    renderLists();
    setNoticeKey('panel_notice_added_fetching');
    refreshCounts().catch(() => {});
  } catch (error) {
    console.error('Failed to add favorite', error);
    setNoticeKey('panel_notice_add_failed');
  }
}

async function editEntry(entry) {
  const currentLabel = entry.label || '';
  const nextLabel = window.prompt(t('panel_prompt_edit_label'), currentLabel);
  if (nextLabel == null) return;
  const currentUrl = entry.url || '';
  const nextUrl = window.prompt(t('panel_prompt_edit_url'), currentUrl);
  if (nextUrl == null) return;
  if (!nextUrl.trim()) {
    setNoticeKey('panel_notice_url_required');
    return;
  }
  if (!isKintoneUrl(nextUrl.trim())) {
    setNoticeKey('panel_notice_kintone_url_required');
    return;
  }
  const parsed = parseKintoneUrl(nextUrl.trim());
  entry.label = nextLabel.trim() || t('panel_entry_no_label');
  entry.url = parsed.url;
  entry.host = parsed.host;
  entry.appId = parsed.appId;
  entry.viewIdOrName = parsed.viewIdOrName;
  await persistFavorites();
  setNoticeKey('panel_notice_updated');
}

async function changeEntryIcon(entry) {
  const current = normalizeIconName(entry.icon);
  const promptText = t('panel_prompt_icon', { icons: ICON_OPTIONS.join(', ') });
  const input = window.prompt(promptText, current);
  if (input == null) return;
  const value = input.trim();
  entry.icon = normalizeIconName(value);
  await persistFavorites();
  setNoticeKey('panel_notice_icon_updated');
}

async function changeEntryCategory(entry) {
  const current = normalizeCategoryName(entry.category);
  const input = window.prompt(t('panel_prompt_category'), current);
  if (input == null) return;
  entry.category = normalizeCategoryName(input);
  await persistFavorites();
  setNoticeKey('panel_notice_category_updated');
}

async function deleteEntry(id) {
  if (!window.confirm(t('panel_confirm_delete_shortcut'))) return;
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
  if (isEditableTarget(event.target) && event.target !== els.filter) {
    return;
  }
  if (event.key === 'Enter' && isImeComposing(event)) {
    return;
  }
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
    setRecordPinsCollapsed(state.recordPinsCollapsed, { persist: false });
  } else {
    disposeRecordPins();
    renderCollapsedSectionsTray();
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
  ensureRecentCollapseControl();
  ensureCollapsedSectionsTray();
  setRecordPinsCollapsed(state.recordPinsCollapsed, { persist: false });
  setWatchlistCollapsed(state.pinsOnly, { persist: false });
  setRecentRecordsCollapsed(state.recentRecordsCollapsed, { persist: false });
  els.toggleShortcuts?.classList.add('collapse-btn');
  els.toggleRecordPins?.classList.add('collapse-btn');
  els.togglePinsOnlyIcon?.classList.add('collapse-btn');
  els.toggleRecentRecords?.classList.add('collapse-btn');
  ensureAppSearchUi();
  if (els.pinList && !state.pinListObserver) {
    state.pinListObserver = new MutationObserver(() => {
      if (state.recordPinsCollapsed) {
        renderCollapsedSectionsTray();
      }
      applySearchVisibility();
    });
    state.pinListObserver.observe(els.pinList, { childList: true, subtree: false });
  }
  if (els.filter) {
    els.filter.placeholder = t('panel_search_placeholder');
  }
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
  els.toggleRecentRecords?.addEventListener('click', () => {
    setRecentRecordsCollapsed(!state.recentRecordsCollapsed);
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
  if (state.pinListObserver) {
    state.pinListObserver.disconnect();
    state.pinListObserver = null;
  }
  if (els.appSearchToggle && state.appSearchToggleHandler) {
    els.appSearchToggle.removeEventListener('click', state.appSearchToggleHandler);
    state.appSearchToggleHandler = null;
  }
  if (els.appSearchInput && state.appSearchInputHandler) {
    els.appSearchInput.removeEventListener('input', state.appSearchInputHandler);
    state.appSearchInputHandler = null;
  }
  if (els.appSearchInput && state.appSearchInputKeydownHandler) {
    els.appSearchInput.removeEventListener('keydown', state.appSearchInputKeydownHandler);
    state.appSearchInputKeydownHandler = null;
  }
  if (state.appSearchOutsideHandler) {
    document.removeEventListener('mousedown', state.appSearchOutsideHandler);
    state.appSearchOutsideHandler = null;
  }
  if (els.toggleShortcuts && state.shortcutToggleHandler) {
    els.toggleShortcuts.removeEventListener('click', state.shortcutToggleHandler);
    state.shortcutToggleHandler = null;
  }
}

async function init() {
  await initializeI18n();
  await loadShortcutSearchOpenMode();
  await loadCollapsedSectionsState();
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
  setNoticeKey('panel_notice_initialization_failed');
});

window.addEventListener('beforeunload', dispose);


