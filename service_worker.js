// service_worker.js
// 目的：アクティブな kintone タブがあるときに、指定ウォッチリスト項目の件数を取得して拡張アイコンのバッジ表示を更新
import {
  createId,
  isKintoneUrl,
  parseKintoneUrl,
  loadShortcuts,
  saveShortcuts,
  upsertRecentRecord,
  saveAppNameMap
} from './core.js';

async function pickBadgeTarget() {
  const { kintoneFavorites = [], kfavBadgeTargetId = null } = await chrome.storage.sync.get([
    'kintoneFavorites',
    'kfavBadgeTargetId'
  ]);
  if (!kintoneFavorites.length) return null;
  const target = kfavBadgeTargetId
    ? kintoneFavorites.find(x => x.id === kfavBadgeTargetId)
    : kintoneFavorites[0]; // 初期は先頭
  return target || null;
}

async function updateBadgeForTab(tabId, tabUrl) {
  if (!tabUrl || !/^https:\/\/.*(kintone|cybozu)\.com/.test(tabUrl)) {
    await chrome.action.setBadgeText({ tabId, text: '' });
    return;
  }
  const target = await pickBadgeTarget();
  if (!target) {
    await chrome.action.setBadgeText({ tabId, text: '' });
    return;
  }
  if (!tabUrl.includes(target.host)) {
    await chrome.action.setBadgeText({ tabId, text: '' });
    return;
  }
  try {
    const res = await chrome.tabs.sendMessage(tabId, {
      type: target.viewIdOrName ? 'COUNT_VIEW' : 'COUNT_APP_QUERY',
      payload: target.viewIdOrName
        ? { appId: target.appId, viewIdOrName: target.viewIdOrName }
        : { appId: target.appId, query: target.query || '' }
    });
    if (res?.ok) {
      await chrome.action.setBadgeText({ tabId, text: String(res.count ?? '') });
      await chrome.action.setBadgeBackgroundColor({ tabId, color: '#777' });
    } else {
      await chrome.action.setBadgeText({ tabId, text: '' });
    }
  } catch (_e) {
    await chrome.action.setBadgeText({ tabId, text: '' });
  }
}

chrome.tabs.onActivated.addListener(async ({ tabId }) => {
  const tab = await chrome.tabs.get(tabId);
  updateBadgeForTab(tabId, tab.url);
});
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (changeInfo.status === 'complete' || changeInfo.url) {
    updateBadgeForTab(tabId, tab.url);
  }
});

// MV3は常駐しないため、定期起床して更新
chrome.runtime.onInstalled.addListener(() => {
  chrome.alarms.create('kfav-badge-refresh', { periodInMinutes: 2 });
 // Edge/Chrome: サイドパネルを有効化 & アイコンクリックで開く挙動を設定
  try {
    if (chrome.sidePanel) {
      if (chrome.sidePanel.setOptions) {
        chrome.sidePanel.setOptions({ path: 'sidepanel.html', enabled: true });
      }
      if (chrome.sidePanel.setPanelBehavior) {
        chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
      }
    }
  } catch (e) {
    // 起動時/導入時に自動でタブを作成しない
  }

  setupContextMenus();
  migrateShortcutsFillIconFieldsOnce().catch(() => {});
});

// 起動時に自動でコネクタタブを作成しない
chrome.runtime.onStartup?.addListener(() => setupContextMenus());
chrome.runtime.onStartup?.addListener(() => {
  migrateShortcutsFillIconFieldsOnce().catch(() => {});
});

const PIN_MENU_ID = 'kfav-pin-record';
const FAVORITE_PARENT_ID = 'kfav-favorite-parent';
const FAVORITE_MENU_APP = 'kfav-favorite-app';
const FAVORITE_MENU_VIEW = 'kfav-favorite-view';
const FAVORITES_KEY = 'kintoneFavorites';
const SHORTCUT_PARENT_ID = 'kfav-shortcut-parent';
const SHORTCUT_MENU_APP = 'kfav-shortcut-app';
const SHORTCUT_MENU_VIEW = 'kfav-shortcut-view';
const SHORTCUT_MENU_CREATE = 'kfav-shortcut-create';

async function setupContextMenus() {
  if (!chrome.contextMenus?.create) return;
  const patterns = [
    'https://*.kintone.com/*',
    'https://*.kintone-dev.com/*',
    'https://*.cybozu.com/*'
  ];
  try {
    await chrome.contextMenus.removeAll();
  } catch (_err) {
    // ignore
  }
  try {
    chrome.contextMenus.create({
      id: PIN_MENU_ID,
      title: 'このレコードをピン留め',
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: FAVORITE_PARENT_ID,
      title: 'ウォッチリストに追加',
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: FAVORITE_MENU_APP,
      parentId: FAVORITE_PARENT_ID,
      title: 'アプリトップ',
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: FAVORITE_MENU_VIEW,
      parentId: FAVORITE_PARENT_ID,
      title: 'ビューを追加',
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: SHORTCUT_PARENT_ID,
      title: 'ショートカットに追加',
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: SHORTCUT_MENU_APP,
      parentId: SHORTCUT_PARENT_ID,
      title: 'アプリトップ',
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: SHORTCUT_MENU_VIEW,
      parentId: SHORTCUT_PARENT_ID,
      title: 'ビューを開く',
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: SHORTCUT_MENU_CREATE,
      parentId: SHORTCUT_PARENT_ID,
      title: 'レコード新規',
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
  } catch (_err) {
    // ignore
  }
}

chrome.contextMenus?.onClicked.addListener(async (info, tab) => {
  if (info.menuItemId === FAVORITE_MENU_APP || info.menuItemId === FAVORITE_MENU_VIEW) {
    await handleFavoriteContextMenu(info.menuItemId, info, tab);
    return;
  }
  if (info.menuItemId === PIN_MENU_ID) {
    const pageUrl = info.pageUrl || tab?.url || '';
    const { host, appId, recordId } = parseKintoneUrl(pageUrl);
    if (!host || !appId || !recordId) return;
    const stored = await chrome.storage.sync.get('kfavPins');
    const raw = stored.kfavPins;
    const entries = Array.isArray(raw) ? raw : raw ? [raw] : [];
    if (entries.some((entry) => entry?.host === host && entry?.appId === appId && entry?.recordId === recordId)) {
      flashBadge(tab?.id, '!', '#e74c3c');
      return;
    }
    const label = (tab?.title || '').replace(/\s+-\s+kintone.*/i, '').trim();
    entries.push({
      id: createId(),
      label,
      host,
      appId,
      recordId,
      titleField: '',
      note: ''
    });
    await chrome.storage.sync.set({ kfavPins: entries });
    flashBadge(tab?.id, '✓', '#2ecc71');
    return;
  }
  if (info.menuItemId === SHORTCUT_MENU_APP || info.menuItemId === SHORTCUT_MENU_VIEW || info.menuItemId === SHORTCUT_MENU_CREATE) {
    await handleShortcutContextMenu(info.menuItemId, info, tab);
  }
});

function flashBadge(tabId, text, color) {
  if (!tabId) return;
  try {
    if (color) chrome.action.setBadgeBackgroundColor({ tabId, color });
    chrome.action.setBadgeText({ tabId, text });
    setTimeout(() => {
      chrome.action.setBadgeText({ tabId, text: '' }).catch(() => {});
    }, 1200);
  } catch (_e) {
    // ignore
  }
}

const DEFAULT_SHORTCUT_ICON = 'file-text';
const DEFAULT_SHORTCUT_ICON_COLOR = 'gray';
const SHORTCUT_ICON_MIGRATION_KEY = 'kfavShortcutsIconMigratedV1';

function toShortcutArray(raw) {
  if (!raw) return [];
  if (Array.isArray(raw)) return raw.slice();
  if (raw && typeof raw === 'object') return [raw];
  return [];
}

async function migrateShortcutsFillIconFieldsOnce() {
  try {
    const mark = await chrome.storage.local.get(SHORTCUT_ICON_MIGRATION_KEY);
    if (mark?.[SHORTCUT_ICON_MIGRATION_KEY]) return;
  } catch (_err) {
    // ignore
  }

  const stored = await chrome.storage.sync.get('kfavShortcuts');
  const list = toShortcutArray(stored.kfavShortcuts);
  if (list.length) {
    const normalized = await loadShortcuts();
    await saveShortcuts(normalized);
  }

  try {
    await chrome.storage.local.set({ [SHORTCUT_ICON_MIGRATION_KEY]: true });
  } catch (_err) {
    // ignore
  }
}

function favoriteUrlOf(host, appId, viewIdOrName) {
  const appKey = String(appId || '').trim();
  const base = `${host}/k/${appKey}/`;
  if (viewIdOrName) {
    return `${base}?view=${encodeURIComponent(String(viewIdOrName))}`;
  }
  return base;
}

async function loadFavoriteList() {
  const stored = await chrome.storage.sync.get(FAVORITES_KEY);
  const raw = stored[FAVORITES_KEY];
  if (!raw) return [];
  if (Array.isArray(raw)) return raw.slice();
  if (raw && typeof raw === 'object') return [raw];
  return [];
}

async function saveFavoriteList(list) {
  await chrome.storage.sync.set({ [FAVORITES_KEY]: list });
}

async function handleFavoriteContextMenu(menuId, info, tab) {
  const pageUrl = info.pageUrl || tab?.url || '';
  const { host, appId, viewIdOrName } = parseKintoneUrl(pageUrl);
  if (!host) {
    flashBadge(tab?.id, '!', '#e74c3c');
    return;
  }
  const appIdStr = String(appId || '').trim();
  if (!appIdStr) {
    flashBadge(tab?.id, '!', '#e74c3c');
    return;
  }
  const type = menuId === FAVORITE_MENU_VIEW ? 'view' : 'app';
  const viewKey = type === 'view' ? String(viewIdOrName || '').trim() : '';
  if (type === 'view' && !viewKey) {
    flashBadge(tab?.id, '!', '#e67e22');
    return;
  }
  const targetUrl = favoriteUrlOf(host, appIdStr, viewKey);
  const list = await loadFavoriteList();
  if (list.some((item) => item?.url === targetUrl)) {
    flashBadge(tab?.id, '!', '#e67e22');
    return;
  }
  const labelFromTab = (tab?.title || '').replace(/\s+-\s+kintone.*/i, '').trim();
  const fallbackLabel = type === 'view' ? `ビュー ${viewKey}` : `App ${appIdStr}`;
  const entry = {
    id: createId(),
    label: labelFromTab || fallbackLabel,
    url: targetUrl,
    host,
    appId: appIdStr,
    viewIdOrName: viewKey,
    query: '',
    order: list.length,
    pinned: false
  };
  list.push(entry);
  list.forEach((item, index) => {
    if (typeof item.order !== 'number') item.order = index;
  });
  await saveFavoriteList(list);
  try {
    await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host });
  } catch (_err) {
    // ignore
  }
  flashBadge(tab?.id, '+', '#f1c40f');
}

function shortcutDefaultLabel(type, appId, viewIdOrName) {
  if (type === 'view') {
    return viewIdOrName ? `ビュー ${viewIdOrName}` : `アプリ ${appId}`;
  }
  if (type === 'create') return 'レコード新規';
  return `アプリ ${appId}`;
}

async function handleShortcutContextMenu(menuId, info, tab) {
  const type = menuId === SHORTCUT_MENU_VIEW
    ? 'view'
    : menuId === SHORTCUT_MENU_CREATE
      ? 'create'
      : 'appTop';
  const pageUrl = info.pageUrl || tab?.url || '';
  const { host, appId, viewIdOrName } = parseKintoneUrl(pageUrl);
  if (!host || !appId) {
    flashBadge(tab?.id, '!', '#e74c3c');
    return;
  }
  if (type === 'view' && !viewIdOrName) {
    flashBadge(tab?.id, '!', '#e67e22');
    return;
  }
  const list = await loadShortcuts();
  const appIdStr = String(appId || '').trim();
  const normalizedView = type === 'view' ? String(viewIdOrName || '') : '';
  const duplicate = list.some((item) => {
    if (!item) return false;
    const sameType = item.type === type;
    const sameHost = item.host === host;
    const sameApp = String(item.appId || '').trim() === appIdStr;
    const sameView = String(item.viewIdOrName || '') === normalizedView;
    return sameType && sameHost && sameApp && sameView;
  });
  if (duplicate) {
    flashBadge(tab?.id, '!', '#e67e22');
    return;
  }
  const title = (tab?.title || '').replace(/\s+-\s+kintone.*/i, '').trim();
  const label = title || shortcutDefaultLabel(type, appIdStr, normalizedView);
  list.push({
    id: createId(),
    type,
    host,
    appId: appIdStr,
    viewIdOrName: normalizedView,
    label,
    initial: '',
    icon: DEFAULT_SHORTCUT_ICON,
    iconColor: DEFAULT_SHORTCUT_ICON_COLOR,
    order: list.length
  });
  await saveShortcuts(list);
  flashBadge(tab?.id, '+', '#2980b9');
}

async function listKintoneTabsByHost(host) {
  const tabs = await chrome.tabs.query({ url: [`${host}/*`] });
  return tabs;
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
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

const READY_CHECK_INTERVAL_MS = 400;
const READY_CHECK_SHORT_ATTEMPTS = 5;
const SIGN_IN_REQUIRED_MESSAGE = 'kintoneにサインインしてから再実行してください';
const APPS_MAP_PAGE_LIMIT = 100;
const APPS_MAP_MAX_PAGES = 200;

async function checkKintoneReady(tabId) {
  if (!tabId) return false;
  try {
    const res = await chrome.tabs.sendMessage(tabId, { type: 'CHECK_KINTONE_READY' });
    return !!res?.ok && !!res.ready;
  } catch (_err) {
    return false;
  }
}

async function prepareConnectorTab(tabId, attempts) {
  if (!tabId) return false;
  const max = attempts ?? READY_CHECK_SHORT_ATTEMPTS;
  for (let i = 0; i < max; i++) {
    const injected = await ensureInjection(tabId);
    if (injected && (await checkKintoneReady(tabId))) {
      return true;
    }
    await sleep(READY_CHECK_INTERVAL_MS);
  }
  return false;
}

async function ensureInjection(tabId) {
  try {
    await chrome.scripting.executeScript({ target: { tabId }, files: ['content.js'] });
    const bridgeSrc = chrome.runtime.getURL('page-bridge.js');
    await chrome.scripting.executeScript({
      target: { tabId },
      world: 'MAIN',
      func: (src) => {
        if (!document.getElementById('__kfav_bridge')) {
          const s = document.createElement('script');
          s.id = '__kfav_bridge';
          s.src = src;
          (document.head || document.documentElement).appendChild(s);
        }
      },
      args: [bridgeSrc]
    });
    return true;
  } catch (e) {
    return false;
  }
}

async function ensureConnectorTab(host) {
  if (!host) return null;
  const existing = await listKintoneTabsByHost(host);
  for (const tab of existing) {
    if (!tab?.id) continue;
    const ready = await prepareConnectorTab(tab.id, READY_CHECK_SHORT_ATTEMPTS);
    if (ready) return tab;
  }
  if (!existing.length) {
    throw new Error('No open kintone tab found for this host');
  }
  throw new Error(SIGN_IN_REQUIRED_MESSAGE);
}

async function runForwardOnHost(host, forward) {
  const safeHost = normalizeHostOrigin(host);
  if (!safeHost) throw new Error('host required');
  const tab = await ensureConnectorTab(safeHost);
  if (!tab) throw new Error('no tab');
  await ensureInjection(tab.id);
  return await chrome.tabs.sendMessage(tab.id, forward);
}

async function fetchAppsMap(host) {
  const safeHost = normalizeHostOrigin(host);
  if (!safeHost) throw new Error('host required');
  const out = {};
  let offset = 0;
  for (let page = 0; page < APPS_MAP_MAX_PAGES; page++) {
    const url = `${safeHost}/k/v1/apps.json?limit=${APPS_MAP_PAGE_LIMIT}&offset=${offset}`;
    const resp = await fetch(url, {
      method: 'GET',
      credentials: 'include',
      headers: {
        Accept: 'application/json',
        'X-Requested-With': 'XMLHttpRequest'
      }
    });
    if (!resp.ok) {
      const text = await resp.text().catch(() => '');
      console.debug('[kfav] apps.json failed', { url, status: resp.status, text: String(text).slice(0, 300) });
      throw new Error(`apps.json failed (${resp.status})`);
    }
    const json = await resp.json();
    const apps = Array.isArray(json?.apps) ? json.apps : [];
    apps.forEach((item) => {
      const appId = String(item?.appId || item?.id || '').trim();
      const name = String(item?.name || '').trim();
      if (!appId || !name) return;
      out[appId] = name;
    });
    if (apps.length < APPS_MAP_PAGE_LIMIT) break;
    offset += apps.length;
  }
  return out;
}

async function fetchAppsMapByAppIds(host, appIds) {
  const safeHost = normalizeHostOrigin(host);
  const uniqueIds = Array.from(new Set((appIds || []).map((id) => String(id || '').trim()).filter((id) => /^\d+$/.test(id))));
  const out = {};
  if (!safeHost || !uniqueIds.length) return out;
  for (const appId of uniqueIds) {
    const url = `${safeHost}/k/v1/app.json?id=${encodeURIComponent(appId)}`;
    try {
      const resp = await fetch(url, {
        method: 'GET',
        credentials: 'include',
        headers: {
          Accept: 'application/json',
          'X-Requested-With': 'XMLHttpRequest'
        }
      });
      if (!resp.ok) {
        const text = await resp.text().catch(() => '');
        console.debug('[kfav] app.json fallback failed', { url, status: resp.status, text: String(text).slice(0, 300) });
        continue;
      }
      const data = await resp.json();
      const name = String(data?.name || data?.app?.name || '').trim();
      if (name) out[appId] = name;
    } catch (error) {
      console.debug('[kfav] app.json fallback error', { appId, error: String(error?.message || error) });
    }
  }
  return out;
}

// ===== UIから「kintoneで実行して結果を返して」の汎用エンドポイント =====
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  (async () => {
    if (msg?.type === 'WARMUP_CONNECTOR') {
      const safeHost = normalizeHostOrigin(msg.host);
      if (!safeHost) { sendResponse({ ok: false, error: 'host required' }); return; }
      try {
        await ensureConnectorTab(safeHost);
        sendResponse({ ok: true });
      } catch (e) {
        sendResponse({ ok: false, error: String(e) });
      }
      return;
    }

    if (msg?.type === 'RECENT_UPSERT') {
      try {
        const list = await upsertRecentRecord(msg.payload || {});
        sendResponse({ ok: true, size: Array.isArray(list) ? list.length : 0 });
      } catch (e) {
        sendResponse({ ok: false, error: String(e) });
      }
      return;
    }

    if (msg?.type === 'GET_APPS_MAP') {
      const safeHost = normalizeHostOrigin(msg?.payload?.host || msg?.host);
      const appIds = Array.isArray(msg?.payload?.appIds)
        ? msg.payload.appIds
        : (Array.isArray(msg?.appIds) ? msg.appIds : []);
      if (!safeHost) {
        sendResponse({ ok: false, error: 'host required' });
        return;
      }
      try {
        const map = await fetchAppsMap(safeHost);
        await saveAppNameMap(safeHost, map);
        console.debug('[kfav] apps map loaded', { host: safeHost, size: Object.keys(map).length });
        sendResponse({ ok: true, host: safeHost, map });
      } catch (error) {
        console.debug('[kfav] apps map load failed', { host: safeHost, error: String(error?.message || error) });
        const fallbackMap = await fetchAppsMapByAppIds(safeHost, appIds);
        if (Object.keys(fallbackMap).length) {
          await saveAppNameMap(safeHost, fallbackMap);
          console.debug('[kfav] apps map fallback loaded', { host: safeHost, size: Object.keys(fallbackMap).length });
          sendResponse({ ok: true, host: safeHost, map: fallbackMap, fallback: true });
          return;
        }
        sendResponse({ ok: false, host: safeHost, error: String(error?.message || error), map: {} });
      }
      return;
    }

    if (msg?.type !== 'RUN_IN_KINTONE') return;
    const safeHost = normalizeHostOrigin(msg.host);
    const { forward } = msg; // forward: { type: 'LIST_FIELDS' | 'LIST_SCHEDULE' | 'LIST_VIEWS' | ... , payload: {...} }
    if (!safeHost || !forward?.type) { sendResponse({ ok:false, error:'bad request' }); return; }
    try {
      const res = await runForwardOnHost(safeHost, forward);
      sendResponse(res || { ok:false, error:'no response' });
    } catch (e) {
      sendResponse({ ok:false, error:String(e) });
    }
  })();
  return true;
});

// ブラウザ起動時も念のため再設定（MV3はSWが落ちるため）
chrome.runtime.onStartup?.addListener(() => {
  try {
    if (chrome.sidePanel) {
      chrome.sidePanel.setOptions?.({ path: 'sidepanel.html', enabled: true });
      chrome.sidePanel.setPanelBehavior?.({ openPanelOnActionClick: true });
    }
  } catch {}
});

chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name !== 'kfav-badge-refresh') return;
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (tab) updateBadgeForTab(tab.id, tab.url);
});
// 拡張アイコンクリック → サイドパネルを開く（非対応ならサイドパネルの内容を新規タブで開く）
chrome.action.onClicked.addListener(async (tab) => {
  if (chrome.sidePanel?.open) {
    try {
      // タブ単位のパス指定（Chrome 114）
      if (chrome.sidePanel.setOptions && tab?.id) {
        await chrome.sidePanel.setOptions({ tabId: tab.id, path: 'sidepanel.html', enabled: true });
      }
      await chrome.sidePanel.open({ tabId: tab?.id });
      return;
    } catch (e) {
      // フォールバックに進む
    }
  }
  // 自動で拡張ページのタブを作成しない
});

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  (async () => {
    if (msg?.type !== 'SWITCH_SIDE_PANEL') return;
    const path = msg.path || 'sidepanel.html';
    const tabId = msg.tabId || sender?.tab?.id;
    try {
      if (chrome.sidePanel?.setOptions) {
        await chrome.sidePanel.setOptions({ tabId, path, enabled: true });
        await chrome.sidePanel.open?.({ tabId });
        sendResponse({ ok: true });
      } else {
        sendResponse({ ok: false, error: 'sidePanel API is not available' });
      }
    } catch (e) {
      sendResponse({ ok: false, error: String(e) });
    }
  })();
  return true;
});







