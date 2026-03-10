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
chrome.storage.onChanged.addListener((changes, area) => {
  if (area !== 'local') return;
  if (!changes || !Object.prototype.hasOwnProperty.call(changes, UI_LANGUAGE_KEY)) return;
  setupContextMenus().catch((error) => {
    console.error('[kfav] failed to rebuild context menus on language change', error);
  });
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
const DEV_PARENT_MENU_ID = 'kf-dev-root';
const DEV_COPY_QUERY_MENU_ID = 'kf-dev-copy-query';
const DEV_COPY_FIELDS_LIST_MENU_ID = 'kf-dev-copy-fields-list';
const DEV_COPY_APP_ID_MENU_ID = 'kf-dev-copy-app-id';
const DEV_COPY_VIEW_ID_MENU_ID = 'kf-dev-copy-view-id';
const PRO_INSTALL_TYPE_MESSAGE_ID = 'PB_GET_INSTALL_TYPE';
const UI_LANGUAGE_KEY = 'uiLanguage';
const UI_LANGUAGE_VALUES = new Set(['auto', 'ja', 'en']);
const latestPageContextByTab = new Map();

const MENU_MESSAGES = {
  ja: {
    menu_pin_record: 'このレコードをピン留め',
    menu_favorite_parent: 'ウォッチリストに追加',
    menu_favorite_app: 'アプリトップ',
    menu_favorite_view: 'ビューを追加',
    menu_shortcut_parent: 'ショートカットに追加',
    menu_shortcut_app: 'アプリトップ',
    menu_shortcut_view: 'ビューを開く',
    menu_shortcut_create: 'レコード新規',
    menu_dev_root: '開発者ツール',
    menu_copy_query: '絞り込みクエリをコピー',
    menu_copy_fields: 'フィールド一覧をコピー',
    menu_copy_app_id: 'appId をコピー',
    menu_copy_view_id: 'viewId をコピー'
  },
  en: {
    menu_pin_record: 'Pin this record',
    menu_favorite_parent: 'Add to watchlist',
    menu_favorite_app: 'App top',
    menu_favorite_view: 'Add view',
    menu_shortcut_parent: 'Add to shortcuts',
    menu_shortcut_app: 'App top',
    menu_shortcut_view: 'Open view',
    menu_shortcut_create: 'New record',
    menu_dev_root: 'Development Tools',
    menu_copy_query: 'Copy Filter Query',
    menu_copy_fields: 'Copy Field List',
    menu_copy_app_id: 'Copy appId',
    menu_copy_view_id: 'Copy viewId'
  }
};

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg?.type !== PRO_INSTALL_TYPE_MESSAGE_ID) return;
  (async () => {
    try {
      if (chrome.management?.getSelf) {
        const info = await chrome.management.getSelf();
        sendResponse({
          ok: true,
          installType: String(info?.installType || ''),
          source: 'management'
        });
        return;
      }
      const manifest = chrome.runtime?.getManifest?.() || {};
      const fallbackDevelopment = !Object.prototype.hasOwnProperty.call(manifest, 'update_url');
      sendResponse({
        ok: false,
        reason: 'management_api_unavailable',
        fallbackDevelopment
      });
    } catch (error) {
      const manifest = chrome.runtime?.getManifest?.() || {};
      const fallbackDevelopment = !Object.prototype.hasOwnProperty.call(manifest, 'update_url');
      sendResponse({
        ok: false,
        reason: String(error?.message || error || 'management_get_self_failed'),
        fallbackDevelopment
      });
    }
  })();
  return true;
});

function normalizeUiLanguage(raw) {
  const value = String(raw || '').trim().toLowerCase();
  if (value === 'ja' || value.startsWith('ja-')) return 'ja';
  return 'en';
}

function normalizeUiLanguageSetting(raw) {
  const value = String(raw || '').trim().toLowerCase();
  if (UI_LANGUAGE_VALUES.has(value)) return value;
  return 'auto';
}

async function resolveUiLanguage() {
  let setting = 'auto';
  try {
    const data = await chrome.storage.local.get(UI_LANGUAGE_KEY);
    setting = normalizeUiLanguageSetting(data?.[UI_LANGUAGE_KEY]);
  } catch (_err) {
    setting = 'auto';
  }
  if (setting === 'ja' || setting === 'en') return setting;
  try {
    const browserLang = navigator.language || (Array.isArray(navigator.languages) ? navigator.languages[0] : '');
    return normalizeUiLanguage(browserLang);
  } catch (_err) {
    return 'ja';
  }
}

function tMenu(lang, key) {
  return MENU_MESSAGES[lang]?.[key]
    ?? MENU_MESSAGES.ja?.[key]
    ?? key;
}

async function getActiveTabForDevAction(fallbackTab) {
  try {
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    if (tabs?.[0]?.id) return tabs[0];
  } catch (_err) {
    // ignore
  }
  return fallbackTab || null;
}

async function fetchFieldsListFromMainWorld(tabId) {
  const injected = await chrome.scripting.executeScript({
    target: { tabId },
    world: 'MAIN',
    func: () => {
      try {
        const sdk = window.kintone;
        if (!sdk || !sdk.app || !sdk.api) {
          return { ok: false, reason: 'not_kintone_page', url: location.href };
        }
        const appId = typeof sdk.app.getId === 'function' ? sdk.app.getId() : '';
        if (!appId) {
          return { ok: false, reason: 'app_id_not_found', url: location.href };
        }
        return sdk.api(
          sdk.api.url('/k/v1/app/form/fields', true),
          'GET',
          { app: appId }
        ).then((resp) => {
          const properties = resp && resp.properties ? resp.properties : {};
          if (!properties || typeof properties !== 'object' || !Object.keys(properties).length) {
            return { ok: false, reason: 'empty_properties', appId: String(appId || ''), url: location.href };
          }

          const rows = [];
          Object.keys(properties).forEach((code) => {
            const field = properties[code] || {};
            const type = String(field.type || '');
            if (type === 'SUBTABLE') {
              const subFields = field.fields && typeof field.fields === 'object' ? field.fields : {};
              Object.keys(subFields).forEach((subCode) => {
                const sub = subFields[subCode] || {};
                rows.push({
                  scope: 'サブテーブル',
                  parentTableCode: code,
                  parentTableLabel: String(field.label || ''),
                  code: String(subCode || ''),
                  label: String(sub.label || ''),
                  type: String(sub.type || '')
                });
              });
              return;
            }

            rows.push({
              scope: '通常',
              parentTableCode: '',
              parentTableLabel: '',
              code: String(code || ''),
              label: String(field.label || ''),
              type: type
            });
          });

          if (!rows.length) {
            return { ok: false, reason: 'empty_properties', appId: String(appId || ''), url: location.href };
          }

          rows.sort((a, b) => {
            const aScope = a.scope === '通常' ? 0 : 1;
            const bScope = b.scope === '通常' ? 0 : 1;
            if (aScope !== bScope) return aScope - bScope;
            const parentCmp = String(a.parentTableCode || '').localeCompare(String(b.parentTableCode || ''), 'ja');
            if (parentCmp !== 0) return parentCmp;
            return String(a.code || '').localeCompare(String(b.code || ''), 'ja');
          });

          const lines = [];
          lines.push('所属\t親テーブルコード\t親テーブル名\tフィールドコード\tフィールド名\tフィールドタイプ');
          rows.forEach((row) => {
            lines.push([
              row.scope,
              row.parentTableCode,
              row.parentTableLabel,
              row.code,
              row.label,
              row.type
            ].join('\t'));
          });

          return {
            ok: true,
            appId: String(appId || ''),
            url: location.href,
            rowCount: rows.length,
            text: lines.join('\n')
          };
        }).catch((error) => ({
          ok: false,
          reason: 'form_fields_fetch_failed',
          appId: String(appId || ''),
          url: location.href,
          detail: String(error && error.message || error)
        }));
      } catch (error) {
        return {
          ok: false,
          reason: 'unexpected_error',
          url: location.href,
          detail: String(error && error.message || error)
        };
      }
    }
  });
  return injected?.[0]?.result || { ok: false, reason: 'unexpected_error' };
}

async function copyTextViaInjectedScript(tabId, text) {
  try {
    const injected = await chrome.scripting.executeScript({
      target: { tabId },
      world: 'ISOLATED',
      func: async (value) => {
        const textValue = String(value || '');
        if (!textValue) return { ok: false, reason: 'empty_text' };
        try {
          if (navigator?.clipboard?.writeText) {
            await navigator.clipboard.writeText(textValue);
            return { ok: true };
          }
        } catch (_err) {
          // fallback below
        }
        try {
          const ta = document.createElement('textarea');
          ta.value = textValue;
          ta.setAttribute('readonly', 'readonly');
          ta.style.position = 'fixed';
          ta.style.left = '-9999px';
          ta.style.top = '-9999px';
          document.body.appendChild(ta);
          ta.select();
          ta.setSelectionRange(0, ta.value.length);
          const ok = document.execCommand('copy');
          ta.remove();
          return ok ? { ok: true } : { ok: false, reason: 'clipboard_failed' };
        } catch (_err) {
          return { ok: false, reason: 'clipboard_failed' };
        }
      },
      args: [text]
    });
    return injected?.[0]?.result || { ok: false, reason: 'copy_fallback_failed' };
  } catch (error) {
    return { ok: false, reason: 'copy_fallback_failed', detail: String(error?.message || error) };
  }
}

async function fetchAppOrViewIdFromMainWorld(tabId, mode) {
  const injected = await chrome.scripting.executeScript({
    target: { tabId },
    world: 'MAIN',
    func: async (targetMode) => {
      try {
        if (!window.kintone || !kintone.app) {
          return { ok: false, reason: 'not_kintone_page', value: '', url: location.href };
        }
        if (targetMode === 'appId') {
          const appId = typeof kintone.app.getId === 'function' ? kintone.app.getId() : '';
          if (!appId) {
            return { ok: false, reason: 'app_id_not_found', value: '', url: location.href };
          }
          return { ok: true, value: String(appId), url: location.href };
        }

        if (targetMode === 'viewId' || targetMode === 'view') {
          try {
            if (typeof kintone.app.getView === 'function') {
              const view = await Promise.resolve(kintone.app.getView());
              const id = view && view.id != null ? String(view.id).trim() : '';
              if (id) {
                return { ok: true, value: id, source: 'getView', url: location.href };
              }
            }
          } catch (_err) {
            // continue to URL fallback
          }

          try {
            const url = new URL(location.href);
            const viewParam = String(url.searchParams.get('view') || '').trim();
            if (viewParam) {
              return { ok: true, value: viewParam, source: 'url', url: location.href };
            }
          } catch (_err) {
            // ignore
          }

          return { ok: false, reason: 'view_id_not_available', value: '', source: 'none', url: location.href };
        }
        return { ok: false, reason: 'invalid_mode', value: '', url: location.href };
      } catch (error) {
        return {
          ok: false,
          reason: 'unexpected_error',
          value: '',
          url: location.href,
          detail: String(error?.message || error)
        };
      }
    },
    args: [mode]
  });
  return injected?.[0]?.result || { ok: false, reason: 'unexpected_error', value: '' };
}

async function copyDevTextToActiveTab(tabId, text, label) {
  let copied = null;
  try {
    copied = await chrome.tabs.sendMessage(tabId, {
      type: 'DEV_COPY_TEXT',
      text: String(text || ''),
      label
    });
  } catch (sendError) {
    copied = {
      ok: false,
      reason: 'receiving_end_does_not_exist',
      detail: String(sendError?.message || sendError)
    };
  }
  if (copied?.ok) return copied;
  const fallback = await copyTextViaInjectedScript(tabId, String(text || ''));
  if (fallback?.ok) return fallback;
  return { ok: false, reason: copied?.reason || fallback?.reason || 'copy_failed' };
}

async function setupContextMenus() {
  if (!chrome.contextMenus?.create) return;
  const lang = await resolveUiLanguage();
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
      title: tMenu(lang, 'menu_pin_record'),
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: FAVORITE_PARENT_ID,
      title: tMenu(lang, 'menu_favorite_parent'),
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: FAVORITE_MENU_APP,
      parentId: FAVORITE_PARENT_ID,
      title: tMenu(lang, 'menu_favorite_app'),
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: FAVORITE_MENU_VIEW,
      parentId: FAVORITE_PARENT_ID,
      title: tMenu(lang, 'menu_favorite_view'),
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: SHORTCUT_PARENT_ID,
      title: tMenu(lang, 'menu_shortcut_parent'),
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: SHORTCUT_MENU_APP,
      parentId: SHORTCUT_PARENT_ID,
      title: tMenu(lang, 'menu_shortcut_app'),
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: SHORTCUT_MENU_VIEW,
      parentId: SHORTCUT_PARENT_ID,
      title: tMenu(lang, 'menu_shortcut_view'),
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: SHORTCUT_MENU_CREATE,
      parentId: SHORTCUT_PARENT_ID,
      title: tMenu(lang, 'menu_shortcut_create'),
      contexts: ['page', 'frame'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: DEV_PARENT_MENU_ID,
      title: tMenu(lang, 'menu_dev_root'),
      contexts: ['all'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: DEV_COPY_QUERY_MENU_ID,
      parentId: DEV_PARENT_MENU_ID,
      title: tMenu(lang, 'menu_copy_query'),
      contexts: ['all'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: DEV_COPY_FIELDS_LIST_MENU_ID,
      parentId: DEV_PARENT_MENU_ID,
      title: tMenu(lang, 'menu_copy_fields'),
      contexts: ['all'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: DEV_COPY_APP_ID_MENU_ID,
      parentId: DEV_PARENT_MENU_ID,
      title: tMenu(lang, 'menu_copy_app_id'),
      contexts: ['all'],
      documentUrlPatterns: patterns
    });
    chrome.contextMenus.create({
      id: DEV_COPY_VIEW_ID_MENU_ID,
      parentId: DEV_PARENT_MENU_ID,
      title: tMenu(lang, 'menu_copy_view_id'),
      contexts: ['all'],
      documentUrlPatterns: patterns
    });
  } catch (_err) {
    // ignore
  }
}

chrome.contextMenus?.onClicked.addListener(async (info, tab) => {
  if (info.menuItemId === DEV_COPY_QUERY_MENU_ID) {
    try {
      const activeTab = await getActiveTabForDevAction(tab);
      if (!activeTab?.id) {
        console.debug('[kfav][DEV] no tab for context menu click');
        return;
      }
      const res = await chrome.tabs.sendMessage(activeTab.id, { type: 'DEV_COPY_QUERY' });
      if (!res?.ok) {
        console.debug('[kfav][DEV] copy failed', { type: 'DEV_COPY_QUERY', reason: res?.reason || res?.error || 'unknown' });
      } else {
        console.debug('[kfav][DEV] copied', { type: 'DEV_COPY_QUERY', value: res?.value || '' });
      }
    } catch (error) {
      console.error('[kfav][DEV] context menu action failed', error);
    }
    return;
  }
  if (info.menuItemId === DEV_COPY_FIELDS_LIST_MENU_ID) {
    try {
      const activeTab = await getActiveTabForDevAction(tab);
      if (!activeTab?.id) {
        console.debug('[kfav][DEV] no tab for fields list copy');
        return;
      }
      const tabId = activeTab.id;
      const tabUrl = activeTab.url || '';
      console.log('[kfav][DEV] copy fields start', { tabId, url: tabUrl });

      const fetched = await fetchFieldsListFromMainWorld(tabId);
      console.log('[kfav][DEV] copy fields fetched', {
        ok: Boolean(fetched?.ok),
        appId: fetched?.appId || '',
        url: fetched?.url || tabUrl,
        rowCount: Number(fetched?.rowCount || 0)
      });
      if (!fetched?.ok || !fetched?.text) {
        console.warn('[kfav][DEV] copy fields failed', {
          reason: fetched?.reason || 'unknown',
          tabId,
          url: fetched?.url || tabUrl
        });
        return;
      }
      let copied = null;
      try {
        copied = await chrome.tabs.sendMessage(tabId, {
          type: 'DEV_COPY_TEXT',
          text: fetched.text,
          label: 'fields-list'
        });
      } catch (sendError) {
        copied = {
          ok: false,
          reason: 'receiving_end_does_not_exist',
          detail: String(sendError?.message || sendError)
        };
      }
      if (!copied?.ok) {
        const fallback = await copyTextViaInjectedScript(tabId, fetched.text);
        if (!fallback?.ok) {
          console.warn('[kfav][DEV] copy fields failed', {
            reason: copied?.reason || fallback?.reason || 'unknown',
            tabId,
            url: fetched?.url || tabUrl
          });
          return;
        }
      }
      console.log('[kfav][DEV] copied fields list', {
        tabId,
        url: fetched?.url || tabUrl
      });
    } catch (error) {
      console.warn('[kfav][DEV] copy fields failed', {
        reason: String(error?.message || error),
        tabId: tab?.id || null,
        url: tab?.url || ''
      });
    }
    return;
  }
  if (info.menuItemId === DEV_COPY_APP_ID_MENU_ID) {
    try {
      const activeTab = await getActiveTabForDevAction(tab);
      if (!activeTab?.id) {
        console.warn('[kfav][DEV] copy appId failed', { reason: 'no_active_tab' });
        return;
      }
      const result = await fetchAppOrViewIdFromMainWorld(activeTab.id, 'appId');
      if (!result?.ok || !String(result?.value || '').trim()) {
        console.warn('[kfav][DEV] copy appId failed', {
          reason: result?.reason || 'app_id_not_found',
          tabId: activeTab.id,
          url: result?.url || activeTab.url || ''
        });
        return;
      }
      const copied = await copyDevTextToActiveTab(activeTab.id, String(result.value), 'appId');
      if (!copied?.ok) {
        console.warn('[kfav][DEV] copy appId failed', {
          reason: copied?.reason || 'clipboard_failed',
          tabId: activeTab.id,
          url: result?.url || activeTab.url || ''
        });
        return;
      }
      console.log('[kfav][DEV] copied appId');
    } catch (error) {
      console.warn('[kfav][DEV] copy appId failed', { reason: String(error?.message || error) });
    }
    return;
  }
  if (info.menuItemId === DEV_COPY_VIEW_ID_MENU_ID) {
    try {
      const activeTab = await getActiveTabForDevAction(tab);
      if (!activeTab?.id) {
        console.warn('[kfav][DEV] copy viewId failed', { reason: 'no_active_tab' });
        return;
      }
      const result = await fetchAppOrViewIdFromMainWorld(activeTab.id, 'view');
      if (!result?.ok) {
        console.warn('[kfav][DEV] copy viewId failed', {
          reason: result?.reason || 'not_kintone_page',
          tabId: activeTab.id,
          href: result?.url || activeTab.url || ''
        });
        return;
      }
      const value = String(result?.value || '').trim();
      if (!value) {
        console.warn('[kfav][DEV] copy viewId failed', {
          reason: 'view_id_not_found',
          tabId: activeTab.id,
          href: result?.url || activeTab.url || ''
        });
        return;
      }
      const copied = await copyDevTextToActiveTab(activeTab.id, value, 'viewId');
      if (!copied?.ok) {
        console.warn('[kfav][DEV] copy viewId failed', {
          reason: copied?.reason || 'clipboard_failed',
          tabId: activeTab.id,
          href: result?.url || activeTab.url || ''
        });
        return;
      }
      console.log('[kfav][DEV] copied viewId', { value, source: result?.source || 'unknown' });
    } catch (error) {
      console.warn('[kfav][DEV] copy viewId failed', {
        reason: String(error?.message || error),
        href: tab?.url || ''
      });
    }
    return;
  }
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
    if (msg?.type === 'PAGE_CONTEXT_UPDATED') {
      const tabId = sender?.tab?.id;
      if (tabId) {
        latestPageContextByTab.set(tabId, {
          href: String(msg?.payload?.href || sender?.tab?.url || ''),
          timestamp: Number(msg?.payload?.timestamp || Date.now()),
          trigger: String(msg?.payload?.trigger || 'unknown')
        });
      }
      sendResponse({ ok: true });
      return;
    }

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







