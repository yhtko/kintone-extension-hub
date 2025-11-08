'use strict';

const FAVORITES_KEY = 'kintoneFavorites';
const SCHEDULE_KEY = 'kfavSchedule';
const PIN_KEY = 'kfavPins';
const SHORTCUT_KEY = 'kfavShortcuts';
const SHORTCUT_VISIBLE_KEY = 'kfavShortcutsVisible';
const MAX_SHORTCUT_INITIAL_LENGTH = 2;

function normalizeShortcutInitial(value) {
  if (value == null) return '';
  const str = String(value).trim();
  if (!str) return '';
  return Array.from(str).slice(0, MAX_SHORTCUT_INITIAL_LENGTH).join('');
}

export function createId() {
  if (globalThis.crypto?.randomUUID) {
    try { return globalThis.crypto.randomUUID(); } catch (_) { /* fall back */ }
  }
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

export async function loadFavorites() {
  const stored = await chrome.storage.sync.get(FAVORITES_KEY);
  return stored[FAVORITES_KEY] || [];
}

export async function saveFavorites(items) {
  await chrome.storage.sync.set({ [FAVORITES_KEY]: items });
}

export function sortFavorites(items) {
  return [...items].sort((a, b) => (b.pinned ? 1 : 0) - (a.pinned ? 1 : 0) || (a.order ?? 0) - (b.order ?? 0));
}

export async function loadScheduleConfig() {
  const stored = await chrome.storage.sync.get(SCHEDULE_KEY);
  const raw = stored[SCHEDULE_KEY];
  if (!raw) return [];
  if (Array.isArray(raw)) {
    return raw
      .filter((item) => item && typeof item === 'object')
      .map((item) => ({ id: item.id || createId(), ...item }));
  }
  if (raw && typeof raw === 'object') {
    const entry = { id: raw.id || createId(), ...raw };
    return [entry];
  }
  return [];
}

export async function saveScheduleConfig(configs) {
  await chrome.storage.sync.set({ [SCHEDULE_KEY]: configs });
}

export function originPatternFor(origin) {
  if (!origin) return null;
  return origin.endsWith('/') ? `${origin}*` : `${origin}/*`;
}

export async function hasHostPermission(origin) {
  const pattern = originPatternFor(origin);
  if (!pattern) return false;
  return await chrome.permissions.contains({ origins: [pattern] });
}

export async function ensureHostPermission(origin) {
  const pattern = originPatternFor(origin);
  if (!pattern) return false;
  if (await chrome.permissions.contains({ origins: [pattern] })) return true;
  return await chrome.permissions.request({ origins: [pattern] });
}

export async function sendRunInKintone(host, forward) {
  if (!host) throw new Error('host is required');
  if (!forward?.type) throw new Error('forward.type is required');
  return await chrome.runtime.sendMessage({
    type: 'RUN_IN_KINTONE',
    host,
    forward
  });
}

export async function sendCountBulk(host, items) {
  return await sendRunInKintone(host, { type: 'COUNT_BULK', payload: { items } });
}

export async function getActiveTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab;
}

export function isKintoneUrl(url) {
  try {
    const hostname = new URL(url).hostname;
    return /\.kintone(?:-dev)?\.com$/.test(hostname) || /\.cybozu\.com$/.test(hostname);
  } catch (_e) {
    return false;
  }
}

export function parseKintoneUrl(u) {
  try {
    const url = new URL(u);
    const match = url.pathname.match(/\/k\/(\d+)\//);
    const appId = match ? match[1] : '';
    const viewParam = url.searchParams.get('view') || '';
    const host = url.origin;
    let recordId = '';
    if (url.hash) {
      const hash = url.hash;
      const recMatch = hash.match(/record=(\d+)/);
      if (recMatch) recordId = recMatch[1];
    }
    return { host, appId, viewIdOrName: viewParam, recordId, url: u };
  } catch (_e) {
    return {};
  }
}

export async function loadPinnedRecords() {
  const stored = await chrome.storage.sync.get(PIN_KEY);
  const raw = stored[PIN_KEY];
  if (!raw) return [];
  if (Array.isArray(raw)) {
    return raw
      .filter((item) => item && typeof item === 'object')
      .map((item) => ({ id: item.id || createId(), ...item }));
  }
  if (raw && typeof raw === 'object') {
    return [{ id: raw.id || createId(), ...raw }];
  }
  return [];
}

export async function savePinnedRecords(entries) {
  await chrome.storage.sync.set({ [PIN_KEY]: entries });
}

const SHORTCUT_TYPES = ['appTop', 'view', 'create'];

function normalizeShortcutEntry(item, fallbackOrder) {
  if (!item || typeof item !== 'object') {
    return null;
  }
  const orderValue = typeof item.order === 'number' ? item.order : fallbackOrder;
  const type = SHORTCUT_TYPES.includes(item.type) ? item.type : 'appTop';
  const appIdRaw = item.appId == null ? '' : String(item.appId).trim();
  const viewRaw = item.viewIdOrName == null ? '' : String(item.viewIdOrName).trim();
  const labelRaw = item.label == null ? '' : String(item.label);
  const initialRaw = normalizeShortcutInitial(item.initial);
  return {
    id: item.id || createId(),
    type,
    host: typeof item.host === 'string' ? item.host : '',
    appId: appIdRaw,
    viewIdOrName: viewRaw,
    label: labelRaw,
    initial: initialRaw,
    order: orderValue
  };
}

export async function loadShortcuts() {
  const stored = await chrome.storage.sync.get(SHORTCUT_KEY);
  const raw = stored[SHORTCUT_KEY];
  if (!raw) return [];
  const list = Array.isArray(raw) ? raw : [raw];
  return list
    .map((item, index) => normalizeShortcutEntry(item, index))
    .filter((entry) => entry);
}

export async function saveShortcuts(entries) {
  const payload = (entries || []).map((item, index) => ({
    id: item.id || createId(),
    type: SHORTCUT_TYPES.includes(item.type) ? item.type : 'appTop',
    host: typeof item.host === 'string' ? item.host : '',
    appId: item.appId == null ? '' : String(item.appId).trim(),
    viewIdOrName: item.viewIdOrName == null ? '' : String(item.viewIdOrName).trim(),
    label: item.label == null ? '' : String(item.label),
    initial: normalizeShortcutInitial(item.initial),
    order: typeof item.order === 'number' ? item.order : index
  }));
  await chrome.storage.sync.set({ [SHORTCUT_KEY]: payload });
}

export function sortShortcuts(entries) {
  return [...(entries || [])].sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
}

export async function loadShortcutVisibility() {
  const stored = await chrome.storage.sync.get(SHORTCUT_VISIBLE_KEY);
  const flag = stored[SHORTCUT_VISIBLE_KEY];
  if (typeof flag === 'boolean') return flag;
  return true;
}

export async function saveShortcutVisibility(visible) {
  await chrome.storage.sync.set({ [SHORTCUT_VISIBLE_KEY]: Boolean(visible) });
}
