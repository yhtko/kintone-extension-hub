// options.js
const DEFAULT_PANE = 'general';
const menuButtons = Array.from(document.querySelectorAll('[data-pane-target]'));
const panes = Array.from(document.querySelectorAll('[data-pane]'));

function activatePane(target, options = {}) {
  const { updateHash = true } = options;
  const next = panes.some((pane) => pane.dataset.pane === target) ? target : DEFAULT_PANE;
  panes.forEach((pane) => {
    pane.classList.toggle('is-active', pane.dataset.pane === next);
  });
  menuButtons.forEach((btn) => {
    btn.classList.toggle('is-active', btn.dataset.paneTarget === next);
  });
  if (updateHash) {
    const newHash = `#${next}`;
    if (window.location.hash !== newHash) {
      let replaced = false;
      if (typeof history !== 'undefined' && typeof history.replaceState === 'function') {
        try {
          history.replaceState(null, '', newHash);
          replaced = true;
        } catch (_err) {
          // ignore fallback below
        }
      }
      if (!replaced) {
        try {
          window.location.hash = newHash;
        } catch (_err) {
          // ignore hash update failures
        }
      }
    }
  }
}

menuButtons.forEach((btn) => {
  btn.addEventListener('click', () => {
    activatePane(btn.dataset.paneTarget);
  });
});

window.addEventListener('hashchange', () => {
  activatePane(window.location.hash.slice(1), { updateHash: false });
});

activatePane(window.location.hash.slice(1) || DEFAULT_PANE, { updateHash: false });

const labelEl = document.getElementById('label');
const urlEl = document.getElementById('url');
const appIdEl = document.getElementById('appId');
const viewEl = document.getElementById('viewIdOrName');
const queryEl = document.getElementById('query');
const addBtn = document.getElementById('add');
const listEl = document.getElementById('list');
const shortcutToggleEl = document.getElementById('shortcut_visible');
const shortcutListEl = document.getElementById('shortcut_list');
const excelListEl = document.getElementById('excel_columns_list');
const excelClearBtn = document.getElementById('excel_columns_clear');
const EXCEL_COLUMN_PREF_KEY = 'kfavExcelColumns';
// ---- schedule form elements ----
const schLabelEl = document.getElementById('sch_label');
const schUrlEl = document.getElementById('sch_url');
const schAppIdEl = document.getElementById('sch_appId');
const schDateFieldEl = document.getElementById('sch_dateField');
const schTitleFieldEl = document.getElementById('sch_titleField');
const schExtraQueryEl = document.getElementById('sch_extraQuery');
const schLimitEl = document.getElementById('sch_limit');
const schSaveBtn = document.getElementById('sch_save');
const schResetBtn = document.getElementById('sch_reset');
const schDetectBtn = document.getElementById('sch_detect');
const dateFieldList = document.getElementById('dateFieldList');
const titleFieldList = document.getElementById('titleFieldList');
const schStatusEl = document.getElementById('sch_status');
const schListEl = document.getElementById('sch_list');
// ---- pinned record elements ----
const pinLabelEl = document.getElementById('pin_label');
const pinUrlEl = document.getElementById('pin_url');
const pinAppIdEl = document.getElementById('pin_appId');
const pinRecordIdEl = document.getElementById('pin_recordId');
const pinTitleFieldEl = document.getElementById('pin_titleField');
const pinNoteEl = document.getElementById('pin_note');
const pinSaveBtn = document.getElementById('pin_save');
const pinResetBtn = document.getElementById('pin_reset');
const pinListEl = document.getElementById('pin_list');
// 編集ダイアログ要素
const editDialog = document.getElementById('editDialog');
const editForm = document.getElementById('editForm');
const editLabelEl = document.getElementById('edit_label');
const editUrlEl = document.getElementById('edit_url');
const editAppIdEl = document.getElementById('edit_appId');
const editViewEl = document.getElementById('edit_viewIdOrName');
const editQueryEl = document.getElementById('edit_query');
const editSaveBtn = document.getElementById('editSave');
let editingId = null;
let scheduleEntries = [];
let scheduleEditingId = null;
let pinnedEntries = [];
let pinEditingId = null;
let shortcutEntries = [];
let shortcutsVisible = true;
let shortcutDraggingId = null;
let excelColumnPrefs = {};
const MAX_SHORTCUT_INITIAL_LENGTH = 2;

function normalizeShortcutInitialLocal(value) {
  if (value == null) return '';
  const str = String(value).trim();
  if (!str) return '';
  return Array.from(str).slice(0, MAX_SHORTCUT_INITIAL_LENGTH).join('');
}

function clearShortcutDropIndicators() {
  if (!shortcutListEl) return;
  const items = shortcutListEl.querySelectorAll('.shortcut-item.is-drop-before, .shortcut-item.is-drop-after');
  items.forEach((item) => {
    item.classList.remove('is-drop-before', 'is-drop-after');
  });
}

function beginShortcutDrag(ev, entryId, li) {
  shortcutDraggingId = entryId;
  clearShortcutDropIndicators();
  if (ev.dataTransfer) {
    try {
      ev.dataTransfer.effectAllowed = 'move';
      ev.dataTransfer.setData('text/plain', entryId);
    } catch (_err) {
      // ignore
    }
  }
  if (li) li.classList.add('dragging');
}

function endShortcutDrag(li) {
  if (li) li.classList.remove('dragging');
  shortcutDraggingId = null;
  clearShortcutDropIndicators();
}

function ensureShortcutListDnDHandlers() {
  if (!shortcutListEl || shortcutListEl.dataset.dndBound === '1') return;
  shortcutListEl.addEventListener('dragover', (ev) => {
    if (!shortcutDraggingId) return;
    ev.preventDefault();
    if (ev.dataTransfer) {
      try {
        ev.dataTransfer.dropEffect = 'move';
      } catch (_err) {
        // ignore
      }
    }
    clearShortcutDropIndicators();
    const target = ev.target.closest('.shortcut-item');
    if (!target || target.dataset.id === shortcutDraggingId) return;
    const rect = target.getBoundingClientRect();
    const isAfter = ev.clientY - rect.top > rect.height / 2;
    target.classList.add(isAfter ? 'is-drop-after' : 'is-drop-before');
  });
  shortcutListEl.addEventListener('dragleave', (ev) => {
    if (!shortcutListEl.contains(ev.relatedTarget)) {
      clearShortcutDropIndicators();
    }
  });
  shortcutListEl.addEventListener('drop', async (ev) => {
    if (!shortcutDraggingId) return;
    ev.preventDefault();
    const targetItem = ev.target.closest('.shortcut-item');
    let insertIndex = shortcutEntries.length;
    if (targetItem) {
      const targetId = targetItem.dataset.id;
      const targetIdx = shortcutEntries.findIndex((entry) => entry.id === targetId);
      if (targetIdx !== -1) {
        const rect = targetItem.getBoundingClientRect();
        const isAfter = ev.clientY - rect.top > rect.height / 2;
        insertIndex = targetIdx + (isAfter ? 1 : 0);
      }
    }
    const fromIdx = shortcutEntries.findIndex((entry) => entry.id === shortcutDraggingId);
    if (fromIdx === -1) {
      endShortcutDrag(null);
      return;
    }
    if (insertIndex < 0) insertIndex = 0;
    if (insertIndex > shortcutEntries.length) insertIndex = shortcutEntries.length;
    const [moved] = shortcutEntries.splice(fromIdx, 1);
    if (fromIdx < insertIndex) insertIndex -= 1;
    shortcutEntries.splice(insertIndex, 0, moved);
    await persistShortcutEntries();
    renderShortcutEntries();
    endShortcutDrag(null);
  });
  shortcutListEl.dataset.dndBound = '1';
}

function parseKintoneUrl(u) {
  try {
    const url = new URL(u);
    const m = url.pathname.match(/\/k\/(\d+)\//);
    const appId = m ? m[1] : '';
    const viewParam = url.searchParams.get('view') || '';
    const host = url.origin;
    return { host, appId, viewIdOrName: viewParam, url: u };
  } catch (_e) {
    return {};
  }
}

const SHORTCUT_TYPES = ['appTop', 'view', 'create'];
const SHORTCUT_TYPE_LABELS = {
  appTop: 'アプリトップ',
  view: 'ビュー',
  create: 'レコード新規'
};

function normalizeShortcutEntryLocal(entry, fallbackOrder = 0) {
  if (!entry || typeof entry !== 'object') return null;
  const type = SHORTCUT_TYPES.includes(entry.type) ? entry.type : 'appTop';
  const host = typeof entry.host === 'string' ? entry.host : '';
  const appId = entry.appId == null ? '' : String(entry.appId).trim();
  const viewIdOrName = entry.viewIdOrName == null ? '' : String(entry.viewIdOrName).trim();
  const label = entry.label == null ? '' : String(entry.label);
  const initial = normalizeShortcutInitialLocal(entry.initial);
  const order = Number.isFinite(entry.order) ? entry.order : fallbackOrder;
  return {
    id: entry.id || createId(),
    type,
    host,
    appId,
    viewIdOrName,
    label,
    initial,
    order
  };
}

function sortShortcutEntriesLocal(list) {
  return [...(list || [])].sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
}

function shortcutTypeLabel(type) {
  return SHORTCUT_TYPE_LABELS[type] || SHORTCUT_TYPE_LABELS.appTop;
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

function shortcutMetaText(entry) {
  const parts = [
    `種別: ${shortcutTypeLabel(entry.type)}`,
    `Host: ${entry.host || '-'}`,
    `App: ${entry.appId || '-'}`
  ];
  if (entry.type === 'view') {
    parts.push(`View: ${entry.viewIdOrName || '-'}`);
  }
  parts.push(`頭文字: ${normalizeShortcutInitialLocal(entry.initial) || '自動'}`);
  return parts.join(' / ');
}

async function persistShortcutEntries() {
  shortcutEntries = shortcutEntries.map((entry, index) => ({
    ...entry,
    order: index,
    initial: normalizeShortcutInitialLocal(entry.initial)
  }));
  const payload = shortcutEntries.map((entry) => ({
    id: entry.id,
    type: entry.type,
    host: entry.host,
    appId: entry.appId,
    viewIdOrName: entry.viewIdOrName,
    label: entry.label,
    initial: entry.initial,
    order: entry.order
  }));
  await chrome.storage.sync.set({ kfavShortcuts: payload });
}

function renderShortcutEntries() {
  if (!shortcutListEl) return;
  ensureShortcutListDnDHandlers();
  shortcutListEl.innerHTML = '';
  if (!shortcutEntries.length) {
    const empty = document.createElement('li');
    empty.className = 'shortcut-empty';
    empty.textContent = 'ショートカットはまだありません。';
    shortcutListEl.appendChild(empty);
    return;
  }

  shortcutEntries = sortShortcutEntriesLocal(shortcutEntries);
  shortcutEntries.forEach((entry, index) => {
    entry.order = index;
    entry.initial = normalizeShortcutInitialLocal(entry.initial);
    const li = document.createElement('li');
    li.className = 'shortcut-item';
    li.dataset.id = entry.id;
    li.draggable = true;

    const handle = document.createElement('span');
    handle.className = 'shortcut-handle';
    handle.textContent = '::';
    handle.title = 'ドラッグで並び替え';

    const info = document.createElement('div');
    info.className = 'shortcut-info';
    const title = document.createElement('div');
    title.className = 'shortcut-title';
    title.textContent = entry.label || shortcutTypeLabel(entry.type);
    const meta = document.createElement('div');
    meta.className = 'shortcut-meta';
    meta.textContent = shortcutMetaText(entry);
    info.appendChild(title);
    info.appendChild(meta);

    const ops = document.createElement('div');
    ops.className = 'shortcut-ops';
    const url = buildShortcutUrl(entry);
    const openA = document.createElement('a');
    openA.textContent = '開く';
    openA.className = 'btn';
    if (url) {
      openA.href = url;
      openA.target = '_blank';
      openA.rel = 'noopener';
      openA.title = url;
    } else {
      openA.href = '#';
      openA.setAttribute('aria-disabled', 'true');
      openA.classList.add('disabled');
      openA.title = '対象URLを生成できません';
      openA.addEventListener('click', (ev) => ev.preventDefault());
    }
    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.textContent = '編集';
    editBtn.addEventListener('click', async () => {
      const current = entry.label || '';
      const next = window.prompt('表示名を入力してください', current);
      if (next === null) return;
      entry.label = next.trim();
      await persistShortcutEntries();
      renderShortcutEntries();
    });
    const delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.textContent = '削除';
    delBtn.addEventListener('click', async () => {
      shortcutEntries = shortcutEntries.filter((item) => item.id !== entry.id);
      await persistShortcutEntries();
      renderShortcutEntries();
    });
    ops.appendChild(openA);
    ops.appendChild(editBtn);
    const initialBtn = document.createElement('button');
    initialBtn.type = 'button';
    initialBtn.textContent = '頭文字';
    initialBtn.addEventListener('click', async () => {
      const current = entry.initial || '';
      const next = window.prompt('ショートカットの頭文字（最大2文字。空欄で自動設定）', current);
      if (next === null) return;
      entry.initial = normalizeShortcutInitialLocal(next);
      await persistShortcutEntries();
      renderShortcutEntries();
    });
    ops.appendChild(initialBtn);
    ops.appendChild(delBtn);

    li.appendChild(handle);
    li.appendChild(info);
    li.appendChild(ops);

    const onDragStart = (ev) => {
      ev.stopPropagation();
      beginShortcutDrag(ev, entry.id, li);
    };
    const onDragEnd = () => {
      endShortcutDrag(li);
    };

    handle.draggable = true;
    handle.addEventListener('dragstart', onDragStart);
    handle.addEventListener('dragend', onDragEnd);

    li.addEventListener('dragstart', onDragStart);
    li.addEventListener('dragend', onDragEnd);

    shortcutListEl.appendChild(li);
  });
}

shortcutToggleEl?.addEventListener('change', async () => {
  shortcutsVisible = Boolean(shortcutToggleEl.checked);
  await chrome.storage.sync.set({ kfavShortcutsVisible: shortcutsVisible });
});

function createId() {
  return crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function setScheduleStatus(msg) {
  if (!schStatusEl) return;
  schStatusEl.textContent = msg || '';
}

function normalizeScheduleEntry(entry) {
  if (!entry || typeof entry !== 'object') return null;
  const limit = Math.max(1, Math.min(100, Number.parseInt(entry.limit ?? 20, 10) || 20));
  return {
    id: entry.id || createId(),
    label: entry.label || '',
    host: entry.host || '',
    appId: String(entry.appId || '').trim(),
    dateField: entry.dateField || '',
    titleField: entry.titleField || '',
    extraQuery: entry.extraQuery || '',
    limit
  };
}

function scheduleEntryLabel(entry) {
  const label = entry.label?.trim();
  if (label) return label;
  const app = entry.appId ? `App ${entry.appId}` : 'App 未指定';
  return `${app} @ ${entry.host}`;
}

async function saveScheduleEntries() {
  await chrome.storage.sync.set({ kfavSchedule: scheduleEntries.map(({ id, label, host, appId, dateField, titleField, extraQuery, limit }) => ({
    id,
    label,
    host,
    appId,
    dateField,
    titleField,
    extraQuery,
    limit
  })) });
}

function renderScheduleEntries() {
  if (!schListEl) return;
  schListEl.innerHTML = '';
  if (!scheduleEntries.length) {
    const li = document.createElement('li');
    li.className = 'muted';
    li.textContent = 'まだ登録されていません';
    schListEl.appendChild(li);
    return;
  }

  scheduleEntries.forEach((entry) => {
    const li = document.createElement('li');
    li.className = 'schedule-item';

    const info = document.createElement('div');
    info.className = 'schedule-info';
    const title = document.createElement('div');
    title.className = 'schedule-title';
    title.textContent = scheduleEntryLabel(entry);
    const meta = document.createElement('div');
    meta.className = 'schedule-meta';
    const detailRows = [
      ['Host', entry.host || '-'],
      ['App ID', entry.appId || '-'],
      ['Date', entry.dateField || '-'],
      ['Title', entry.titleField || '-'],
      ['Limit', String(entry.limit || '')]
    ];
    if (entry.extraQuery) detailRows.push(['Query', entry.extraQuery]);
    detailRows.forEach(([label, value]) => {
      const line = document.createElement('div');
      line.textContent = `${label}: ${value}`;
      meta.appendChild(line);
    });
    info.appendChild(title);
    info.appendChild(meta);

    const ops = document.createElement('div');
    ops.className = 'schedule-ops';
    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.textContent = '編集';
    editBtn.addEventListener('click', () => fillScheduleForm(entry));
    const delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.textContent = '削除';
    delBtn.addEventListener('click', async () => {
      scheduleEntries = scheduleEntries.filter((item) => item.id !== entry.id);
      await saveScheduleEntries();
      renderScheduleEntries();
    });
    ops.appendChild(editBtn);
    ops.appendChild(delBtn);


    li.appendChild(info);
    li.appendChild(ops);
    schListEl.appendChild(li);
  });
}

function resetScheduleForm() {
  scheduleEditingId = null;
  if (schLabelEl) schLabelEl.value = '';
  schUrlEl.value = '';
  schAppIdEl.value = '';
  schDateFieldEl.value = '';
  schTitleFieldEl.value = '';
  schExtraQueryEl.value = '';
  schLimitEl.value = '';
  setScheduleStatus('');
  updatePermStatus('');
  if (schSaveBtn) schSaveBtn.textContent = '追加 / 更新';
}

function fillScheduleForm(entry) {
  scheduleEditingId = entry.id;
  if (schLabelEl) schLabelEl.value = entry.label || '';
  schUrlEl.value = entry.host ? `${entry.host}/k/${entry.appId || ''}/` : '';
  schAppIdEl.value = entry.appId || '';
  schDateFieldEl.value = entry.dateField || '';
  schTitleFieldEl.value = entry.titleField || '';
  schExtraQueryEl.value = entry.extraQuery || '';
  schLimitEl.value = entry.limit ? String(entry.limit) : '';
  if (schSaveBtn) schSaveBtn.textContent = '更新';
  updatePermStatus(entry.host);
  setScheduleStatus('編集モードです。変更後「更新」を押してください。');
}

function normalizePinnedEntry(entry) {
  if (!entry || typeof entry !== 'object') return null;
  return {
    id: entry.id || createId(),
    label: entry.label || '',
    host: entry.host || '',
    appId: String(entry.appId || '').trim(),
    recordId: String(entry.recordId || '').trim(),
    titleField: entry.titleField || '',
    note: entry.note || ''
  };
}

function pinnedEntryLabel(entry) {
  const label = entry.label?.trim();
  if (label) return label;
  const app = entry.appId ? `App ${entry.appId}` : 'App 未指定';
  const record = entry.recordId ? `Record ${entry.recordId}` : 'Record 未指定';
  return `${app} / ${record}`;
}

async function savePinnedEntries() {
  await chrome.storage.sync.set({ kfavPins: pinnedEntries.map(({ id, label, host, appId, recordId, titleField, note }) => ({
    id,
    label,
    host,
    appId,
    recordId,
    titleField,
    note
  })) });
}

function renderPinnedEntries() {
  if (!pinListEl) return;
  pinListEl.innerHTML = '';
  if (!pinnedEntries.length) {
    const empty = document.createElement('li');
    empty.className = 'muted';
    empty.textContent = '登録済みのピンはありません。';
    pinListEl.appendChild(empty);
    return;
  }

  pinnedEntries.forEach((entry) => {
    const li = document.createElement('li');
    li.className = 'pin-item';

    const info = document.createElement('div');
    info.className = 'pin-info';
    const title = document.createElement('div');
    title.className = 'pin-title';
    title.textContent = pinnedEntryLabel(entry);
    const meta = document.createElement('div');
    meta.className = 'pin-meta';
    const details = [
      ['Host', entry.host || '-'],
      ['App ID', entry.appId || '-'],
      ['Record', entry.recordId || '-'],
      ['Title field', entry.titleField || '-']
    ];
    if (entry.note) details.push(['Note', entry.note]);
    details.forEach(([label, value]) => {
      const line = document.createElement('div');
      line.textContent = `${label}: ${value}`;
      meta.appendChild(line);
    });
    info.appendChild(title);
    info.appendChild(meta);

    const ops = document.createElement('div');
    ops.className = 'pin-ops';
    const openBtn = document.createElement('a');
    openBtn.textContent = '開く';
    openBtn.className = 'btn';
    if (entry.host && entry.appId && entry.recordId) {
      openBtn.href = `${entry.host}/k/${entry.appId}/show#record=${entry.recordId}`;
      openBtn.target = '_blank';
      openBtn.rel = 'noopener';
    } else {
      openBtn.href = '#';
      openBtn.setAttribute('aria-disabled', 'true');
      openBtn.classList.add('disabled');
      openBtn.addEventListener('click', (ev) => ev.preventDefault());
    }
    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.textContent = '編集';
    editBtn.addEventListener('click', () => fillPinnedForm(entry));
    const delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.textContent = '削除';
    delBtn.addEventListener('click', async () => {
      pinnedEntries = pinnedEntries.filter((item) => item.id !== entry.id);
      await savePinnedEntries();
      renderPinnedEntries();
    });
    ops.appendChild(openBtn);
    ops.appendChild(editBtn);
    ops.appendChild(delBtn);

    li.appendChild(info);
    li.appendChild(ops);
    pinListEl.appendChild(li);
  });
}

function resetPinnedForm() {
  pinEditingId = null;
  if (pinLabelEl) pinLabelEl.value = '';
  pinUrlEl.value = '';
  pinAppIdEl.value = '';
  pinRecordIdEl.value = '';
  pinTitleFieldEl.value = '';
  pinNoteEl.value = '';
  if (pinSaveBtn) pinSaveBtn.textContent = '追加 / 更新';
}

function fillPinnedForm(entry) {
  pinEditingId = entry.id;
  if (pinLabelEl) pinLabelEl.value = entry.label || '';
  pinUrlEl.value = entry.host && entry.appId && entry.recordId
    ? `${entry.host}/k/${entry.appId}/show#record=${entry.recordId}`
    : '';
  pinAppIdEl.value = entry.appId || '';
  pinRecordIdEl.value = entry.recordId || '';
  pinTitleFieldEl.value = entry.titleField || '';
  pinNoteEl.value = entry.note || '';
  if (pinSaveBtn) pinSaveBtn.textContent = '更新';
}
async function loadFavorites() {
  const { kintoneFavorites = [] } = await chrome.storage.sync.get('kintoneFavorites');
  return kintoneFavorites;
}
async function saveFavorites(items) {
  await chrome.storage.sync.set({ kintoneFavorites: items });
}

function sortItems(items) {
  return [...items].sort((a,b) => (b.pinned?1:0)-(a.pinned?1:0) || (a.order??0)-(b.order??0));
}

async function setBadgeTarget(id) {
  await chrome.storage.sync.set({ kfavBadgeTargetId: id });
}

async function getBadgeTarget() {
  const { kfavBadgeTargetId = null } = await chrome.storage.sync.get('kfavBadgeTargetId');
  return kfavBadgeTargetId;
}

async function render(items) {
  if (!listEl) return;
  const sorted = sortItems(items);
  const badgeTargetId = await getBadgeTarget();

  listEl.innerHTML = '';
  if (!sorted.length) {
    listEl.innerHTML = '<li style="color:#888;font-size:12px;">未登録</li>';
    return;
  }

  sorted.forEach((it, idx) => {
    if (typeof it.order !== 'number') it.order = idx;

    const li = document.createElement('li');
    li.className = 'item';
    li.draggable = true;
    li.dataset.id = it.id;

    const handle = document.createElement('div');
    handle.className = 'drag-handle';
    handle.textContent = '⋮⋮';

    const left = document.createElement('div');
    left.className = 'meta';
    const title = document.createElement('div');
    title.className = 'title';
    title.textContent = it.label || '(no label)';
    const meta = document.createElement('div');
    meta.className = 'meta';
    meta.innerHTML = [
      `URL: ${it.url}`,
      `App: ${it.appId || '-'}`,
      `View: ${it.viewIdOrName || '-'}`,
      `Query: ${it.query || '-'}`,
      `Host: ${it.host}`
    ].join('<br/>');
    left.appendChild(title);
    left.appendChild(meta);

    const right = document.createElement('div');
    right.className = 'ops';

    // バッジ対象ラジオ
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = 'badgeTarget';
    radio.className = 'badge-radio';
    radio.checked = badgeTargetId ? (badgeTargetId === it.id) : (sorted[0]?.id === it.id);
    radio.title = 'このウォッチリストをバッジ対象にする';
    radio.addEventListener('change', () => setBadgeTarget(it.id));

    // ピン切替
    const pinBtn = document.createElement('button');
    pinBtn.className = 'pin';
    pinBtn.textContent = it.pinned ? '📌 解除' : '📌 固定';
    pinBtn.addEventListener('click', async () => {
      const all = await loadFavorites();
      const me = all.find(x => x.id === it.id);
      me.pinned = !me.pinned;
      await saveFavorites(all);
      render(all);
    });

    // 開く・削除
    const openA = document.createElement('a');
    openA.textContent = '開く';
    openA.className = 'btn';
    openA.href = it.url;
    openA.target = '_blank';
    openA.rel = 'noopener';
    // 編集ボタン
    const editBtn = document.createElement('button');
    editBtn.textContent = '編集';
    editBtn.addEventListener('click', () => openEdit(it));

    const delBtn = document.createElement('button');
    delBtn.textContent = '削除';
    delBtn.addEventListener('click', async () => {
      const next = (await loadFavorites()).filter(x => x.id !== it.id);
      await saveFavorites(next);
      render(next);
    });

    right.appendChild(radio);
    right.appendChild(pinBtn);
    right.appendChild(openA);
    right.appendChild(editBtn);
    right.appendChild(delBtn);

    // D&D
    li.addEventListener('dragstart', e => {
      e.dataTransfer.setData('text/plain', it.id);
      li.classList.add('dragging');
    });
    li.addEventListener('dragend', () => li.classList.remove('dragging'));
    li.addEventListener('dragover', e => e.preventDefault());
    li.addEventListener('drop', async e => {
      e.preventDefault();
      const fromId = e.dataTransfer.getData('text/plain');
      const toId = it.id;
      if (fromId === toId) return;
      const all = await loadFavorites();
      const fromIdx = all.findIndex(x => x.id === fromId);
      const toIdx = all.findIndex(x => x.id === toId);
      const [moved] = all.splice(fromIdx, 1);
      all.splice(toIdx, 0, moved);
      all.forEach((x, i) => x.order = i);
      await saveFavorites(all);
      render(all);
    });

    li.appendChild(handle);
    li.appendChild(left);
    li.appendChild(right);
    listEl.appendChild(li);
  });
}

// ---- 編集処理（renderの外に1回だけ定義）----
function openEdit(item) {
  editingId = item.id;
  editLabelEl.value = item.label || '';
  editUrlEl.value = item.url || '';
  editAppIdEl.value = item.appId || '';
  editViewEl.value = item.viewIdOrName || '';
  editQueryEl.value = item.query || '';
  if (typeof editDialog.showModal === 'function') {
    editDialog.showModal();
  } else {
    alert('このブラウザはdialog要素に対応していません');
  }
}
editDialog.addEventListener('close', () => { /* noop */ });
editSaveBtn.addEventListener('click', async (e) => {
  e.preventDefault();
  if (!editingId) { editDialog.close(); return; }
  const newLabel = editLabelEl.value.trim();
  const newUrl = editUrlEl.value.trim();
  let { host, appId, viewIdOrName } = parseKintoneUrl(newUrl);
  if (editAppIdEl.value.trim()) appId = editAppIdEl.value.trim();
  if (editViewEl.value.trim()) viewIdOrName = editViewEl.value.trim();
  const newQuery = editQueryEl.value.trim() || '';
  if (!newUrl || !host) { alert('正しい kintone URL を入力してください'); return; }
  const all = await loadFavorites();
  const idx = all.findIndex(x => x.id === editingId);
  if (idx === -1) { editDialog.close(); return; }
  if (all.some((x, i) => i !== idx && x.url === newUrl)) {
    alert('このURLは既に登録済みです'); return;
  }
  const old = all[idx];
  all[idx] = { ...old, label: newLabel, url: newUrl, host, appId: appId||'', viewIdOrName: viewIdOrName||'', query: newQuery };
  await saveFavorites(all);
  await render(all);
  editingId = null;
  editDialog.close();
});

addBtn.addEventListener('click', async () => {
  const label = labelEl.value.trim();
  const url = urlEl.value.trim();
  let { host, appId, viewIdOrName } = parseKintoneUrl(url);

  if (appIdEl.value.trim()) appId = appIdEl.value.trim();
  if (viewEl.value.trim()) viewIdOrName = viewEl.value.trim();
  const query = queryEl.value.trim() || '';

  if (!url || !host) {
    alert('正しい kintone URL を入力してください');
    return;
  }

  const list = await loadFavorites();
  if (list.some(x => x.url === url)) {
    alert('このURLは既に登録済みです');
    return;
  }

  const item = {
    id: createId(),
    label,
    url,
    host,
    appId: appId || '',
    viewIdOrName: viewIdOrName || '',
    query: query || '',
    initial: '',
    order: list.length,
    pinned: false
  };

  const next = [...list, item];
  await saveFavorites(next);
  await render(next);

  // 入力クリア（ラベルは残す）
  urlEl.value = '';
  appIdEl.value = '';
  viewEl.value = '';
  queryEl.value = '';
});

// 初期表示
(async () => {
  const items = await loadFavorites();
  await render(items);

  const shortcutStored = await chrome.storage.sync.get(['kfavShortcuts', 'kfavShortcutsVisible']);
  const rawShortcuts = shortcutStored.kfavShortcuts;
  const shortcutArr = Array.isArray(rawShortcuts) ? rawShortcuts : rawShortcuts ? [rawShortcuts] : [];
  shortcutEntries = shortcutArr.map((item, idx) => normalizeShortcutEntryLocal(item, idx)).filter(Boolean);
  shortcutEntries = sortShortcutEntriesLocal(shortcutEntries);
  const visibleFlag = shortcutStored.kfavShortcutsVisible;
  shortcutsVisible = typeof visibleFlag === 'boolean' ? visibleFlag : true;
  if (shortcutToggleEl) shortcutToggleEl.checked = shortcutsVisible;
  renderShortcutEntries();

  const stored = await chrome.storage.sync.get('kfavSchedule');
  const raw = stored.kfavSchedule;
  const arr = Array.isArray(raw) ? raw : raw ? [raw] : [];
  scheduleEntries = arr.map(normalizeScheduleEntry).filter(Boolean);
  renderScheduleEntries();
  resetScheduleForm();

  const pinsStored = await chrome.storage.sync.get('kfavPins');
  const pinRaw = pinsStored.kfavPins;
  const pinArr = Array.isArray(pinRaw) ? pinRaw : pinRaw ? [pinRaw] : [];
  pinnedEntries = pinArr.map(normalizePinnedEntry).filter(Boolean);
  renderPinnedEntries();
  resetPinnedForm();

  await loadExcelColumnPrefs();
})();

if (chrome?.storage?.onChanged) {
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area === 'sync' && Object.prototype.hasOwnProperty.call(changes, 'kfavPins')) {
      const next = changes.kfavPins.newValue;
      const arr = Array.isArray(next) ? next : next ? [next] : [];
      pinnedEntries = arr.map(normalizePinnedEntry).filter(Boolean);
      renderPinnedEntries();
    }
    if (area === 'sync' && Object.prototype.hasOwnProperty.call(changes, 'kfavSchedule')) {
      const next = changes.kfavSchedule.newValue;
      const arr = Array.isArray(next) ? next : next ? [next] : [];
      scheduleEntries = arr.map(normalizeScheduleEntry).filter(Boolean);
      renderScheduleEntries();
    }
    if (area === 'sync' && Object.prototype.hasOwnProperty.call(changes, EXCEL_COLUMN_PREF_KEY)) {
      excelColumnPrefs = normalizeExcelColumnPrefs(changes[EXCEL_COLUMN_PREF_KEY].newValue);
      renderExcelColumnPrefs();
    }
  });
}

// ---- schedule save ----
schUrlEl?.addEventListener('change', async () => {
  // URLからhost/appIdを補完
  const raw = schUrlEl.value.trim();
  const parsed = parseKintoneUrl(raw);
  if (parsed.host && !schAppIdEl.value) schAppIdEl.value = parsed.appId || '';
  const host = parsed.host || hostFromUrl(raw);
  await updatePermStatus(host);
  setScheduleStatus('');
});

schResetBtn?.addEventListener('click', () => {
  resetScheduleForm();
});
schSaveBtn?.addEventListener('click', async () => {
  const label = schLabelEl?.value.trim() || '';
  const url = schUrlEl.value.trim();
  const parsed = parseKintoneUrl(url);
  const host = parsed.host || hostFromUrl(url);
  const appId = (schAppIdEl.value || parsed.appId || '').toString().trim();
  const dateField = schDateFieldEl.value.trim();
  const titleField = schTitleFieldEl.value.trim();
  const extraQuery = schExtraQueryEl.value.trim();
  const limit = Math.max(1, Math.min(100, parseInt(schLimitEl.value || '20', 10)));

  if (!host || !appId || !dateField) {
    alert('host / App ID / 日付フィールドは必須です');
    return;
  }

  const entry = normalizeScheduleEntry({
    id: scheduleEditingId || createId(),
    label,
    host,
    appId,
    dateField,
    titleField,
    extraQuery,
    limit
  });
  const idx = scheduleEntries.findIndex((item) => item.id === entry.id);
  if (idx >= 0) scheduleEntries[idx] = entry;
  else scheduleEntries.push(entry);

  await saveScheduleEntries();
  renderScheduleEntries();
  setScheduleStatus(scheduleEditingId ? '更新しました。' : '追加しました。');
  resetScheduleForm();
});

// ---- kintoneフィールド検出 ----
function hostFromUrl(u){
  try {
    return new URL(u).origin;
  } catch {
    return '';
  }
}

schDetectBtn?.addEventListener('click', async () => {
  const hostHint = hostFromUrl(schUrlEl.value || '');
  const appId = schAppIdEl.value || parseKintoneUrl(schUrlEl.value || '').appId || '';
  if (!hostHint) { setScheduleStatus('先に kintone URL を入力し、ホストを特定してください。'); return; }
  if (!appId) { alert('App ID を入力してから実行してください'); return; }
  if (!(await hasHostPermission(hostHint))) {
    setScheduleStatus('このホストの許可が必要です。「このホストを許可」を押してください。');
    return;
  }
  try {
    setScheduleStatus('フィールド情報を取得中…');
    const res = await chrome.runtime.sendMessage({
      type: 'RUN_IN_KINTONE',
      host: hostHint,
      forward: { type: 'LIST_FIELDS', payload: { appId } }
    });
    if (!res?.ok) throw new Error(res?.error || 'LIST_FIELDS failed');
    const fields = res.fields || [];
    const dateCandidates = fields.filter(f => f.type === 'DATE' || f.type === 'DATETIME');
    dateFieldList.innerHTML = '';
    dateCandidates.forEach(f => {
      const opt = document.createElement('option');
      opt.value = f.code;
      opt.label = `${f.label} (${f.code})`;
      dateFieldList.appendChild(opt);
    });
    const titleCandidates = fields.filter(f => ['SINGLE_LINE_TEXT','MULTI_LINE_TEXT','RICH_TEXT','LINK'].includes(f.type));
    titleFieldList.innerHTML = '';
    titleCandidates.forEach(f => {
      const opt = document.createElement('option');
      opt.value = f.code;
      opt.label = `${f.label} (${f.code})`;
      titleFieldList.appendChild(opt);
    });
    setScheduleStatus('フィールド候補を更新しました（入力欄の候補から選択できます）');
  } catch (e) {
    setScheduleStatus('フィールド検出に失敗：' + String(e?.message || e));
  }
});

// ---- Excel column prefs ----
function normalizeExcelColumnPrefs(raw) {
  if (!raw || typeof raw !== 'object') return {};
  const next = {};
  Object.entries(raw).forEach(([key, value]) => {
    if (!value || typeof value !== 'object') return;
    next[key] = {
      order: Array.isArray(value.order) ? value.order.filter(Boolean) : [],
      appId: value.appId || '',
      appName: value.appName || '',
      viewName: value.viewName || '',
      viewKey: value.viewKey || '',
      savedAt: value.savedAt || 0,
      widths: value.widths && typeof value.widths === 'object' ? { ...value.widths } : {}
    };
  });
  return next;
}

function formatPrefDate(ts) {
  if (!ts) return '-';
  const d = new Date(ts);
  if (Number.isNaN(d.getTime())) return '-';
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  const hh = String(d.getHours()).padStart(2, '0');
  const mi = String(d.getMinutes()).padStart(2, '0');
  return `${yyyy}/${mm}/${dd} ${hh}:${mi}`;
}

function renderExcelColumnPrefs() {
  if (!excelListEl) return;
  excelListEl.innerHTML = '';
  const entries = Object.entries(excelColumnPrefs).map(([key, value]) => ({
    key,
    appId: value.appId || '',
    appName: value.appName || '',
    viewName: value.viewName || '',
    viewKey: value.viewKey || '',
    savedAt: value.savedAt || 0,
    count: Array.isArray(value.order) ? value.order.length : 0,
    widthCount: value.widths ? Object.keys(value.widths).length : 0
  })).sort((a, b) => (b.savedAt || 0) - (a.savedAt || 0));

  if (!entries.length) {
    const li = document.createElement('li');
    li.className = 'excel-columns-empty';
    li.textContent = '保存済みの列順はありません。';
    excelListEl.appendChild(li);
    return;
  }

  entries.forEach((entry) => {
    const li = document.createElement('li');
    li.className = 'excel-columns-item';

    const info = document.createElement('div');
    info.className = 'excel-columns-info';
    const title = document.createElement('div');
    title.className = 'excel-columns-title';
    const appLabel = entry.appName || (entry.appId ? `App ${entry.appId}` : 'App 未指定');
    const viewLabel = entry.viewName || 'ビュー未指定';
    title.textContent = `${appLabel} / ${viewLabel}`;
    const meta = document.createElement('div');
    meta.className = 'excel-columns-meta';
    const count = document.createElement('span');
    count.textContent = `列数: ${entry.count}`;
    const saved = document.createElement('span');
    saved.textContent = `保存: ${formatPrefDate(entry.savedAt)}`;
    meta.appendChild(count);
    if (entry.widthCount > 0) {
      const widthInfo = document.createElement('span');
      widthInfo.textContent = `幅設定: ${entry.widthCount}列`;
      meta.appendChild(widthInfo);
    }
    meta.appendChild(saved);
    info.appendChild(title);
    info.appendChild(meta);

    const actions = document.createElement('div');
    actions.className = 'excel-columns-actions';
    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.textContent = '削除';
    removeBtn.addEventListener('click', () => { removeExcelPref(entry.key); });
    actions.appendChild(removeBtn);

    li.appendChild(info);
    li.appendChild(actions);
    excelListEl.appendChild(li);
  });
}

async function loadExcelColumnPrefs() {
  if (!excelListEl) return;
  const stored = await chrome.storage.sync.get(EXCEL_COLUMN_PREF_KEY);
  excelColumnPrefs = normalizeExcelColumnPrefs(stored[EXCEL_COLUMN_PREF_KEY]);
  renderExcelColumnPrefs();
}

async function removeExcelPref(key) {
  if (!key || !excelColumnPrefs[key]) return;
  const next = { ...excelColumnPrefs };
  delete next[key];
  excelColumnPrefs = next;
  if (Object.keys(next).length) {
    await chrome.storage.sync.set({ [EXCEL_COLUMN_PREF_KEY]: next });
  } else {
    await chrome.storage.sync.remove(EXCEL_COLUMN_PREF_KEY);
  }
  renderExcelColumnPrefs();
}

excelClearBtn?.addEventListener('click', async () => {
  if (!Object.keys(excelColumnPrefs).length) return;
  if (!window.confirm('保存済みの列順をすべて削除します。よろしいですか？')) return;
  excelColumnPrefs = {};
  await chrome.storage.sync.remove(EXCEL_COLUMN_PREF_KEY);
  renderExcelColumnPrefs();
});

// ---- Host permission helpers ----
function originPatternFor(origin){ if(!origin) return null; return origin.endsWith('/') ? origin+'*' : origin+'/*'; }
async function hasHostPermission(origin){
  const pat = originPatternFor(origin); if(!pat) return false;
  return await chrome.permissions.contains({ origins: [pat] });
}
async function requestHostPermission(origin){
  const pat = originPatternFor(origin); if(!pat) return false;
  // ここは“クリック直後”で呼ぶこと（ユーザー操作扱い）
  return await chrome.permissions.request({ origins: [pat] });
}
async function updatePermStatus(origin){
  const el = document.getElementById('sch_permStatus'); if(!el) return false;
  const granted = await hasHostPermission(origin);
  el.textContent = granted ? '✅ 許可済み' : '❌ 未許可';
  return granted;
}
document.getElementById('sch_grantHost')?.addEventListener('click', async () => {
  const host = hostFromUrl(schUrlEl.value||'');
  if (!host) { alert('先に kintone URL を入力してください'); return; }
  const ok = await requestHostPermission(host);   // ← クリック直下で実行
  await updatePermStatus(host);
  if (!ok) { alert('許可がキャンセルされました'); return; }
  setScheduleStatus('ホストの許可が完了しました。必要に応じてフィールド検出を実行してください。');
});

// ---- pinned events ----
pinResetBtn?.addEventListener('click', () => {
  resetPinnedForm();
});
pinSaveBtn?.addEventListener('click', async () => {
  const label = pinLabelEl?.value.trim() || '';
  const url = pinUrlEl.value.trim();
  const parsed = parseKintoneUrl(url);
  const host = parsed.host || hostFromUrl(url);
  const appId = (pinAppIdEl.value || parsed.appId || '').toString().trim();
  const recordId = (pinRecordIdEl.value || parsed.recordId || '').toString().trim();
  const titleField = pinTitleFieldEl.value.trim();
  const note = pinNoteEl.value.trim();

  if (!host || !appId || !recordId) {
    alert('host / App ID / レコードID は必須です');
    return;
  }

  const entry = normalizePinnedEntry({
    id: pinEditingId || createId(),
    label,
    host,
    appId,
    recordId,
    titleField,
    note
  });
  const idx = pinnedEntries.findIndex((item) => item.id === entry.id);
  if (idx >= 0) pinnedEntries[idx] = entry;
  else pinnedEntries.push(entry);

  await savePinnedEntries();
  renderPinnedEntries();
  resetPinnedForm();
});
pinUrlEl?.addEventListener('change', () => {
  const raw = pinUrlEl.value.trim();
  const parsed = parseKintoneUrl(raw);
  if (parsed.appId && !pinAppIdEl.value) pinAppIdEl.value = parsed.appId;
  if (parsed.recordId && !pinRecordIdEl.value) pinRecordIdEl.value = parsed.recordId;
});
















