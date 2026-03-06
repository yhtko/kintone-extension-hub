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
const hostPermInputEl = document.getElementById('perm_host_input');
const hostPermRequestBtn = document.getElementById('perm_request');
const hostPermCheckBtn = document.getElementById('perm_check');
const hostPermStatusEl = document.getElementById('perm_status');
const iconEl = document.getElementById('icon');
const iconColorEl = document.getElementById('iconColor');
const categoryEl = document.getElementById('category');
const addBtn = document.getElementById('add');
const listEl = document.getElementById('list');
const shortcutToggleEl = document.getElementById('shortcut_visible');
const shortcutListEl = document.getElementById('shortcut_list');
const excelListEl = document.getElementById('excel_columns_list');
const excelClearBtn = document.getElementById('excel_columns_clear');
const excelModeInputs = Array.from(document.querySelectorAll('input[name="excel_overlay_mode"]'));
const EXCEL_COLUMN_PREF_KEY = 'kfavExcelColumns';
const EXCEL_OVERLAY_MODE_KEY = 'kfavExcelOverlayMode';
const DEFAULT_EXCEL_OVERLAY_MODE = 'edit';
const EXCEL_OVERLAY_MODE_VALUES = ['off', 'view', 'edit'];
const PIN_VISIBLE_KEY = 'kfavRecordPinsVisible';
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
const pinVisibleToggleEl = document.getElementById('pin_visible');
// 編集ダイアログ要素
const editDialog = document.getElementById('editDialog');
const editForm = document.getElementById('editForm');
const editLabelEl = document.getElementById('edit_label');
const editUrlEl = document.getElementById('edit_url');
const editAppIdEl = document.getElementById('edit_appId');
const editViewEl = document.getElementById('edit_viewIdOrName');
const editQueryEl = document.getElementById('edit_query');
const editIconEl = document.getElementById('edit_icon');
const editIconColorEl = document.getElementById('edit_iconColor');
const editCategoryEl = document.getElementById('edit_category');
const iconPreviewEl = document.getElementById('iconPreview');
const editIconPreviewEl = document.getElementById('editIconPreview');
const categorySuggestionsEl = document.getElementById('categorySuggestions');
const editCancelBtn = document.getElementById('editCancel');
const editCloseXBtn = document.getElementById('editCloseX');
const editSaveBtn = document.getElementById('editSave');
const pinEditDialog = document.getElementById('pinEditDialog');
const pinEditForm = document.getElementById('pinEditForm');
const pinEditLabelEl = document.getElementById('pin_edit_label');
const pinEditUrlEl = document.getElementById('pin_edit_url');
const pinEditAppIdEl = document.getElementById('pin_edit_appId');
const pinEditRecordIdEl = document.getElementById('pin_edit_recordId');
const pinEditTitleFieldEl = document.getElementById('pin_edit_titleField');
const pinEditNoteEl = document.getElementById('pin_edit_note');
const pinEditCancelBtn = document.getElementById('pinEditCancel');
const pinEditCloseXBtn = document.getElementById('pinEditCloseX');
const pinEditSaveBtn = document.getElementById('pinEditSave');
editForm?.addEventListener('click', (event) => event.stopPropagation());
pinEditForm?.addEventListener('click', (event) => event.stopPropagation());
let editingId = null;
let pinnedEntries = [];
let pinModalEditingId = null;
let shortcutEntries = [];
let shortcutsVisible = true;
let shortcutDraggingId = null;
let excelColumnPrefs = {};
const ensuredOrigins = new Set();
const ICON_CHOICES = [
  'clipboard', 'file-text', 'package', 'box', 'truck', 'factory', 'wrench', 'calendar',
  'list-checks', 'search', 'chart-bar', 'receipt', 'users', 'settings', 'bookmark', 'star'
];
const DEFAULT_ICON = 'file-text';
const ICON_COLOR_CHOICES = ['gray', 'blue', 'green', 'orange', 'red', 'purple'];
const ICON_COLOR_LABELS = {
  gray: 'グレー',
  blue: 'ブルー',
  green: 'グリーン',
  orange: 'オレンジ',
  red: 'レッド',
  purple: 'パープル'
};
const DEFAULT_ICON_COLOR = 'gray';
const DEFAULT_CATEGORY = 'その他';
const MAX_SHORTCUT_INITIAL_LENGTH = 2;

function normalizeShortcutInitialLocal(value) {
  if (value == null) return '';
  const str = String(value).trim();
  if (!str) return '';
  return Array.from(str).slice(0, MAX_SHORTCUT_INITIAL_LENGTH).join('');
}

function sanitizeIcon(value) {
  const name = (value || '').trim();
  return ICON_CHOICES.includes(name) ? name : DEFAULT_ICON;
}

function sanitizeIconColor(value) {
  const color = String(value || '').trim().toLowerCase();
  return ICON_COLOR_CHOICES.includes(color) ? color : DEFAULT_ICON_COLOR;
}

function sanitizeShortcutIcon(value) {
  return sanitizeIcon(value);
}

function sanitizeShortcutIconColor(value) {
  return sanitizeIconColor(value);
}

function sanitizeCategory(value) {
  const trimmed = (value || '').trim();
  return trimmed;
}

function getCategoryLabel(value) {
  const name = sanitizeCategory(value);
  return name || DEFAULT_CATEGORY;
}

function normalizeFavoriteEntryLocal(item, fallbackOrder) {
  return {
    ...item,
    icon: sanitizeIcon(item.icon),
    iconColor: sanitizeIconColor(item.iconColor),
    category: sanitizeCategory(item.category),
    order: typeof item.order === 'number' ? item.order : fallbackOrder
  };
}

function toPascalIconName(name) {
  return String(name || '')
    .trim()
    .split(/[-_\s]+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
    .join('');
}

function iconToSvgLocal(name, size = 18) {
  const lucide = globalThis.lucide;
  const icons = lucide?.icons;
  if (!icons) return '';
  const kebab = sanitizeIcon(name);
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

function renderLucideIconsLocal(root = document) {
  const nodes = root.querySelectorAll('.lc[data-icon]');
  nodes.forEach((el) => {
    const iconName = sanitizeIcon(el.dataset.icon);
    el.dataset.icon = iconName;
    el.innerHTML = iconToSvgLocal(iconName, 18);
  });
}

function renderIconPreview(targetEl, iconName, iconColor = DEFAULT_ICON_COLOR) {
  if (!targetEl) return;
  const name = sanitizeIcon(iconName);
  const color = sanitizeIconColor(iconColor);
  targetEl.dataset.icon = name;
  targetEl.dataset.icoColor = color;
  targetEl.innerHTML = iconToSvgLocal(name, 18);
}

function updateCategorySuggestions(items = []) {
  if (!categorySuggestionsEl) return;
  const categories = Array.from(
    new Set(
      (items || [])
        .map((item) => sanitizeCategory(item.category))
        .filter((name) => name && name !== DEFAULT_CATEGORY)
    )
  ).sort((a, b) => a.localeCompare(b, 'ja'));
  categorySuggestionsEl.innerHTML = '';
  categories.forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    categorySuggestionsEl.appendChild(option);
  });
}

function normalizeOrigin(input) {
  if (!input) return '';
  const trimmed = String(input).trim();
  if (!trimmed) return '';
  const withoutWildcard = trimmed.endsWith('*') ? trimmed.slice(0, -1) : trimmed;
  try {
    const url = new URL(withoutWildcard);
    return url.origin;
  } catch (_e) {
    return '';
  }
}

function originPatternFor(input) {
  const origin = normalizeOrigin(input);
  if (!origin) return null;
  return `${origin}/*`;
}

function rememberGrantedOrigin(input) {
  const origin = normalizeOrigin(input);
  if (origin) ensuredOrigins.add(origin);
}

async function primeExistingPermissionCache() {
  if (!chrome?.permissions?.getAll) return;
  try {
    const perms = await chrome.permissions.getAll();
    (perms?.origins || []).forEach((pattern) => rememberGrantedOrigin(pattern));
  } catch (error) {
    console.debug('permissions cache prime failed', error);
  }
}

primeExistingPermissionCache();

function populateIconSelect(selectEl) {
  if (!selectEl) return;
  selectEl.innerHTML = '';
  ICON_CHOICES.forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    selectEl.appendChild(option);
  });
  selectEl.value = DEFAULT_ICON;
}

function populateIconColorSelect(selectEl) {
  if (!selectEl) return;
  selectEl.innerHTML = '';
  ICON_COLOR_CHOICES.forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = ICON_COLOR_LABELS[name] || name;
    selectEl.appendChild(option);
  });
  selectEl.value = DEFAULT_ICON_COLOR;
}

function populateShortcutIconSelect(selectEl, currentValue = '') {
  if (!selectEl) return;
  const normalized = sanitizeShortcutIcon(currentValue);
  selectEl.innerHTML = '';
  ICON_CHOICES.forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    selectEl.appendChild(option);
  });
  selectEl.value = normalized;
}

function populateShortcutIconColorSelect(selectEl, currentValue = '') {
  if (!selectEl) return;
  const normalized = sanitizeShortcutIconColor(currentValue);
  selectEl.innerHTML = '';
  ICON_COLOR_CHOICES.forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = ICON_COLOR_LABELS[name] || name;
    selectEl.appendChild(option);
  });
  selectEl.value = normalized;
}

function shortcutEditorPreviewState(iconValue, colorValue) {
  const icon = sanitizeShortcutIcon(iconValue);
  const color = sanitizeShortcutIconColor(colorValue);
  return { icon, color };
}

populateIconSelect(iconEl);
populateIconSelect(editIconEl);
populateIconColorSelect(iconColorEl);
populateIconColorSelect(editIconColorEl);
renderIconPreview(iconPreviewEl, DEFAULT_ICON, DEFAULT_ICON_COLOR);
renderIconPreview(editIconPreviewEl, DEFAULT_ICON, DEFAULT_ICON_COLOR);
iconEl?.addEventListener('change', () => renderIconPreview(iconPreviewEl, iconEl.value, iconColorEl?.value));
iconColorEl?.addEventListener('change', () => renderIconPreview(iconPreviewEl, iconEl?.value, iconColorEl.value));
editIconEl?.addEventListener('change', () => renderIconPreview(editIconPreviewEl, editIconEl.value, editIconColorEl?.value));
editIconColorEl?.addEventListener('change', () => renderIconPreview(editIconPreviewEl, editIconEl?.value, editIconColorEl.value));

async function warmHostPermissionStatus(items = []) {
  const hosts = Array.from(new Set(items.map((item) => item.host).filter(Boolean)));
  for (const host of hosts) {
    const origin = normalizeOrigin(host);
    if (!origin || ensuredOrigins.has(origin)) continue;
    await hasHostPermission(host);
  }
}

async function ensureHostPermissionFor(host) {
  const origin = normalizeOrigin(host);
  if (!origin) return true;
  if (ensuredOrigins.has(origin)) return true;
  return await requestHostPermission(origin);
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
  const icon = sanitizeShortcutIcon(entry.icon);
  const iconColor = sanitizeShortcutIconColor(entry.iconColor);
  const order = Number.isFinite(entry.order) ? entry.order : fallbackOrder;
  return {
    id: entry.id || createId(),
    type,
    host,
    appId,
    viewIdOrName,
    label,
    initial,
    icon,
    iconColor,
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
    `App: ${entry.appId || '-'}`,
    `Icon: ${sanitizeShortcutIcon(entry.icon)}`,
    `色: ${sanitizeShortcutIconColor(entry.iconColor)}`
  ];
  if (entry.type === 'view') {
    parts.push(`View: ${entry.viewIdOrName || '-'}`);
  }
  return parts.join(' / ');
}

async function persistShortcutEntries() {
  shortcutEntries = shortcutEntries.map((entry, index) => ({
    ...entry,
    order: index,
    initial: normalizeShortcutInitialLocal(entry.initial),
    icon: sanitizeShortcutIcon(entry.icon),
    iconColor: sanitizeShortcutIconColor(entry.iconColor)
  }));
  const core = await import('./core.js');
  await core.saveShortcuts(shortcutEntries);
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
    entry.icon = sanitizeShortcutIcon(entry.icon);
    entry.iconColor = sanitizeShortcutIconColor(entry.iconColor);
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
    const titleIcon = document.createElement('span');
    titleIcon.className = 'item-icon lc';
    const rowPreview = shortcutEditorPreviewState(entry.icon, entry.iconColor);
    titleIcon.dataset.icon = rowPreview.icon;
    titleIcon.dataset.icoColor = rowPreview.color;
    titleIcon.setAttribute('aria-hidden', 'true');
    const titleText = document.createElement('span');
    titleText.textContent = entry.label || shortcutTypeLabel(entry.type);
    title.appendChild(titleIcon);
    title.appendChild(titleText);
    const meta = document.createElement('div');
    meta.className = 'shortcut-meta';
    meta.textContent = shortcutMetaText(entry);
    info.appendChild(title);
    info.appendChild(meta);

    const editor = document.createElement('div');
    editor.className = 'shortcut-editor';

    const editorGrid = document.createElement('div');
    editorGrid.className = 'grid icon-cat-grid shortcut-editor-grid';

    const labelWrap = document.createElement('label');
    labelWrap.textContent = '表示名';
    labelWrap.appendChild(document.createElement('br'));
    const labelInput = document.createElement('input');
    labelInput.value = entry.label || '';
    labelWrap.appendChild(labelInput);

    const iconWrap = document.createElement('label');
    iconWrap.textContent = 'アイコン';
    iconWrap.appendChild(document.createElement('br'));
    const iconSelect = document.createElement('select');
    populateShortcutIconSelect(iconSelect, entry.icon);
    const iconPreview = document.createElement('span');
    iconPreview.className = 'icon-preview lc';
    iconPreview.setAttribute('aria-hidden', 'true');
    iconWrap.appendChild(iconSelect);
    iconWrap.appendChild(iconPreview);

    const colorWrap = document.createElement('label');
    colorWrap.textContent = 'アイコン色';
    colorWrap.appendChild(document.createElement('br'));
    const colorSelect = document.createElement('select');
    populateShortcutIconColorSelect(colorSelect, entry.iconColor);
    colorWrap.appendChild(colorSelect);

    editorGrid.appendChild(labelWrap);
    editorGrid.appendChild(iconWrap);
    editorGrid.appendChild(colorWrap);

    const editorActions = document.createElement('div');
    editorActions.className = 'shortcut-editor-actions';
    const saveBtn = document.createElement('button');
    saveBtn.type = 'button';
    saveBtn.textContent = '保存';
    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.textContent = 'キャンセル';

    editorActions.appendChild(saveBtn);
    editorActions.appendChild(cancelBtn);
    editor.appendChild(editorGrid);
    editor.appendChild(editorActions);
    info.appendChild(editor);

    const updateEditorPreview = () => {
      const preview = shortcutEditorPreviewState(iconSelect.value, colorSelect.value);
      iconPreview.dataset.icon = preview.icon;
      iconPreview.dataset.icoColor = preview.color;
      renderLucideIconsLocal(editor);
    };
    updateEditorPreview();

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
    editBtn.addEventListener('click', () => {
      const opening = !editor.classList.contains('is-open');
      editor.classList.toggle('is-open', opening);
      editBtn.textContent = opening ? '閉じる' : '編集';
      if (opening) {
        labelInput.focus();
        labelInput.select();
      }
    });
    const delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.textContent = '削除';
    delBtn.addEventListener('click', async () => {
      shortcutEntries = shortcutEntries.filter((item) => item.id !== entry.id);
      await persistShortcutEntries();
      renderShortcutEntries();
    });
    iconSelect.addEventListener('change', updateEditorPreview);
    colorSelect.addEventListener('change', updateEditorPreview);
    saveBtn.addEventListener('click', async () => {
      entry.label = labelInput.value.trim();
      entry.icon = sanitizeShortcutIcon(iconSelect.value);
      entry.iconColor = sanitizeShortcutIconColor(colorSelect.value);
      await persistShortcutEntries();
      renderShortcutEntries();
    });
    cancelBtn.addEventListener('click', () => {
      labelInput.value = entry.label || '';
      populateShortcutIconSelect(iconSelect, entry.icon);
      populateShortcutIconColorSelect(colorSelect, entry.iconColor);
      updateEditorPreview();
      editor.classList.remove('is-open');
      editBtn.textContent = '編集';
    });
    ops.appendChild(openA);
    ops.appendChild(editBtn);
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
  renderLucideIconsLocal(shortcutListEl);
}

shortcutToggleEl?.addEventListener('change', async () => {
  shortcutsVisible = Boolean(shortcutToggleEl.checked);
  await chrome.storage.sync.set({ kfavShortcutsVisible: shortcutsVisible });
});

function createId() {
  return crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
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
    li.className = 'pin-item pin-card-item';
    li.dataset.id = entry.id;

    const main = document.createElement('div');
    main.className = 'pin-main';
    const top = document.createElement('div');
    top.className = 'pin-top';

    const title = document.createElement('div');
    title.className = 'pin-title';
    const pinMark = document.createElement('span');
    pinMark.className = 'pin-mark';
    pinMark.textContent = '📌';
    pinMark.setAttribute('aria-hidden', 'true');
    const titleText = document.createElement('span');
    titleText.className = 'pin-title-text';
    titleText.textContent = pinnedEntryLabel(entry);
    titleText.title = titleText.textContent;
    title.appendChild(pinMark);
    title.appendChild(titleText);

    const ops = document.createElement('div');
    ops.className = 'pin-ops pin-actions';
    const openBtn = document.createElement('a');
    openBtn.className = 'watch-icon-btn pin-icon-btn pin-open-btn';
    openBtn.textContent = '↗';
    openBtn.title = '開く';
    openBtn.setAttribute('aria-label', '開く');
    if (entry.host && entry.appId && entry.recordId) {
      openBtn.href = `${entry.host}/k/${entry.appId}/show#record=${entry.recordId}`;
      openBtn.target = '_blank';
      openBtn.rel = 'noopener';
    } else {
      openBtn.href = '#';
      openBtn.setAttribute('aria-disabled', 'true');
      openBtn.classList.add('disabled', 'is-disabled');
      openBtn.addEventListener('click', (ev) => ev.preventDefault());
    }
    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.className = 'watch-icon-btn pin-icon-btn pin-edit-btn';
    editBtn.textContent = '✎';
    editBtn.title = '編集';
    editBtn.setAttribute('aria-label', '編集');
    editBtn.addEventListener('click', () => openPinEditModal(entry, editBtn));
    const delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.className = 'watch-icon-btn pin-icon-btn pin-del-btn';
    delBtn.textContent = '×';
    delBtn.title = '削除';
    delBtn.setAttribute('aria-label', '削除');
    delBtn.addEventListener('click', async () => {
      pinnedEntries = pinnedEntries.filter((item) => item.id !== entry.id);
      await savePinnedEntries();
      renderPinnedEntries();
    });
    ops.appendChild(openBtn);
    ops.appendChild(editBtn);
    ops.appendChild(delBtn);

    top.appendChild(title);
    top.appendChild(ops);

    const summary = document.createElement('div');
    summary.className = 'watch-summary pin-summary';
    const appChip = document.createElement('span');
    appChip.className = 'watch-chip pin-chip';
    appChip.textContent = `App ${entry.appId || '-'}`;
    const recordChip = document.createElement('span');
    recordChip.className = 'watch-chip pin-chip';
    recordChip.textContent = `Record ${entry.recordId || '-'}`;
    summary.appendChild(appChip);
    summary.appendChild(recordChip);

    const noteText = (entry.note || '').trim();
    let noteEl = null;
    if (noteText) {
      noteEl = document.createElement('div');
      noteEl.className = 'watch-query pin-note-line';
      noteEl.textContent = noteText;
      noteEl.title = noteText;
    }

    const details = document.createElement('details');
    details.className = 'watch-details pin-details';
    const detailSummary = document.createElement('summary');
    detailSummary.textContent = '詳細';
    const detailBody = document.createElement('div');
    detailBody.className = 'watch-details-body pin-details-body';
    const fullUrl = entry.host && entry.appId && entry.recordId
      ? `${entry.host}/k/${entry.appId}/show#record=${entry.recordId}`
      : '-';
    [
      ['Host', entry.host || '-'],
      ['URL', fullUrl],
      ['App ID', entry.appId || '-'],
      ['Record ID', entry.recordId || '-'],
      ['Title field', entry.titleField || '-']
    ].forEach(([label, value]) => {
      const row = document.createElement('div');
      row.className = 'watch-detail-row pin-detail-row';
      row.textContent = `${label}: ${value}`;
      row.title = String(value);
      detailBody.appendChild(row);
    });
    details.appendChild(detailSummary);
    details.appendChild(detailBody);

    main.appendChild(top);
    main.appendChild(summary);
    if (noteEl) main.appendChild(noteEl);
    main.appendChild(details);
    li.appendChild(main);
    pinListEl.appendChild(li);
  });
}

function resetPinnedForm() {
  if (pinLabelEl) pinLabelEl.value = '';
  pinUrlEl.value = '';
  pinAppIdEl.value = '';
  pinRecordIdEl.value = '';
  pinTitleFieldEl.value = '';
  pinNoteEl.value = '';
  if (pinSaveBtn) pinSaveBtn.textContent = '追加';
}
async function loadFavorites() {
  const { kintoneFavorites = [] } = await chrome.storage.sync.get('kintoneFavorites');
  return kintoneFavorites.map((item, idx) => normalizeFavoriteEntryLocal(item, idx));
}
async function saveFavorites(items) {
  const normalized = items.map((item, idx) => normalizeFavoriteEntryLocal(item, idx));
  await chrome.storage.sync.set({ kintoneFavorites: normalized });
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
  updateCategorySuggestions(sorted);

  listEl.innerHTML = '';
  if (!sorted.length) {
    const empty = document.createElement('li');
    empty.className = 'item item-empty';
    empty.textContent = '登録済みのウォッチリストはまだありません。';
    listEl.appendChild(empty);
    return;
  }

  sorted.forEach((it, idx) => {
    if (typeof it.order !== 'number') it.order = idx;

    const li = document.createElement('li');
    li.className = 'item watch-item';
    li.draggable = true;
    li.dataset.id = it.id;

    const handle = document.createElement('div');
    handle.className = 'drag-handle watch-drag-handle';
    handle.textContent = '⋮⋮';
    handle.title = 'ドラッグで並び替え';

    const body = document.createElement('div');
    body.className = 'watch-main';
    const top = document.createElement('div');
    top.className = 'watch-top';
    const title = document.createElement('div');
    title.className = 'title watch-title';
    const iconPreview = document.createElement('span');
    iconPreview.className = 'item-icon lc';
    iconPreview.dataset.icon = sanitizeIcon(it.icon);
    iconPreview.dataset.icoColor = sanitizeIconColor(it.iconColor);
    iconPreview.setAttribute('aria-hidden', 'true');
    const titleText = document.createElement('span');
    titleText.className = 'watch-title-text';
    titleText.textContent = it.label || '(no label)';
    title.appendChild(iconPreview);
    title.appendChild(titleText);
    top.appendChild(title);

    const right = document.createElement('div');
    right.className = 'ops watch-actions';

    // バッジ対象ラジオ
    const badgeWrap = document.createElement('label');
    badgeWrap.className = 'watch-icon-btn watch-badge-btn';
    badgeWrap.title = 'このウォッチリストをバッジ対象にする';
    badgeWrap.setAttribute('aria-label', 'このウォッチリストをバッジ対象にする');
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = 'badgeTarget';
    radio.className = 'badge-radio watch-badge-radio';
    radio.checked = badgeTargetId ? (badgeTargetId === it.id) : (sorted[0]?.id === it.id);
    const badgeMark = document.createElement('span');
    badgeMark.className = 'watch-badge-mark';
    badgeMark.textContent = '◎';
    radio.addEventListener('change', async () => {
      await setBadgeTarget(it.id);
      const wraps = listEl.querySelectorAll('.watch-badge-btn');
      wraps.forEach((el) => el.classList.remove('is-active'));
      badgeWrap.classList.add('is-active');
    });
    badgeWrap.classList.toggle('is-active', radio.checked);
    badgeWrap.appendChild(radio);
    badgeWrap.appendChild(badgeMark);

    // ピン切替
    const pinBtn = document.createElement('button');
    pinBtn.className = 'pin watch-icon-btn watch-pin-btn';
    pinBtn.type = 'button';
    pinBtn.textContent = '📌';
    pinBtn.title = it.pinned ? '固定を解除' : '固定する';
    pinBtn.setAttribute('aria-label', it.pinned ? '固定を解除' : '固定する');
    pinBtn.classList.toggle('is-active', Boolean(it.pinned));
    pinBtn.addEventListener('click', async () => {
      const all = await loadFavorites();
      const me = all.find(x => x.id === it.id);
      me.pinned = !me.pinned;
      await saveFavorites(all);
      render(all);
    });

    // 開く・削除
    const openA = document.createElement('a');
    openA.textContent = '↗';
    openA.className = 'btn watch-icon-btn watch-open-btn';
    openA.title = '開く';
    openA.setAttribute('aria-label', '開く');
    openA.href = it.url;
    openA.target = '_blank';
    openA.rel = 'noopener';
    // 編集ボタン
    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.textContent = '✎';
    editBtn.className = 'watch-icon-btn watch-edit-btn';
    editBtn.title = '編集';
    editBtn.setAttribute('aria-label', '編集');
    editBtn.addEventListener('click', () => openEdit(it, editBtn));

    const delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.textContent = '✕';
    delBtn.className = 'watch-icon-btn watch-del-btn';
    delBtn.title = '削除';
    delBtn.setAttribute('aria-label', '削除');
    delBtn.addEventListener('click', async () => {
      const next = (await loadFavorites()).filter(x => x.id !== it.id);
      await saveFavorites(next);
      render(next);
    });

    right.appendChild(badgeWrap);
    right.appendChild(pinBtn);
    right.appendChild(openA);
    right.appendChild(editBtn);
    right.appendChild(delBtn);
    top.appendChild(right);

    const summary = document.createElement('div');
    summary.className = 'watch-summary';
    const addChip = (text) => {
      const chip = document.createElement('span');
      chip.className = 'watch-chip';
      chip.textContent = text;
      summary.appendChild(chip);
    };
    if (it.appId) addChip(`App ${it.appId}`);
    if (it.viewIdOrName) addChip(`View ${it.viewIdOrName}`);
    addChip(getCategoryLabel(it.category));

    const urlLine = document.createElement('div');
    urlLine.className = 'watch-url';
    urlLine.textContent = it.url || '-';
    urlLine.title = it.url || '';

    const queryLine = document.createElement('div');
    queryLine.className = 'watch-query';
    if (it.query) {
      queryLine.textContent = `Query: ${it.query}`;
      queryLine.title = it.query;
    }

    const details = document.createElement('details');
    details.className = 'watch-details';
    const detailsSummary = document.createElement('summary');
    detailsSummary.textContent = '詳細';
    const detailsBody = document.createElement('div');
    detailsBody.className = 'watch-details-body';
    const metaRows = [
      ['Host', it.host || '-'],
      ['URL', it.url || '-'],
      ['App', it.appId || '-'],
      ['View', it.viewIdOrName || '-'],
      ['Query', it.query || '-'],
      ['Icon', sanitizeIcon(it.icon)],
      ['色', sanitizeIconColor(it.iconColor)],
      ['カテゴリ', getCategoryLabel(it.category)]
    ];
    metaRows.forEach(([label, value]) => {
      const row = document.createElement('div');
      row.className = 'watch-detail-row';
      row.textContent = `${label}: ${value}`;
      detailsBody.appendChild(row);
    });
    details.appendChild(detailsSummary);
    details.appendChild(detailsBody);

    body.appendChild(top);
    body.appendChild(summary);
    body.appendChild(urlLine);
    if (it.query) body.appendChild(queryLine);
    body.appendChild(details);

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
    li.appendChild(body);
    listEl.appendChild(li);
  });
  renderLucideIconsLocal(listEl);
}

// ---- 編集処理（renderの外に1回だけ定義）----
const editModalState = {
  triggerEl: null,
  previousBodyOverflow: '',
  handlers: null
};

function getDialogFocusableElements() {
  if (!editDialog) return [];
  const selector = [
    'button:not([disabled])',
    'a[href]',
    'input:not([disabled])',
    'select:not([disabled])',
    'textarea:not([disabled])',
    '[tabindex]:not([tabindex="-1"])'
  ].join(',');
  return Array.from(editDialog.querySelectorAll(selector))
    .filter((el) => !el.hasAttribute('hidden') && el.offsetParent !== null);
}

function handleDialogTabTrap(event) {
  if (event.key !== 'Tab' || !editDialog?.open) return;
  const focusables = getDialogFocusableElements();
  if (!focusables.length) {
    event.preventDefault();
    return;
  }
  const first = focusables[0];
  const last = focusables[focusables.length - 1];
  const active = document.activeElement;
  if (event.shiftKey && active === first) {
    event.preventDefault();
    last.focus();
  } else if (!event.shiftKey && active === last) {
    event.preventDefault();
    first.focus();
  }
}

function closeEditDialog({ restoreFocus = true } = {}) {
  if (!editDialog) return;
  const closingId = editingId;
  const handlers = editModalState.handlers;
  if (handlers) {
    document.removeEventListener('keydown', handlers.onDocumentKeydown, true);
    editDialog.removeEventListener('click', handlers.onDialogClick);
    editDialog.removeEventListener('cancel', handlers.onDialogCancel);
    editDialog.removeEventListener('keydown', handlers.onDialogKeydown);
    editModalState.handlers = null;
  }
  if (editDialog.open && typeof editDialog.close === 'function') {
    editDialog.close();
  }
  document.body.style.overflow = editModalState.previousBodyOverflow;
  let focusTarget = editModalState.triggerEl;
  if ((!focusTarget || !focusTarget.isConnected) && closingId) {
    const escaped = globalThis.CSS?.escape ? CSS.escape(String(closingId)) : String(closingId);
    focusTarget = document.querySelector(`#list .watch-item[data-id="${escaped}"] .watch-edit-btn`);
  }
  if (restoreFocus && focusTarget && typeof focusTarget.focus === 'function') {
    try {
      focusTarget.focus();
    } catch (_err) {
      // ignore
    }
  }
  editModalState.triggerEl = null;
  editingId = null;
}

function openEdit(item, triggerEl = null) {
  editingId = item.id;
  editLabelEl.value = item.label || '';
  editUrlEl.value = item.url || '';
  editAppIdEl.value = item.appId || '';
  editViewEl.value = item.viewIdOrName || '';
  editQueryEl.value = item.query || '';
  if (editIconEl) {
    editIconEl.value = sanitizeIcon(item.icon);
  }
  if (editIconColorEl) {
    editIconColorEl.value = sanitizeIconColor(item.iconColor);
  }
  if (editIconEl) {
    renderIconPreview(editIconPreviewEl, editIconEl.value, editIconColorEl?.value);
  }
  if (editCategoryEl) {
    editCategoryEl.value = sanitizeCategory(item.category);
  }
  if (typeof editDialog.showModal === 'function') {
    if (editDialog.open) {
      closeEditDialog({ restoreFocus: false });
    }
    editModalState.triggerEl = triggerEl || document.activeElement;
    editModalState.previousBodyOverflow = document.body.style.overflow;
    const onDocumentKeydown = (event) => {
      if (event.isComposing) return;
      if (event.key === 'Escape') {
        event.preventDefault();
        closeEditDialog();
      }
    };
    const onDialogClick = (event) => {
      if (event.target === editDialog) {
        closeEditDialog();
      }
    };
    const onDialogCancel = (event) => {
      event.preventDefault();
      closeEditDialog();
    };
    const onDialogKeydown = (event) => {
      handleDialogTabTrap(event);
    };
    editModalState.handlers = {
      onDocumentKeydown,
      onDialogClick,
      onDialogCancel,
      onDialogKeydown
    };
    document.addEventListener('keydown', onDocumentKeydown, true);
    editDialog.addEventListener('click', onDialogClick);
    editDialog.addEventListener('cancel', onDialogCancel);
    editDialog.addEventListener('keydown', onDialogKeydown);
    editDialog.showModal();
    document.body.style.overflow = 'hidden';
    setTimeout(() => {
      try {
        editLabelEl.focus();
        editLabelEl.select();
      } catch (_err) {
        // ignore
      }
    }, 0);
  } else {
    alert('このブラウザはdialog要素に対応していません');
  }
}

editCancelBtn?.addEventListener('click', () => closeEditDialog());
editCloseXBtn?.addEventListener('click', () => closeEditDialog());

editSaveBtn.addEventListener('click', async (e) => {
  e.preventDefault();
  if (!editingId) { closeEditDialog(); return; }
  const newLabel = editLabelEl.value.trim();
  const newUrl = editUrlEl.value.trim();
  let { host, appId, viewIdOrName } = parseKintoneUrl(newUrl);
  if (editAppIdEl.value.trim()) appId = editAppIdEl.value.trim();
  if (editViewEl.value.trim()) viewIdOrName = editViewEl.value.trim();
  const newQuery = editQueryEl.value.trim() || '';
  const newIcon = sanitizeIcon(editIconEl?.value);
  const newIconColor = sanitizeIconColor(editIconColorEl?.value);
  const newCategory = sanitizeCategory(editCategoryEl?.value);
  if (!newUrl || !host) { alert('有効な kintone URL を入力してください'); return; }
  if (!(await ensureHostPermissionFor(host))) {
    alert('このドメインへのアクセス権限が許可されませんでした');
    return;
  }
  const all = await loadFavorites();
  const idx = all.findIndex(x => x.id === editingId);
  if (idx === -1) { closeEditDialog(); return; }
  if (all.some((x, i) => i !== idx && x.url === newUrl)) {
    alert('このURLは既に登録済みです'); return;
  }
  const old = all[idx];
  all[idx] = {
    ...old,
    label: newLabel,
    url: newUrl,
    host,
    appId: appId||'',
    viewIdOrName: viewIdOrName||'',
    query: newQuery,
    icon: newIcon,
    iconColor: newIconColor,
    category: newCategory
  };
  await saveFavorites(all);
  await render(all);
  closeEditDialog();
});

const pinEditModalState = {
  triggerEl: null,
  previousBodyOverflow: '',
  handlers: null
};

function getPinDialogFocusableElements() {
  if (!pinEditDialog) return [];
  const selector = [
    'button:not([disabled])',
    'a[href]',
    'input:not([disabled])',
    'select:not([disabled])',
    'textarea:not([disabled])',
    '[tabindex]:not([tabindex="-1"])'
  ].join(',');
  return Array.from(pinEditDialog.querySelectorAll(selector))
    .filter((el) => !el.hasAttribute('hidden') && el.offsetParent !== null);
}

function handlePinDialogTabTrap(event) {
  if (event.key !== 'Tab' || !pinEditDialog?.open) return;
  const focusables = getPinDialogFocusableElements();
  if (!focusables.length) {
    event.preventDefault();
    return;
  }
  const first = focusables[0];
  const last = focusables[focusables.length - 1];
  const active = document.activeElement;
  if (event.shiftKey && active === first) {
    event.preventDefault();
    last.focus();
  } else if (!event.shiftKey && active === last) {
    event.preventDefault();
    first.focus();
  }
}

function closePinEditDialog({ restoreFocus = true } = {}) {
  if (!pinEditDialog) return;
  const closingId = pinModalEditingId;
  const handlers = pinEditModalState.handlers;
  if (handlers) {
    document.removeEventListener('keydown', handlers.onDocumentKeydown, true);
    pinEditDialog.removeEventListener('click', handlers.onDialogClick);
    pinEditDialog.removeEventListener('cancel', handlers.onDialogCancel);
    pinEditDialog.removeEventListener('keydown', handlers.onDialogKeydown);
    pinEditModalState.handlers = null;
  }
  if (pinEditDialog.open && typeof pinEditDialog.close === 'function') {
    pinEditDialog.close();
  }
  document.body.style.overflow = pinEditModalState.previousBodyOverflow;
  let focusTarget = pinEditModalState.triggerEl;
  if ((!focusTarget || !focusTarget.isConnected) && closingId) {
    const escaped = globalThis.CSS?.escape ? CSS.escape(String(closingId)) : String(closingId);
    focusTarget = document.querySelector(`#pin_list .pin-item[data-id="${escaped}"] .pin-edit-btn`);
  }
  if (restoreFocus && focusTarget && typeof focusTarget.focus === 'function') {
    try {
      focusTarget.focus();
    } catch (_err) {
      // ignore
    }
  }
  pinEditModalState.triggerEl = null;
  pinModalEditingId = null;
}

function openPinEditModal(entry, triggerEl = null) {
  if (!entry || !pinEditDialog) return;
  pinModalEditingId = entry.id;
  if (pinEditLabelEl) pinEditLabelEl.value = entry.label || '';
  if (pinEditUrlEl) {
    pinEditUrlEl.value = entry.host && entry.appId && entry.recordId
      ? `${entry.host}/k/${entry.appId}/show#record=${entry.recordId}`
      : '';
  }
  if (pinEditAppIdEl) pinEditAppIdEl.value = entry.appId || '';
  if (pinEditRecordIdEl) pinEditRecordIdEl.value = entry.recordId || '';
  if (pinEditTitleFieldEl) pinEditTitleFieldEl.value = entry.titleField || '';
  if (pinEditNoteEl) pinEditNoteEl.value = entry.note || '';

  if (typeof pinEditDialog.showModal === 'function') {
    if (pinEditDialog.open) {
      closePinEditDialog({ restoreFocus: false });
    }
    pinEditModalState.triggerEl = triggerEl || document.activeElement;
    pinEditModalState.previousBodyOverflow = document.body.style.overflow;
    const onDocumentKeydown = (event) => {
      if (event.isComposing) return;
      if (event.key === 'Escape') {
        event.preventDefault();
        closePinEditDialog();
      }
    };
    const onDialogClick = (event) => {
      if (event.target === pinEditDialog) {
        closePinEditDialog();
      }
    };
    const onDialogCancel = (event) => {
      event.preventDefault();
      closePinEditDialog();
    };
    const onDialogKeydown = (event) => {
      handlePinDialogTabTrap(event);
    };
    pinEditModalState.handlers = {
      onDocumentKeydown,
      onDialogClick,
      onDialogCancel,
      onDialogKeydown
    };
    document.addEventListener('keydown', onDocumentKeydown, true);
    pinEditDialog.addEventListener('click', onDialogClick);
    pinEditDialog.addEventListener('cancel', onDialogCancel);
    pinEditDialog.addEventListener('keydown', onDialogKeydown);
    pinEditDialog.showModal();
    document.body.style.overflow = 'hidden';
    setTimeout(() => {
      try {
        (pinEditLabelEl || pinEditUrlEl)?.focus();
      } catch (_err) {
        // ignore
      }
    }, 0);
  } else {
    alert('このブラウザはdialog要素に対応していません');
  }
}

pinEditCancelBtn?.addEventListener('click', () => closePinEditDialog());
pinEditCloseXBtn?.addEventListener('click', () => closePinEditDialog());

pinEditSaveBtn?.addEventListener('click', async (event) => {
  event.preventDefault();
  if (!pinModalEditingId) {
    closePinEditDialog();
    return;
  }
  const label = pinEditLabelEl?.value.trim() || '';
  const url = pinEditUrlEl?.value.trim() || '';
  const parsed = parseKintoneUrl(url);
  const host = parsed.host || hostFromUrl(url);
  const appId = (pinEditAppIdEl?.value || parsed.appId || '').toString().trim();
  const recordId = (pinEditRecordIdEl?.value || parsed.recordId || '').toString().trim();
  const titleField = pinEditTitleFieldEl?.value.trim() || '';
  const note = pinEditNoteEl?.value.trim() || '';

  if (!host || !appId || !recordId) {
    alert('host / App ID / レコードID は必須です');
    return;
  }

  const idx = pinnedEntries.findIndex((item) => item.id === pinModalEditingId);
  if (idx === -1) {
    closePinEditDialog();
    return;
  }
  pinnedEntries[idx] = normalizePinnedEntry({
    id: pinModalEditingId,
    label,
    host,
    appId,
    recordId,
    titleField,
    note
  });
  await savePinnedEntries();
  renderPinnedEntries();
  closePinEditDialog();
});

addBtn.addEventListener('click', async () => {
  const label = labelEl.value.trim();
  const url = urlEl.value.trim();
  let { host, appId, viewIdOrName } = parseKintoneUrl(url);

  if (appIdEl.value.trim()) appId = appIdEl.value.trim();
  if (viewEl.value.trim()) viewIdOrName = viewEl.value.trim();
  const query = queryEl.value.trim() || '';
  const icon = sanitizeIcon(iconEl?.value);
  const iconColor = sanitizeIconColor(iconColorEl?.value);
  const category = sanitizeCategory(categoryEl?.value);

  if (!url || !host) {
    alert('有効な kintone URL を入力してください');
    return;
  }
  if (!(await ensureHostPermissionFor(host))) {
    alert('このドメインへのアクセス権限が許可されませんでした');
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
    icon,
    iconColor,
    category,
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
  if (iconEl) {
    iconEl.value = DEFAULT_ICON;
  }
  if (iconColorEl) {
    iconColorEl.value = DEFAULT_ICON_COLOR;
  }
  if (iconEl) {
    renderIconPreview(iconPreviewEl, iconEl.value, iconColorEl?.value);
  }
  if (categoryEl) categoryEl.value = '';
});

// 初期表示
(async () => {
  const items = await loadFavorites();
  await warmHostPermissionStatus(items);
  await render(items);

  const [core, shortcutStored] = await Promise.all([
    import('./core.js'),
    chrome.storage.sync.get(['kfavShortcutsVisible'])
  ]);
  shortcutEntries = core
    .sortShortcuts(await core.loadShortcuts())
    .map((item, idx) => normalizeShortcutEntryLocal(item, idx))
    .filter(Boolean);
  const visibleFlag = shortcutStored.kfavShortcutsVisible;
  shortcutsVisible = typeof visibleFlag === 'boolean' ? visibleFlag : true;
  if (shortcutToggleEl) shortcutToggleEl.checked = shortcutsVisible;
  renderShortcutEntries();

  const pinsStored = await chrome.storage.sync.get(['kfavPins', PIN_VISIBLE_KEY]);
  const pinRaw = pinsStored.kfavPins;
  const pinArr = Array.isArray(pinRaw) ? pinRaw : pinRaw ? [pinRaw] : [];
  pinnedEntries = pinArr.map(normalizePinnedEntry).filter(Boolean);
  renderPinnedEntries();
  resetPinnedForm();
  const pinVisibleFlag = pinsStored[PIN_VISIBLE_KEY];
  const pinVisible = typeof pinVisibleFlag === 'boolean' ? pinVisibleFlag : true;
  if (pinVisibleToggleEl) pinVisibleToggleEl.checked = pinVisible;

  await Promise.all([
    loadExcelColumnPrefs(),
    loadExcelOverlayMode()
  ]);
})();

if (chrome?.storage?.onChanged) {
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area === 'sync' && Object.prototype.hasOwnProperty.call(changes, 'kfavPins')) {
      const next = changes.kfavPins.newValue;
      const arr = Array.isArray(next) ? next : next ? [next] : [];
      pinnedEntries = arr.map(normalizePinnedEntry).filter(Boolean);
      renderPinnedEntries();
    }
    if (area === 'sync' && Object.prototype.hasOwnProperty.call(changes, PIN_VISIBLE_KEY)) {
      const flag = typeof changes[PIN_VISIBLE_KEY].newValue === 'boolean'
        ? changes[PIN_VISIBLE_KEY].newValue
        : true;
      if (pinVisibleToggleEl) pinVisibleToggleEl.checked = flag;
    }
    if (area === 'sync' && Object.prototype.hasOwnProperty.call(changes, EXCEL_COLUMN_PREF_KEY)) {
      excelColumnPrefs = normalizeExcelColumnPrefs(changes[EXCEL_COLUMN_PREF_KEY].newValue);
      renderExcelColumnPrefs();
    }
    if (area === 'sync' && Object.prototype.hasOwnProperty.call(changes, EXCEL_OVERLAY_MODE_KEY)) {
      const mode = normalizeExcelOverlayMode(changes[EXCEL_OVERLAY_MODE_KEY].newValue);
      excelModeInputs.forEach((input) => {
        input.checked = input.value === mode;
      });
    }
  });
}

function hostFromUrl(u){
  if (!u) return '';
  try {
    return new URL(u).origin;
  } catch {
    try {
      return new URL(`https://${u}`).origin;
    } catch {
      return '';
    }
  }
}

function normalizeExcelOverlayMode(value) {
  return EXCEL_OVERLAY_MODE_VALUES.includes(value) ? value : DEFAULT_EXCEL_OVERLAY_MODE;
}

async function loadExcelOverlayMode() {
  if (!excelModeInputs.length) return;
  const stored = await chrome.storage.sync.get(EXCEL_OVERLAY_MODE_KEY);
  const mode = normalizeExcelOverlayMode(stored[EXCEL_OVERLAY_MODE_KEY]);
  excelModeInputs.forEach((input) => {
    input.checked = input.value === mode;
  });
}

async function saveExcelOverlayMode(modeValue) {
  const mode = normalizeExcelOverlayMode(modeValue);
  await chrome.storage.sync.set({ [EXCEL_OVERLAY_MODE_KEY]: mode });
}

excelModeInputs.forEach((input) => {
  input.addEventListener('change', async () => {
    if (!input.checked) return;
    await saveExcelOverlayMode(input.value);
  });
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
    li.innerHTML = '保存済みの列レイアウトはありません。<br/>一覧画面で列並びや列幅を変更すると、ここに保存されます。';
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
    count.textContent = `列順: ${entry.count}列`;
    const widthInfo = document.createElement('span');
    widthInfo.textContent = `幅設定: ${entry.widthCount}列`;
    const saved = document.createElement('span');
    saved.textContent = `保存: ${formatPrefDate(entry.savedAt)}`;
    meta.appendChild(count);
    meta.appendChild(widthInfo);
    meta.appendChild(saved);
    info.appendChild(title);
    info.appendChild(meta);

    const actions = document.createElement('div');
    actions.className = 'excel-columns-actions';
    const clearOrderBtn = document.createElement('button');
    clearOrderBtn.type = 'button';
    clearOrderBtn.textContent = '並びを解除';
    clearOrderBtn.addEventListener('click', () => { clearExcelPrefOrder(entry.key); });
    const clearWidthBtn = document.createElement('button');
    clearWidthBtn.type = 'button';
    clearWidthBtn.textContent = '幅を解除';
    clearWidthBtn.addEventListener('click', () => { clearExcelPrefWidths(entry.key); });
    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.textContent = '削除';
    removeBtn.className = 'btn-danger-subtle';
    removeBtn.addEventListener('click', () => { removeExcelPref(entry.key); });
    actions.appendChild(clearOrderBtn);
    actions.appendChild(clearWidthBtn);
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

async function saveExcelColumnPrefsMap(next) {
  if (Object.keys(next).length) {
    await chrome.storage.sync.set({ [EXCEL_COLUMN_PREF_KEY]: next });
  } else {
    await chrome.storage.sync.remove(EXCEL_COLUMN_PREF_KEY);
  }
}

async function clearExcelPrefOrder(key) {
  if (!key || !excelColumnPrefs[key]) return;
  const current = excelColumnPrefs[key];
  const nextEntry = { ...current, order: [], savedAt: Date.now() };
  const shouldDelete = !nextEntry.order.length && !Object.keys(nextEntry.widths || {}).length;
  const next = { ...excelColumnPrefs };
  if (shouldDelete) delete next[key];
  else next[key] = nextEntry;
  excelColumnPrefs = next;
  await saveExcelColumnPrefsMap(next);
  renderExcelColumnPrefs();
}

async function clearExcelPrefWidths(key) {
  if (!key || !excelColumnPrefs[key]) return;
  const current = excelColumnPrefs[key];
  const nextEntry = { ...current, widths: {}, savedAt: Date.now() };
  const shouldDelete = !nextEntry.order.length && !Object.keys(nextEntry.widths || {}).length;
  const next = { ...excelColumnPrefs };
  if (shouldDelete) delete next[key];
  else next[key] = nextEntry;
  excelColumnPrefs = next;
  await saveExcelColumnPrefsMap(next);
  renderExcelColumnPrefs();
}

async function removeExcelPref(key) {
  if (!key || !excelColumnPrefs[key]) return;
  const next = { ...excelColumnPrefs };
  delete next[key];
  excelColumnPrefs = next;
  await saveExcelColumnPrefsMap(next);
  renderExcelColumnPrefs();
}

excelClearBtn?.addEventListener('click', async () => {
  if (!Object.keys(excelColumnPrefs).length) return;
  if (!window.confirm('保存済みの列レイアウトをすべて削除します。よろしいですか？')) return;
  excelColumnPrefs = {};
  await chrome.storage.sync.remove(EXCEL_COLUMN_PREF_KEY);
  renderExcelColumnPrefs();
});

// ---- Host permission helpers ----
async function hasHostPermission(origin) {
  const pattern = originPatternFor(origin);
  if (!pattern || !chrome?.permissions?.contains) return false;
  try {
    const granted = await chrome.permissions.contains({ origins: [pattern] });
    if (granted) rememberGrantedOrigin(origin);
    return granted;
  } catch (error) {
    console.warn('host permission check failed', error);
    return false;
  }
}

async function requestHostPermission(origin) {
  const pattern = originPatternFor(origin);
  if (!pattern || !chrome?.permissions?.request) return false;
  try {
    const granted = await chrome.permissions.request({ origins: [pattern] });
    if (granted) rememberGrantedOrigin(origin);
    return granted;
  } catch (error) {
    console.warn('host permission request failed', error);
    return false;
  }
}

async function updatePermStatus(origin, statusEl = hostPermStatusEl) {
  const el = statusEl;
  if (!el) return false;
  const normalized = normalizeOrigin(origin);
  if (!normalized) {
    el.textContent = '権限: 未確認';
    return false;
  }
  const granted = await hasHostPermission(normalized);
  el.textContent = granted ? '権限: 許可済み' : '権限: 未許可';
  return granted;
}

hostPermRequestBtn?.addEventListener('click', async () => {
  const raw = hostPermInputEl?.value?.trim() || '';
  const host = hostFromUrl(raw);
  if (!host) {
    alert('有効な kintone URL またはホストを入力してください');
    await updatePermStatus('');
    return;
  }
  const ok = await requestHostPermission(host);
  await updatePermStatus(host);
  if (!ok) {
    alert('許可がキャンセルされました');
    return;
  }
  if (hostPermInputEl) hostPermInputEl.value = host;
  alert('許可が完了しました');
});

hostPermCheckBtn?.addEventListener('click', async () => {
  const raw = hostPermInputEl?.value?.trim() || '';
  const host = hostFromUrl(raw);
  if (!host) {
    alert('有効な kintone URL またはホストを入力してください');
    await updatePermStatus('');
    return;
  }
  await updatePermStatus(host);
});


// ---- pinned events ----
pinResetBtn?.addEventListener('click', () => {
  resetPinnedForm();
});
pinVisibleToggleEl?.addEventListener('change', async () => {
  const visible = Boolean(pinVisibleToggleEl.checked);
  await chrome.storage.sync.set({ [PIN_VISIBLE_KEY]: visible });
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
    id: createId(),
    label,
    host,
    appId,
    recordId,
    titleField,
    note
  });
  pinnedEntries.push(entry);

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
pinEditUrlEl?.addEventListener('change', () => {
  const raw = pinEditUrlEl.value.trim();
  const parsed = parseKintoneUrl(raw);
  if (parsed.appId && pinEditAppIdEl && !pinEditAppIdEl.value) pinEditAppIdEl.value = parsed.appId;
  if (parsed.recordId && pinEditRecordIdEl && !pinEditRecordIdEl.value) pinEditRecordIdEl.value = parsed.recordId;
});
