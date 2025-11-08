'use strict';

import {
  loadPinnedRecords,
  savePinnedRecords,
  ensureHostPermission,
  sendRunInKintone,
  createId,
  getActiveTab,
  parseKintoneUrl,
  isKintoneUrl
} from './core.js';

export function initPins(options = {}) {
  const doc = options.document || document;
  const pinListEl = doc.getElementById('pinList');
  const pinNoticeEl = doc.getElementById('pinNotice');
  const refreshBtn = doc.getElementById('refreshPins');
  const addBtn = doc.getElementById('addPinFromTab');

  if (!pinListEl || !refreshBtn) {
    return {
      refreshPins() {},
      quickAdd() { return false; },
      dispose() {}
    };
  }

  const state = {
    entries: [],
    detailMap: new Map(),
    storageListener: null,
    fetchToken: 0,
    isRefreshing: false,
    skipStorageOnce: false
  };

  let noticeTimer = null;
  let persistTimer = null;
  let dragSourceId = null;
  const escapeSelector = (value) => {
    if (value == null) return '';
    if (globalThis.CSS?.escape) return CSS.escape(String(value));
    return String(value).replace(/[^\w-]/g, (ch) => `\\${ch}`);
  };

  async function persistPins() {
    state.skipStorageOnce = true;
    await savePinnedRecords(state.entries.map(({ id, label, host, appId, recordId, titleField, note }) => ({
      id,
      label,
      host,
      appId,
      recordId,
      titleField,
      note
    })));
  }

  function schedulePersist() {
    if (persistTimer) clearTimeout(persistTimer);
    persistTimer = setTimeout(() => {
      persistTimer = null;
      persistPins().catch(() => {});
    }, 400);
  }

  function setNotice(msg, ttlMs = 0) {
    if (noticeTimer) {
      clearTimeout(noticeTimer);
      noticeTimer = null;
    }
    if (!pinNoticeEl) return;
    if (msg) {
      pinNoticeEl.textContent = msg;
      pinNoticeEl.classList.remove('hidden');
      if (ttlMs > 0) {
        noticeTimer = setTimeout(() => {
          setNotice('');
        }, ttlMs);
      }
    } else {
      pinNoticeEl.textContent = '';
      pinNoticeEl.classList.add('hidden');
    }
  }

  function deriveRecordTitle(entry, detail) {
    const record = detail?.record;
    if (record && entry.titleField && typeof record[entry.titleField]?.value === 'string') {
      const val = record[entry.titleField].value.trim();
      if (val) return val;
    }
    if (record?.$id?.value) return `#${record.$id.value}`;
    return '';
  }

  function defaultLabel(entry) {
    const app = entry.appId ? `App ${entry.appId}` : 'App 未指定';
    const record = entry.recordId ? `Record ${entry.recordId}` : 'Record 未指定';
    return `${app} / ${record}`;
  }

  function createCard(entry, detail) {
    const li = doc.createElement('li');
    li.className = 'pin-card';
    li.dataset.id = entry.id;

    li.addEventListener('dragenter', handleDragEnter, true);
    li.addEventListener('dragover', handleDragOver);
    li.addEventListener('dragleave', handleDragLeave, true);
    li.addEventListener('drop', handleDrop);
    li.addEventListener('dragend', handleDragEnd);

    const header = doc.createElement('div');
    header.className = 'pin-card-header';

    const titleWrap = doc.createElement('div');
    titleWrap.className = 'pin-title-wrap';
    const dragHandle = doc.createElement('span');
    dragHandle.className = 'pin-drag-handle';
    dragHandle.title = 'ドラッグで並び替え';
    dragHandle.textContent = '⋮⋮';
    dragHandle.draggable = true;
    dragHandle.dataset.cardId = entry.id;
    dragHandle.setAttribute('role', 'button');
    dragHandle.setAttribute('aria-label', 'ドラッグで並び替え');
    dragHandle.addEventListener('dragstart', handleDragStart);
    dragHandle.addEventListener('dragend', handleDragEnd);
    dragHandle.addEventListener('click', (ev) => ev.preventDefault());

    const titleInput = doc.createElement('input');
    titleInput.type = 'text';
    titleInput.className = 'pin-title-input';
    titleInput.value = entry.label || '';
    titleInput.placeholder = deriveRecordTitle(entry, detail) || defaultLabel(entry);
    titleInput.addEventListener('input', () => {
      entry.label = titleInput.value;
      schedulePersist();
    });
    titleInput.addEventListener('blur', () => {
      entry.label = titleInput.value.trim();
      titleInput.value = entry.label;
      schedulePersist();
    });
    titleWrap.appendChild(dragHandle);
    titleWrap.appendChild(titleInput);
    header.appendChild(titleWrap);

    const actions = doc.createElement('div');
    actions.className = 'pin-actions';
    const openBtn = doc.createElement('a');
    openBtn.textContent = '開く';
    openBtn.href = `${entry.host}/k/${entry.appId}/show#record=${entry.recordId}`;
    openBtn.target = '_blank';
    openBtn.rel = 'noopener';

    const removeBtn = doc.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'pin-remove';
    removeBtn.title = 'ピンを外す';
    removeBtn.textContent = '🗑';
    removeBtn.addEventListener('click', (ev) => {
      ev.stopPropagation();
      removeEntry(entry.id).catch(() => {});
    });

    actions.appendChild(openBtn);
    actions.appendChild(removeBtn);
    header.appendChild(actions);
    li.appendChild(header);

    const recordTitle = deriveRecordTitle(entry, detail);
    if (recordTitle) {
      const recordTitleEl = doc.createElement('div');
      recordTitleEl.className = 'pin-record-title';
      recordTitleEl.textContent = recordTitle;
      li.appendChild(recordTitleEl);
    }

    const meta = doc.createElement('div');
    meta.className = 'pin-meta';
    const metaItems = [
      `Host: ${entry.host || '-'}`,
      `App: ${entry.appId || '-'}`,
      `Record: ${entry.recordId || '-'}`
    ];
    metaItems.forEach((text) => {
      const span = doc.createElement('span');
      span.textContent = text;
      meta.appendChild(span);
    });
    li.appendChild(meta);

    const note = doc.createElement('textarea');
    note.className = 'pin-note-input';
    note.placeholder = 'メモを入力…';
    note.value = entry.note || '';
    note.addEventListener('input', () => {
      entry.note = note.value;
      schedulePersist();
    });
    note.addEventListener('blur', () => {
      entry.note = note.value;
      schedulePersist();
    });
    li.appendChild(note);

    if (detail?.error) {
      const err = doc.createElement('div');
      err.className = 'pin-error';
      err.textContent = `⚠ ${detail.error}`;
      li.appendChild(err);
    }

    return li;
  }

  function renderList(detailsMap) {
    if (detailsMap instanceof Map) {
      state.detailMap = detailsMap;
    } else if (!(state.detailMap instanceof Map)) {
      state.detailMap = new Map();
    }

    pinListEl.innerHTML = '';
    if (!state.entries.length) {
      const empty = doc.createElement('li');
      empty.className = 'muted';
      empty.textContent = 'まだピン留めされたレコードはありません。設定ページから追加してください。';
      pinListEl.appendChild(empty);
      state.detailMap = new Map();
      return;
    }

    const map = state.detailMap;
    state.entries.forEach((entry) => {
      const card = createCard(entry, map.get(entry.id));
      pinListEl.appendChild(card);
    });
  }

  function attachStorageListener() {
    if (!chrome?.storage?.onChanged) return;
    const listener = async (changes, area) => {
      if (area !== 'sync' || !Object.prototype.hasOwnProperty.call(changes, 'kfavPins')) return;
      if (state.skipStorageOnce) {
        state.skipStorageOnce = false;
        return;
      }
      const next = changes.kfavPins.newValue;
      const arr = Array.isArray(next) ? next : next ? [next] : [];
      state.entries = arr.map(normalizePinnedEntry).filter(Boolean);
      const hosts = new Set(state.entries.map((entry) => entry.host).filter(Boolean));
      for (const host of hosts) {
        try {
          await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host });
        } catch (_warmErr) {}
      }
      renderList();
      refreshPins({ silent: true }).catch(() => {});
    };
    chrome.storage.onChanged.addListener(listener);
    state.storageListener = listener;
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

  async function quickAddCurrentRecord() {
    try {
      const tab = await getActiveTab();
      if (!tab || !isKintoneUrl(tab.url || '')) {
        setNotice('kintone のレコードを表示してから実行してください。', 2000);
        return false;
      }
      const parsed = parseKintoneUrl(tab.url || '');
      if (!parsed.host || !parsed.appId || !parsed.recordId) {
        setNotice('レコード詳細画面で実行してください。', 2000);
        return false;
      }
      if (state.entries.some((entry) => entry.host === parsed.host && entry.appId === parsed.appId && entry.recordId === parsed.recordId)) {
        setNotice('このレコードは既にピン留めされています。', 2000);
        return false;
      }
      const labelFromTab = (tab.title || '').replace(/\s+-\s+kintone.*/i, '').trim();
      const entry = normalizePinnedEntry({
        id: createId(),
        label: labelFromTab,
        host: parsed.host,
        appId: parsed.appId,
        recordId: parsed.recordId,
        titleField: '',
        note: ''
      });
      state.entries.push(entry);
      if (!(state.detailMap instanceof Map)) state.detailMap = new Map();
      state.detailMap.set(entry.id, state.detailMap.get(entry.id) || null);
      await persistPins();
      renderList();
      setNotice('ピン留めしました。', 2000);
      refreshPins({ silent: true }).catch(() => {});
      return true;
    } catch (error) {
      setNotice(String(error?.message || error), 2000);
      return false;
    }
  }

  async function removeEntry(id) {
    const idx = state.entries.findIndex((entry) => entry.id === id);
    if (idx === -1) return;
    state.entries.splice(idx, 1);
    if (state.detailMap instanceof Map) state.detailMap.delete(id);
    await persistPins();
    renderList();
    setNotice('ピンを外しました。', 2000);
    refreshPins({ silent: true }).catch(() => {});
  }

  function handleDragStart(event) {
    const sourceId = event.currentTarget.dataset.cardId
      || event.currentTarget.dataset.id
      || event.currentTarget.closest('.pin-card')?.dataset.id;
    dragSourceId = sourceId || null;
    if (!dragSourceId) return;
    event.stopPropagation();
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', dragSourceId);
    const card = pinListEl.querySelector(`.pin-card[data-id="${escapeSelector(dragSourceId)}"]`);
    card?.classList.add('dragging');
  }

  function handleDragEnter(event) {
    if (!dragSourceId) return;
    if (event.target !== event.currentTarget) return;
    if (event.currentTarget.dataset.id === dragSourceId) return;
    event.currentTarget.classList.add('drag-over');
  }

  function handleDragOver(event) {
    if (!dragSourceId) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
  }

  function handleDragLeave(event) {
    if (event.target !== event.currentTarget) return;
    event.currentTarget.classList.remove('drag-over');
  }

  async function handleDrop(event) {
    event.preventDefault();
    const targetId = event.currentTarget.dataset.id;
    event.currentTarget.classList.remove('drag-over');
    if (!dragSourceId || !targetId || dragSourceId === targetId) return;
    const fromIdx = state.entries.findIndex((entry) => entry.id === dragSourceId);
    const toIdx = state.entries.findIndex((entry) => entry.id === targetId);
    if (fromIdx === -1 || toIdx === -1) return;
    const [moved] = state.entries.splice(fromIdx, 1);
    state.entries.splice(toIdx, 0, moved);
    await persistPins();
    renderList(state.detailMap);
    dragSourceId = null;
    Array.from(pinListEl.querySelectorAll('.pin-card.dragging')).forEach((el) => el.classList.remove('dragging'));
  }

  function handleDragEnd(event) {
    event.stopPropagation();
    const card = pinListEl.querySelector('.pin-card.dragging');
    card?.classList.remove('dragging');
    dragSourceId = null;
    Array.from(pinListEl.querySelectorAll('.pin-card.drag-over')).forEach((el) => el.classList.remove('drag-over'));
  }

  async function refreshPins({ silent = false } = {}) {
    if (!state.entries.length) {
      if (!silent) setNotice('');
      renderList();
      return;
    }

    if (state.isRefreshing) return;
    state.isRefreshing = true;
    if (!silent) setNotice('更新中…');
    const token = ++state.fetchToken;
    const results = new Map();

    try {
      const byHost = new Map();
      state.entries.forEach((entry) => {
        if (!entry.host || !entry.appId || !entry.recordId) {
          results.set(entry.id, { error: '設定が不足しています（host / App ID / Record）' });
          return;
        }
        if (!byHost.has(entry.host)) byHost.set(entry.host, []);
        byHost.get(entry.host).push(entry);
      });

      for (const [host, entries] of byHost.entries()) {
        try {
          const granted = await ensureHostPermission(host);
          if (!granted) {
            entries.forEach((entry) => {
              results.set(entry.id, { error: 'このホストの許可が必要です' });
            });
            continue;
          }
        } catch (error) {
          entries.forEach((entry) => {
            results.set(entry.id, { error: String(error?.message || error) });
          });
          continue;
        }

        const payloadItems = entries.map((entry) => ({
          id: entry.id,
          appId: entry.appId,
          recordId: entry.recordId,
          fields: entry.titleField ? [entry.titleField] : []
        }));

        try {
          try {
            await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host });
          } catch (_warmErr) {}
          const res = await sendRunInKintone(host, { type: 'PIN_FETCH', payload: { items: payloadItems } });
          if (token !== state.fetchToken) return;
          if (res?.ok && res.results) {
            for (const entry of entries) {
              const detail = res.results[entry.id];
              if (detail?.ok) {
                results.set(entry.id, { record: detail.record });
              } else {
                results.set(entry.id, { error: detail?.error || '取得失敗' });
              }
            }
          } else {
            const errorText = res?.error || '取得失敗';
            entries.forEach((entry) => {
              results.set(entry.id, { error: errorText });
            });
          }
        } catch (error) {
          const message = String(error?.message || error);
          entries.forEach((entry) => {
            results.set(entry.id, { error: message });
          });
        }
      }

      if (token !== state.fetchToken) return;

      renderList(results);

      const errors = Array.from(results.values()).filter((detail) => detail?.error);
      if (!silent) {
        if (errors.length && errors.length === state.entries.length) {
          setNotice(errors[0].error);
        } else if (errors.length) {
          setNotice('一部のレコードで取得に失敗しました。');
        } else {
          setNotice('');
        }
      }
    } finally {
      state.isRefreshing = false;
    }
  }

  refreshBtn.addEventListener('click', () => {
    refreshPins().catch((error) => setNotice(String(error?.message || error)));
  });
  addBtn?.addEventListener('click', () => {
    quickAddCurrentRecord().catch((error) => setNotice(String(error?.message || error)));
  });

  async function boot() {
    state.entries = (await loadPinnedRecords()).map(normalizePinnedEntry).filter(Boolean);
    const hosts = new Set(state.entries.map((entry) => entry.host).filter(Boolean));
    for (const host of hosts) {
      try {
        await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host });
      } catch (_warmErr) {}
    }
    renderList();
    await refreshPins();
  }

  attachStorageListener();
  boot().catch((error) => setNotice(String(error?.message || error)));

  return {
    refreshPins,
    quickAdd: quickAddCurrentRecord,
    dispose() {
      if (state.storageListener) {
        chrome.storage.onChanged.removeListener(state.storageListener);
        state.storageListener = null;
      }
      if (noticeTimer) {
        clearTimeout(noticeTimer);
        noticeTimer = null;
      }
      if (persistTimer) {
        clearTimeout(persistTimer);
        persistTimer = null;
      }
    }
  };
}
