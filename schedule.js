'use strict';

import {
  loadScheduleConfig,
  saveScheduleConfig,
  hasHostPermission,
  ensureHostPermission,
  sendRunInKintone,
  createId
} from './core.js';

export function initSchedule(options = {}) {
  const doc = options.document || document;
  const scheduleListEl = doc.getElementById('scheduleList');
  const scheduleNoticeEl = doc.getElementById('scheduleNotice');
  const refreshScheduleBtn = doc.getElementById('refreshSchedule');
  const scheduleEntrySel = doc.getElementById('scheduleEntrySelect');
  const scheduleTitleSel = doc.getElementById('scheduleTitleField');
  const detectScheduleFieldsBtn = doc.getElementById('detectScheduleFields');

  if (!scheduleListEl || !refreshScheduleBtn) {
    return {
      dispose() {},
      async refreshSchedule() {}
    };
  }

  const state = {
    refreshInterval: options.refreshInterval ?? 120000,
    timer: null,
    entries: [],
    currentEntryId: null,
    storageListener: null,
    titleFetchToken: 0,
    scheduleFetchToken: 0
  };

  function normalizeEntry(entry) {
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

  function serializeEntries() {
    return state.entries.map(({ id, label, host, appId, dateField, titleField, extraQuery, limit }) => ({
      id,
      label,
      host,
      appId,
      dateField,
      titleField,
      extraQuery,
      limit
    }));
  }

  function entryLabel(entry) {
    const label = entry.label?.trim();
    if (label) return label;
    const app = entry.appId ? `App ${entry.appId}` : 'App 未指定';
    return `${app} @ ${entry.host}`;
  }

  function getCurrentEntry() {
    return state.entries.find((entry) => entry.id === state.currentEntryId) || null;
  }

  function renderScheduleNotice(msg) {
    if (!scheduleNoticeEl) return;
    if (msg) {
      scheduleNoticeEl.textContent = msg;
      scheduleNoticeEl.classList.remove('hidden');
    } else {
      scheduleNoticeEl.textContent = '';
      scheduleNoticeEl.classList.add('hidden');
    }
  }

  function updateControlsState() {
    const hasEntries = state.entries.length > 0;
    const hasCurrent = Boolean(getCurrentEntry());
    if (scheduleEntrySel) scheduleEntrySel.disabled = !hasEntries;
    if (scheduleTitleSel) scheduleTitleSel.disabled = !hasCurrent;
    if (detectScheduleFieldsBtn) detectScheduleFieldsBtn.disabled = !hasCurrent;
    if (refreshScheduleBtn) refreshScheduleBtn.disabled = !hasCurrent;
  }

  function renderEntryOptions() {
    if (!scheduleEntrySel) return;
    scheduleEntrySel.innerHTML = '';
    if (!state.entries.length) {
      const opt = new Option('（設定がありません）', '', true, true);
      scheduleEntrySel.appendChild(opt);
      updateControlsState();
      scheduleListEl.innerHTML = '';
      renderScheduleNotice('設定ページで直近スケジュールを追加してください。');
      return;
    }

    state.entries.forEach((entry) => {
      const opt = new Option(entryLabel(entry), entry.id, false, entry.id === state.currentEntryId);
      scheduleEntrySel.appendChild(opt);
    });
    if (!state.entries.some((entry) => entry.id === state.currentEntryId)) {
      state.currentEntryId = state.entries[0].id;
      scheduleEntrySel.value = state.currentEntryId;
    }
    updateControlsState();
  }

  function renderSchedule(records, entry) {
    scheduleListEl.innerHTML = '';
    if (!entry) return;

    const header = doc.createElement('li');
    header.className = 'sch-entry';
    header.textContent = `${entryLabel(entry)} — ${entry.dateField || '日付フィールド未設定'}`;
    scheduleListEl.appendChild(header);

    const rows = (records || []).filter((record) => record?.[entry.dateField]?.value);
    if (!rows.length) {
      const li = doc.createElement('li');
      li.className = 'muted';
      li.style.padding = '8px';
      li.textContent = '該当なし';
      scheduleListEl.appendChild(li);
      return;
    }

    rows.forEach((record) => {
      const id = record.$id?.value;
      const dateRaw = record[entry.dateField]?.value || '';
      const dateText = formatDateValue(dateRaw);
      const title = entry.titleField
        ? (String(record[entry.titleField]?.value || '').trim() || `#${id}`)
        : `#${id}`;

      const li = doc.createElement('li');
      li.className = 'sch-item';

      const dateDiv = doc.createElement('div');
      dateDiv.className = 'sch-date';
      dateDiv.textContent = dateText;

      const titleDiv = doc.createElement('div');
      titleDiv.className = 'sch-title';
      titleDiv.textContent = title;

      const openDiv = doc.createElement('div');
      openDiv.className = 'sch-open';
      const anchor = doc.createElement('a');
      anchor.textContent = '開く';
      anchor.target = '_blank';
      anchor.rel = 'noopener';
      anchor.href = recordUrlOf(entry.host, entry.appId, id);
      openDiv.appendChild(anchor);

      li.appendChild(dateDiv);
      li.appendChild(titleDiv);
      li.appendChild(openDiv);
      scheduleListEl.appendChild(li);
    });
  }

  function formatDateValue(value) {
    if (!value) return '—';
    if (/^\d{4}-\d{2}-\d{2}T/.test(value)) {
      const date = new Date(value);
      if (!Number.isNaN(date.getTime())) {
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        const hh = String(date.getHours()).padStart(2, '0');
        const mm = String(date.getMinutes()).padStart(2, '0');
        return `${y}-${m}-${d} ${hh}:${mm}`;
      }
      return value;
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
    return value;
  }

  function recordUrlOf(host, appId, recId) {
    return `${host}/k/${appId}/show#record=${recId}`;
  }

  async function populateScheduleTitleCandidates({ force = false } = {}) {
    if (!scheduleTitleSel) return;
    const entry = getCurrentEntry();
    scheduleTitleSel.innerHTML = '';
    if (!entry) {
      scheduleTitleSel.disabled = true;
      return;
    }

    scheduleTitleSel.disabled = false;
    const baseOption = doc.createElement('option');
    baseOption.value = '';
    baseOption.textContent = '（レコード番号）';
    scheduleTitleSel.appendChild(baseOption);

    if (!entry.host || !entry.appId) {
      scheduleTitleSel.value = entry.titleField || '';
      return;
    }

    const token = ++state.titleFetchToken;
    const permitted = await hasHostPermission(entry.host);
    if (!permitted) {
      scheduleTitleSel.value = entry.titleField || '';
      return;
    }

    try {
      if (!force && entry._fieldCache) {
        entry._fieldCache.forEach(({ code, label }) => {
          const opt = doc.createElement('option');
          opt.value = code;
          opt.textContent = `${label} (${code})`;
          scheduleTitleSel.appendChild(opt);
        });
        scheduleTitleSel.value = entry.titleField || '';
        return;
      }

      const res = await sendRunInKintone(entry.host, {
        type: 'LIST_FIELDS',
        payload: { appId: entry.appId }
      });
      if (token !== state.titleFetchToken) return;
      if (!res?.ok) throw new Error(res?.error || 'LIST_FIELDS failed');
      const fields = (res.fields || []).filter((field) => ['SINGLE_LINE_TEXT', 'MULTI_LINE_TEXT', 'RICH_TEXT', 'LINK'].includes(field.type));
      entry._fieldCache = fields.map((field) => ({ code: field.code, label: field.label }));
      entry._fieldCache.forEach(({ code, label }) => {
        const opt = doc.createElement('option');
        opt.value = code;
        opt.textContent = `${label} (${code})`;
        scheduleTitleSel.appendChild(opt);
      });
      scheduleTitleSel.value = entry.titleField || '';
    } catch (error) {
      if (token !== state.titleFetchToken) return;
      renderScheduleNotice(String(error?.message || error));
    }
  }

  async function refreshSchedule() {
    const entry = getCurrentEntry();
    if (!entry) {
      scheduleListEl.innerHTML = '';
      renderScheduleNotice('設定が未完了です（設定→「直近スケジュール設定」）。');
      return;
    }

    if (!entry.host || !entry.appId || !entry.dateField) {
      scheduleListEl.innerHTML = '';
      renderScheduleNotice('host / App ID / 日付フィールドを設定してください。');
      return;
    }

    renderScheduleNotice('更新中…');
    const token = ++state.scheduleFetchToken;

    try {
      const granted = await ensureHostPermission(entry.host);
      if (token !== state.scheduleFetchToken) return;
      if (!granted) {
        renderScheduleNotice('このホストの許可が必要です。⟳を押して許可してください。');
        scheduleListEl.innerHTML = '';
        return;
      }
    } catch (_error) {
      // ignore and continue
    }

    try {
      try {
        await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host: entry.host });
      } catch (_warmErr) {}
      const end = new Date();
      end.setDate(end.getDate() + 7);
      const response = await sendRunInKintone(entry.host, {
        type: 'LIST_SCHEDULE',
        payload: {
          appId: entry.appId,
          dateField: entry.dateField,
          titleField: entry.titleField || '',
          extraQuery: entry.extraQuery || '',
          endDate: ymd(end),
          limit: entry.limit || 20
        }
      });
      if (token !== state.scheduleFetchToken) return;
      if (response?.ok) {
        renderSchedule(response.records || [], entry);
        renderScheduleNotice('');
      } else {
        scheduleListEl.innerHTML = '';
        renderScheduleNotice(response?.error || '取得に失敗しました');
      }
    } catch (error) {
      if (token !== state.scheduleFetchToken) return;
      scheduleListEl.innerHTML = '';
      renderScheduleNotice(String(error?.message || error));
    }
  }

  function ymd(dt) {
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, '0');
    const d = String(dt.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  async function persistEntries() {
    await saveScheduleConfig(serializeEntries());
  }

  function attachStorageListener() {
    if (!chrome?.storage?.onChanged) return;
    const listener = async (changes, area) => {
      if (area !== 'sync') return;
      if (!Object.prototype.hasOwnProperty.call(changes, 'kfavSchedule')) return;
      const nextValue = changes.kfavSchedule.newValue;
      const arr = Array.isArray(nextValue) ? nextValue : nextValue ? [nextValue] : [];
      const normalized = arr.map(normalizeEntry).filter(Boolean);
      state.entries = normalized;
      if (!state.entries.some((entry) => entry.id === state.currentEntryId)) {
        state.currentEntryId = state.entries[0]?.id || null;
      }
      const hosts = new Set(state.entries.map((entry) => entry.host).filter(Boolean));
      for (const host of hosts) {
        try {
          await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host });
        } catch (_warmErr) {}
      }
      renderEntryOptions();
      populateScheduleTitleCandidates().then(() => refreshSchedule()).catch(() => {});
    };
    chrome.storage.onChanged.addListener(listener);
    state.storageListener = listener;
  }

  function wireEvents() {
    if (scheduleEntrySel) {
      scheduleEntrySel.addEventListener('change', () => {
        state.currentEntryId = scheduleEntrySel.value || null;
        populateScheduleTitleCandidates().then(() => refreshSchedule()).catch(() => {});
      });
    }

    if (scheduleTitleSel) {
      scheduleTitleSel.addEventListener('change', async () => {
        const entry = getCurrentEntry();
        if (!entry) return;
        entry.titleField = scheduleTitleSel.value || '';
        await persistEntries();
        await refreshSchedule();
      });
    }

    detectScheduleFieldsBtn?.addEventListener('click', async () => {
      const entry = getCurrentEntry();
      if (!entry) return;
      await populateScheduleTitleCandidates({ force: true });
    });

    refreshScheduleBtn.addEventListener('click', () => {
      refreshSchedule().catch(() => {});
    });
  }

  async function boot() {
    const loaded = (await loadScheduleConfig()).map(normalizeEntry).filter(Boolean);
    state.entries = loaded;
    state.currentEntryId = state.entries[0]?.id || null;
    const hosts = new Set(state.entries.map((entry) => entry.host).filter(Boolean));
    for (const host of hosts) {
      try {
        await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host });
      } catch (_warmErr) {}
    }
    renderEntryOptions();
    updateControlsState();
    await populateScheduleTitleCandidates();
    await refreshSchedule();
    if (state.refreshInterval > 0) {
      state.timer = setInterval(() => {
        refreshSchedule().catch(() => {});
      }, state.refreshInterval);
    }
  }

  wireEvents();
  attachStorageListener();
  boot().catch((error) => {
    renderScheduleNotice(String(error?.message || error));
  });

  return {
    refreshSchedule,
    dispose() {
      if (state.timer) {
        clearInterval(state.timer);
        state.timer = null;
      }
      if (state.storageListener) {
        chrome.storage.onChanged.removeListener(state.storageListener);
        state.storageListener = null;
      }
    }
  };
}
