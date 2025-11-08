// content.js
// 1) page-bridge.js 繧偵・繝ｼ繧ｸ縺ｫ豕ｨ蜈･  2) popup/background 縺九ｉ縺ｮ隕∵ｱゅｒ window.postMessage 縺ｧ讖区ｸ｡縺・
(function bootstrap() {
  if (window.__kfav_content_injected) return;
  window.__kfav_content_injected = true;

  (function injectBridge() {
    const s = document.createElement('script');
    s.src = chrome.runtime.getURL('page-bridge.js');
    s.async = false;
    (document.head || document.documentElement).appendChild(s);
    s.onload = () => s.remove();
  })();

  const pendingKey = '__kfav_pending_map';
  const pending = window[pendingKey] instanceof Map ? window[pendingKey] : new Map();
  window[pendingKey] = pending;

  function uuid() {
    return crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2);
  }

  window.addEventListener('message', (ev) => {
    if (ev.origin !== location.origin) return;
    const data = ev.data || {};
    if (!data || data.__kfav__ !== true || !data.replyTo) return;
    const entry = pending.get(data.replyTo);
    if (entry) {
      pending.delete(data.replyTo);
      entry.resolve(data);
    }
  });

  function postToPage(type, payload) {
    const id = uuid();
    const p = new Promise((resolve) => {
      pending.set(id, { resolve });
      setTimeout(() => {
        if (pending.has(id)) {
          pending.delete(id);
          resolve({ ok: false, error: 'Timeout' });
        }
      }, 10000);
    });
    window.postMessage({ __kfav__: true, id, type, payload }, location.origin);
    return p;
  }

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    (async () => {
      try {
        if (msg?.type === 'COUNT_BULK') {
          const res = await postToPage('COUNT_BULK', msg.payload);
          sendResponse(res); return;
        }
        if (msg?.type === 'PING_CONTENT') {
          sendResponse({ ok: true }); return;
        }
        if (msg?.type === 'COUNT_VIEW') {
          const res = await postToPage('COUNT_VIEW', msg.payload);
          sendResponse(res); return;
        }
        if (msg?.type === 'COUNT_APP_QUERY') {
          const res = await postToPage('COUNT_APP_QUERY', msg.payload);
          sendResponse(res); return;
        }
        if (msg?.type === 'LIST_VIEWS') {
          const res = await postToPage('LIST_VIEWS', msg.payload);
          sendResponse(res); return;
        }
        if (msg?.type === 'LIST_SCHEDULE') {
          const res = await postToPage('LIST_SCHEDULE', msg.payload);
          sendResponse(res); return;
        }
        if (msg?.type === 'LIST_FIELDS') {
          const res = await postToPage('LIST_FIELDS', msg.payload);
          sendResponse(res); return;
        }
        if (msg?.type === 'PIN_FETCH') {
          const res = await postToPage('PIN_FETCH', msg.payload);
          sendResponse(res); return;
        }
        if (msg?.type === 'CHECK_KINTONE_READY') {
          const res = await postToPage('CHECK_KINTONE_READY', msg.payload);
          sendResponse(res); return;
        }
        sendResponse({ ok: false, error: 'Unknown message type' });
      } catch (e) {
        sendResponse({ ok: false, error: String(e?.message || e) });
      }
    })();
    return true;
  });

  const overlayTexts = {
    ja: {
      title: "Excelモード",
      loading: "読み込み中…",
      loadFailed: "データの取得に失敗しました",
      noEditableFields: "編集可能なフィールドがありません",
      btnSave: "保存",
      btnClose: "閉じる",
      toastNoChanges: "変更はありません",
      toastInvalidCells: "エラーを解消してから保存してください",
      toastSaveSuccess: "保存しました",
      toastSaveFailed: "保存に失敗しました",
      confirmClose: "未保存の変更があります。閉じますか？",
      reloading: "一覧を更新しています…",
      conflictRetry: "最新のデータを反映して再保存します…",
      conflictNoChanges: "最新のデータと一致しました",
      conflictFailed: "他の変更と競合し保存できませんでした",
      toastCopySuccess: "コピーしました",
      toastCopyFailed: "コピーに失敗しました",
      dirtyLabel: (n) => `未保存: ${n}`,
      statusSum: "合計",
      statusAvg: "平均",
      statusCount: "件数",
      btnColumns: "列順",
      columnsTitle: "列の並び替え",
      columnsHint: "ヘッダーをドラッグして並び順を変え、幅はドラッグまたは自動調整ボタンで変更できます。",
      columnsAutoWidth: "自動調整",
      columnsSave: "列順を保存",
      columnsReset: "リセット",
      columnsClose: "閉じる",
      columnsEmpty: "列情報がありません",
      toastColumnsSaved: "列順を保存しました",
      toastColumnsReset: "列順をリセットしました",
      toastColumnsAutoWidth: "列幅を自動調整しました",
      errorUnsupportedFilter: "この一覧には一時的なフィルター/ソートが含まれているため、Excelモードを起動できません。ビューを標準状態に戻してから再試行してください。"
    },
    en: {
      title: "Excel mode",
      loading: "Loading…",
      loadFailed: "Failed to load data",
      noEditableFields: "No editable fields found",
      btnSave: "Save",
      btnClose: "Close",
      toastNoChanges: "No changes",
      toastInvalidCells: "Fix errors before saving",
      toastSaveSuccess: "Changes saved",
      toastSaveFailed: "Failed to save changes",
      confirmClose: "You have unsaved changes. Close anyway?",
      reloading: "Refreshing list…",
      conflictRetry: "Data updated, retrying save…",
      conflictNoChanges: "Your edits now match the latest data",
      conflictFailed: "Conflicted with other changes; save aborted",
      toastCopySuccess: "Copied",
      toastCopyFailed: "Copy failed",
      dirtyLabel: (n) => `Unsaved: ${n}`,
      statusSum: "Sum",
      statusAvg: "Avg",
      statusCount: "Count",
      btnColumns: "Columns",
      columnsTitle: "Reorder columns",
      columnsHint: "Drag the header chips below to reorder columns, then save the layout for future sessions.",
      columnsAutoWidth: "Auto width",
      columnsSave: "Save order",
      columnsReset: "Reset",
      columnsClose: "Close",
      columnsEmpty: "No columns available",
      toastColumnsSaved: "Column order saved",
      toastColumnsReset: "Column order reset",
      toastColumnsAutoWidth: "Column widths adjusted",
      errorUnsupportedFilter: "This list has ad-hoc filters/sorting that cannot be replayed in Excel mode. Clear the filter and try again."
    }
  };

  const overlayConfig = {
    limit: 100,
    allowedTypes: new Set(['SINGLE_LINE_TEXT', 'NUMBER', 'DATE', 'RADIO_BUTTON'])
  };

  const COLUMN_PREF_STORAGE_KEY = 'kfavExcelColumns';
  const MAX_COLUMN_PREF_ENTRIES = 80;
  const DEFAULT_COLUMN_WIDTH = 160;
  const MIN_COLUMN_WIDTH = 96;
  const MAX_COLUMN_WIDTH = 320;

  let overlayCssLoaded = false;

  function resolveText(language, key, ...args) {
    const table = overlayTexts[language] || overlayTexts.en;
    const value = table[key];
    if (typeof value === 'function') return value(...args);
    if (typeof value === 'string') return value;
    const fallback = overlayTexts.en[key];
    return typeof fallback === 'function' ? fallback(...args) : (fallback || '');
  }

  function columnLabel(index) {
    let n = index;
    let label = '';
    while (n >= 0) {
      const char = String.fromCharCode(65 + (n % 26));
      label = char + label;
      n = Math.floor(n / 26) - 1;
    }
    return label;
  }

  class ExcelOverlayController {
    constructor(postFn) {
      this.postFn = postFn;
      this.isOpen = false;
      this.root = null;
      this.surface = null;
      this.bodyScroll = null;
      this.columnHeaderScroll = null;
      this.rowHeaderScroll = null;
      this.toastHost = null;
      this.loadingEl = null;
      this.dirtyBadge = null;
      this.saveButton = null;
      this.closeButton = null;
      this.titleElement = null;
      this.appName = '';
      this.statsElements = { sum: null, avg: null, count: null };
      this.language = 'ja';
      this.appName = '';
      this.appId = null;
      this.query = '';
      this.fields = [];
      this.fieldMap = new Map();
      this.fieldIndexMap = new Map();
      this.rows = [];
      this.rowMap = new Map();
      this.rowIndexMap = new Map();
      this.diff = new Map();
      this.invalidCells = new Set();
      this.revisionMap = new Map();
      this.isComposing = false;
      this.saving = false;
      this.handleScroll = this.syncScroll.bind(this);
      this.handleKeyOnOverlay = this.handleOverlayKeydown.bind(this);
      this.handleCompositionStart = () => { this.isComposing = true; };
      this.handleCompositionEnd = () => { this.isComposing = false; };
      this.requireReload = false;
      this.previousFocus = null;
      this.bodyOverflowBackup = '';
      this.reloading = false;
      this.rowHeight = 32;
      this.virtualStart = 0;
      this.virtualEnd = 0;
      this.rowHeaderContainer = null;
      this.tableContainer = null;
      this.inputsByRow = new Map();
      this.pendingFocus = null;
      this.invalidValueCache = new Map();
      this.selection = null;
      this.dragSelecting = false;
      this.editingCell = null;
      this.handleMouseUp = this.endSelection.bind(this);
      this.columnPrefCache = null;
      this.viewKey = 'default';
      this.viewName = '';
      this.viewId = '';
      this.columnPanelLayer = null;
      this.columnPreviewRow = null;
      this.columnOrderDraft = [];
      this.columnButton = null;
      this.defaultFieldCodes = [];
      this.draggingColumnCode = null;
      this.columnWidthsDraft = new Map();
      this.radioPicker = null;
      this.boundRadioPickerScrollHandler = () => { this.repositionRadioPicker(); };
      this.boundRadioPickerResizeHandler = () => { this.repositionRadioPicker(); };
      this.boundRadioPickerOutsideClick = (event) => {
        if (!this.radioPicker) return;
        const panel = this.radioPicker.panel;
        const input = this.radioPicker.input;
        if (!panel) return;
        if (panel.contains(event.target) || input === event.target) return;
        this.closeRadioPicker();
      };
    }

    async open() {
      if (this.isOpen) return;
      this.isOpen = true;
      try {
        this.previousFocus = document.activeElement instanceof HTMLElement ? document.activeElement : null;
        this.bodyOverflowBackup = document.body.style.overflow || '';
        document.body.style.overflow = 'hidden';
        this.requireReload = false;
        await this.ensureStyles();
        await this.loadLanguage();
        this.mountShell();
        this.showLoading(resolveText(this.language, 'loading'));
        await this.fetchData();
        this.hideLoading();
        if (!this.fields.length) {
          this.renderNoFields();
          return;
        }
        this.renderGrid();
        this.attachEvents();
        this.updateDirtyBadge();
        this.updateStats();
        this.focusFirstCell();
      } catch (error) {
        this.hideLoading();
        this.notify(resolveText(this.language, 'loadFailed'));
        console.error('[kintone-excel-overlay] failed to open overlay', error);
        this.close(true);
      }
    }

    async ensureStyles() {
      if (overlayCssLoaded) return;
      overlayCssLoaded = true;
      const href = chrome.runtime.getURL('overlay.css');
      const link = document.createElement('link');
      link.rel = 'stylesheet';
      link.href = href;
      link.dataset.pbOverlay = 'true';
      document.head.appendChild(link);
    }

    mountShell() {
      this.root = document.createElement('div');
      this.root.id = 'pb-excel-overlay';
      this.root.className = 'pb-overlay';

      const surface = document.createElement('div');
      surface.className = 'pb-overlay__surface';
      this.surface = surface;

      const toolbar = this.buildToolbar();
      const grid = this.buildGridShell();
      const status = this.buildStatusBar();
      const toastHost = document.createElement('div');
      toastHost.className = 'pb-overlay__toast-host';
      this.toastHost = toastHost;

      surface.appendChild(toolbar);
      surface.appendChild(grid);
      surface.appendChild(status);
      surface.appendChild(toastHost);
      this.root.appendChild(surface);
      document.body.appendChild(this.root);
    }

    buildToolbar() {
      const toolbar = document.createElement('div');
      toolbar.className = 'pb-overlay__toolbar';

      const titleWrap = document.createElement('div');
      titleWrap.className = 'pb-overlay__toolbar-left';
      const title = document.createElement('span');
      title.className = 'pb-overlay__title';
      title.textContent = this.appName || resolveText(this.language, 'title');
      this.titleElement = title;
      titleWrap.appendChild(title);

      const actions = document.createElement('div');
      actions.className = 'pb-overlay__toolbar-actions';

      const columnsBtn = document.createElement('button');
      columnsBtn.type = 'button';
      columnsBtn.className = 'pb-overlay__btn';
      columnsBtn.textContent = resolveText(this.language, 'btnColumns');
      columnsBtn.disabled = true;
      columnsBtn.addEventListener('click', () => { this.openColumnManager(); });
      this.columnButton = columnsBtn;

      const dirty = document.createElement('span');
      dirty.className = 'pb-overlay__dirty';
      dirty.textContent = resolveText(this.language, 'dirtyLabel', 0);
      this.dirtyBadge = dirty;

      const saveBtn = document.createElement('button');
      saveBtn.type = 'button';
      saveBtn.className = 'pb-overlay__btn pb-overlay__btn--primary';
      saveBtn.textContent = resolveText(this.language, 'btnSave');
      this.saveButton = saveBtn;

      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.className = 'pb-overlay__btn';
      closeBtn.textContent = resolveText(this.language, 'btnClose');
      this.closeButton = closeBtn;

      actions.appendChild(columnsBtn);
      actions.appendChild(dirty);
      actions.appendChild(saveBtn);
      actions.appendChild(closeBtn);

      toolbar.appendChild(titleWrap);
      toolbar.appendChild(actions);
      return toolbar;
    }

    buildGridShell() {
      const grid = document.createElement('div');
      grid.className = 'pb-overlay__grid';

      const head = document.createElement('div');
      head.className = 'pb-overlay__grid-head';

      const corner = document.createElement('div');
      corner.className = 'pb-overlay__corner';
      head.appendChild(corner);

      const colScroll = document.createElement('div');
      colScroll.className = 'pb-overlay__col-scroll';

      const colRow = document.createElement('div');
      colRow.className = 'pb-overlay__col-row';
      colScroll.appendChild(colRow);
      head.appendChild(colScroll);
      this.columnHeaderScroll = colScroll;

      const body = document.createElement('div');
      body.className = 'pb-overlay__grid-body';

      const rowScroll = document.createElement('div');
      rowScroll.className = 'pb-overlay__row-scroll';
      const rowsWrap = document.createElement('div');
      rowsWrap.className = 'pb-overlay__row-wrap';
      rowScroll.appendChild(rowsWrap);
      this.rowHeaderScroll = rowScroll;

      const cellScroll = document.createElement('div');
      cellScroll.className = 'pb-overlay__cell-scroll';
      cellScroll.tabIndex = 0;
      const cellTable = document.createElement('div');
      cellTable.className = 'pb-overlay__table';
      cellScroll.appendChild(cellTable);
      this.bodyScroll = cellScroll;

      body.appendChild(rowScroll);
      body.appendChild(cellScroll);

      grid.appendChild(head);
      grid.appendChild(body);

      return grid;
    }

    buildStatusBar() {
      const status = document.createElement('div');
      status.className = 'pb-overlay__status';

      const sum = document.createElement('span');
      sum.className = 'pb-overlay__status-item';
      this.statsElements.sum = sum;

      const avg = document.createElement('span');
      avg.className = 'pb-overlay__status-item';
      this.statsElements.avg = avg;

      const count = document.createElement('span');
      count.className = 'pb-overlay__status-item';
      this.statsElements.count = count;

      status.appendChild(sum);
      status.appendChild(avg);
      status.appendChild(count);

      return status;
    }

    attachEvents() {
      if (!this.root) return;
      this.bodyScroll.addEventListener('scroll', this.handleScroll, { passive: true });
      this.saveButton.addEventListener('click', () => { this.save(); });
      this.closeButton.addEventListener('click', () => { this.tryClose(); });
      this.root.addEventListener('keydown', this.handleKeyOnOverlay, true);
      document.addEventListener('mouseup', this.handleMouseUp, true);
    }

    detachEvents() {
      if (!this.root) return;
      if (this.bodyScroll) {
        this.bodyScroll.removeEventListener('scroll', this.handleScroll);
      }
      this.root.removeEventListener('keydown', this.handleKeyOnOverlay, true);
      document.removeEventListener('mouseup', this.handleMouseUp, true);
    }

    getCurrentViewId() {
      const appApi = typeof window !== 'undefined' && window.kintone && window.kintone.app ? window.kintone.app : null;
      if (appApi && typeof appApi.getSelectedViewId === 'function') {
        const value = appApi.getSelectedViewId();
        if (value !== undefined && value !== null) return String(value);
      }
      const search = location.search;
      if (search) {
        const params = new URLSearchParams(search);
        const viewParam = params.get('view');
        if (viewParam) return viewParam;
      }
      const hash = location.hash || '';
      const match = hash.match(/[?&]view=(\d+)/);
      if (match) return match[1];
      return null;
    }

    async loadLanguage() {
      try {
        const res = await this.postFn('EXCEL_GET_LOGIN_USER');
        if (res?.ok && res.user?.language) {
          const lang = String(res.user.language).toLowerCase();
          this.language = overlayTexts[lang] ? lang : 'en';
        } else {
          this.language = 'en';
        }
      } catch (_e) {
        this.language = 'en';
      }
    }

    async fetchData() {
      const contextRes = await this.postFn('EXCEL_GET_APP_CONTEXT');
      if (!contextRes?.ok || !contextRes.appId) {
        throw new Error('Context unavailable');
      }
      this.appId = contextRes.appId;
      this.appName = contextRes.appName || '';
      if (this.titleElement) {
        this.titleElement.textContent = this.appName || resolveText(this.language, 'title');
      }
      const currentQuery = contextRes.query || '';
      const viewId = this.getCurrentViewId();
      this.viewId = '';
      this.viewName = '';
      this.viewKey = 'default';
      let resolvedQuery = currentQuery;
      let viewFieldOrder = [];
      try {
        const viewInfo = await this.postFn('EXCEL_GET_VIEW_INFO', {
          appId: this.appId,
          viewId,
          currentQuery
        });
        if (viewInfo?.ok) {
          if (typeof viewInfo.query === 'string') {
            resolvedQuery = viewInfo.query;
          }
          if (Array.isArray(viewInfo.fieldOrder)) {
            viewFieldOrder = viewInfo.fieldOrder;
          }
          if (viewInfo.viewId) {
            this.viewId = String(viewInfo.viewId);
          }
          if (viewInfo.viewName) {
            this.viewName = String(viewInfo.viewName);
          }
        }
      } catch (_e) {
        // ignore and fall back to currentQuery
      }
      this.viewKey = this.computeViewKey(this.viewId, this.viewName, viewId);
      this.query = this.applyLimit(resolvedQuery, overlayConfig.limit);

      const fieldsRes = await this.postFn('EXCEL_GET_FIELDS', { appId: this.appId });
      if (!fieldsRes?.ok) throw new Error('Failed to get fields');
      const rawFields = Array.isArray(fieldsRes.fields)
        ? fieldsRes.fields.filter((f) => overlayConfig.allowedTypes.has(f.type))
        : [];

      const sorted = rawFields.slice().sort((a, b) => {
        const aIdx = viewFieldOrder.indexOf(a.code);
        const bIdx = viewFieldOrder.indexOf(b.code);
        const ap = aIdx === -1 ? Number.MAX_SAFE_INTEGER : aIdx;
        const bp = bIdx === -1 ? Number.MAX_SAFE_INTEGER : bIdx;
        if (ap !== bp) return ap - bp;
        return rawFields.indexOf(a) - rawFields.indexOf(b);
      });
      this.defaultFieldCodes = sorted.map((field) => field.code);
      const savedPref = await this.getSavedColumnPref(this.defaultFieldCodes);
      const savedOrder = savedPref?.order || null;
      if (savedOrder?.length) {
        const savedIndex = new Map(savedOrder.map((code, index) => [code, index]));
        const defaultIndex = new Map(this.defaultFieldCodes.map((code, index) => [code, index]));
        sorted.sort((a, b) => {
          const ap = savedIndex.has(a.code) ? savedIndex.get(a.code) : Number.MAX_SAFE_INTEGER;
          const bp = savedIndex.has(b.code) ? savedIndex.get(b.code) : Number.MAX_SAFE_INTEGER;
          if (ap !== bp) return ap - bp;
          const da = defaultIndex.get(a.code) ?? Number.MAX_SAFE_INTEGER;
          const db = defaultIndex.get(b.code) ?? Number.MAX_SAFE_INTEGER;
          return da - db;
        });
      }
      const widthMap = new Map();
      if (savedPref?.widths) {
        Object.entries(savedPref.widths).forEach(([code, value]) => {
          const numeric = Number(value);
          if (Number.isFinite(numeric)) {
            widthMap.set(code, Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, numeric)));
          }
        });
      }
      this.fields = sorted.map((field) => ({
        ...field,
        choices: Array.isArray(field.choices) ? field.choices.map((choice) => String(choice)) : [],
        width: widthMap.get(field.code) || DEFAULT_COLUMN_WIDTH
      }));
      this.rebuildFieldMaps();

      const fieldCodes = this.fields.map((f) => f.code);
      const recordsRes = await this.postFn('EXCEL_GET_RECORDS', {
        appId: this.appId,
        fields: fieldCodes,
        query: this.query
      });

      if (!recordsRes?.ok) {
        const errorCode = recordsRes?.errorCode || '';
        if (errorCode === 'GAIA_IL02' || /GAIA_IL02/.test(recordsRes?.error || '')) {
          throw new Error(resolveText(this.language, 'errorUnsupportedFilter'));
        }
        throw new Error(recordsRes?.error || 'Failed to get records');
      }

      const records = Array.isArray(recordsRes.records) ? recordsRes.records : [];
      this.rows = records.map((record) => this.transformRecord(record));
      this.rowMap.clear();
      this.rowIndexMap.clear();
      this.revisionMap.clear();
      this.rows.forEach((row, index) => {
        this.rowMap.set(row.id, row);
        this.rowIndexMap.set(row.id, index);
        this.revisionMap.set(row.id, row.revision);
      });
      if (this.columnButton) {
        this.columnButton.disabled = !this.fields.length;
      }
    }

    computeViewKey(viewIdValue, viewNameValue, fallbackParam) {
      if (viewIdValue) return `id:${viewIdValue}`;
      if (viewNameValue) return `name:${viewNameValue}`;
      if (fallbackParam) return `param:${fallbackParam}`;
      return 'default';
    }

    rebuildFieldMaps() {
      this.fieldMap.clear();
      this.fieldIndexMap.clear();
      this.fields.forEach((field, index) => {
        this.fieldMap.set(field.code, field);
        this.fieldIndexMap.set(field.code, index);
      });
    }

    async loadColumnPrefMap(force = false) {
      if (!force && this.columnPrefCache) return this.columnPrefCache;
      try {
        const stored = await chrome.storage.sync.get(COLUMN_PREF_STORAGE_KEY);
        const raw = stored?.[COLUMN_PREF_STORAGE_KEY];
        const map = raw && typeof raw === 'object' ? raw : {};
        this.columnPrefCache = { ...map };
        return this.columnPrefCache;
      } catch (_err) {
        this.columnPrefCache = {};
        return this.columnPrefCache;
      }
    }

    async writeColumnPrefMap(map) {
      if (!map || !Object.keys(map).length) {
        await chrome.storage.sync.remove(COLUMN_PREF_STORAGE_KEY);
        this.columnPrefCache = {};
        return;
      }
      await chrome.storage.sync.set({ [COLUMN_PREF_STORAGE_KEY]: map });
      this.columnPrefCache = { ...map };
    }

    getColumnPrefKey() {
      const app = this.appId ? String(this.appId) : '0';
      const viewKey = this.viewKey || 'default';
      return `${app}::${viewKey}`;
    }

    async getSavedColumnPref(defaultCodes = []) {
      const map = await this.loadColumnPrefMap();
      const entry = map[this.getColumnPrefKey()];
      if (!entry) return null;
      const order = Array.isArray(entry.order) ? entry.order : [];
      const allowed = new Set(defaultCodes.length ? defaultCodes : this.fields.map((f) => f.code));
      const filteredOrder = order.filter((code) => allowed.has(code));
      const widths = entry.widths && typeof entry.widths === 'object' ? entry.widths : null;
      return { order: filteredOrder, widths, savedAt: entry.savedAt || 0 };
    }

    async persistColumnPref(order, widths) {
      if (!Array.isArray(order) || !order.length) return;
      const map = { ...(await this.loadColumnPrefMap()) };
      const key = this.getColumnPrefKey();
      const widthPayload = widths && typeof widths === 'object' && Object.keys(widths).length ? { ...widths } : undefined;
      const entry = {
        order: order.slice(),
        appId: this.appId || '',
        viewKey: this.viewKey || '',
        viewName: this.viewName || '',
        appName: this.appName || '',
        savedAt: Date.now()
      };
      if (widthPayload) entry.widths = widthPayload;
      map[key] = entry;
      const entries = Object.entries(map);
      if (entries.length > MAX_COLUMN_PREF_ENTRIES) {
        entries.sort((a, b) => {
          const aTime = a[1]?.savedAt || 0;
          const bTime = b[1]?.savedAt || 0;
          return aTime - bTime;
        });
        while (entries.length > MAX_COLUMN_PREF_ENTRIES) {
          const [oldKey] = entries.shift();
          if (oldKey) delete map[oldKey];
        }
      }
      await this.writeColumnPrefMap(map);
    }

    async removeColumnOrder() {
      const map = { ...(await this.loadColumnPrefMap()) };
      const key = this.getColumnPrefKey();
      if (!Object.prototype.hasOwnProperty.call(map, key)) return;
      delete map[key];
      await this.writeColumnPrefMap(map);
    }

    applyColumnOrder(order, options = {}) {
      if (!Array.isArray(order) || !order.length || !this.fields.length) return;
      const updateDraft = options.updateDraft !== false;
      const lookup = new Map(this.fields.map((field) => [field.code, field]));
      const next = [];
      const seen = new Set();
      order.forEach((code) => {
        if (seen.has(code)) return;
        const field = lookup.get(code);
        if (field) {
          next.push(field);
          seen.add(code);
        }
      });
      this.fields.forEach((field) => {
        if (!seen.has(field.code)) {
          next.push(field);
        }
      });
      this.fields = next;
      this.rebuildFieldMaps();
      this.selection = null;
      this.editingCell = null;
      this.pendingFocus = null;
      this.renderGrid();
      this.updateVirtualRows(true);
      this.syncTableWidth();
      if (updateDraft) {
        this.columnOrderDraft = this.fields.map((field) => field.code);
      }
    }

    openColumnManager() {
      if (!this.surface || this.columnPanelLayer || !this.fields.length) return;
      const layer = document.createElement('div');
      layer.className = 'pb-overlay__column-layer';
      layer.addEventListener('click', (event) => {
        if (event.target === layer) this.closeColumnManager();
      });

      const panel = document.createElement('div');
      panel.className = 'pb-overlay__column-panel';
      panel.setAttribute('role', 'dialog');
      panel.setAttribute('aria-label', resolveText(this.language, 'columnsTitle'));
      panel.addEventListener('keydown', (event) => {
        event.stopPropagation();
      });

      const header = document.createElement('div');
      header.className = 'pb-overlay__column-panel-head';
      const titleWrap = document.createElement('div');
      const title = document.createElement('div');
      title.className = 'pb-overlay__column-panel-title';
      title.textContent = resolveText(this.language, 'columnsTitle');
      const hint = document.createElement('p');
      hint.className = 'pb-overlay__column-panel-hint';
      hint.textContent = resolveText(this.language, 'columnsHint');
      titleWrap.appendChild(title);
      titleWrap.appendChild(hint);
      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.className = 'pb-overlay__column-close';
      closeBtn.setAttribute('aria-label', resolveText(this.language, 'columnsClose'));
      closeBtn.textContent = '×';
      closeBtn.addEventListener('click', () => { this.closeColumnManager(); });
      header.appendChild(titleWrap);
      header.appendChild(closeBtn);

      const preview = document.createElement('div');
      preview.className = 'pb-overlay__column-preview';
      const previewScroll = document.createElement('div');
      previewScroll.className = 'pb-overlay__column-preview-scroll';
      const previewRow = document.createElement('div');
      previewRow.className = 'pb-overlay__column-preview-row';
      previewScroll.appendChild(previewRow);
      preview.appendChild(previewScroll);

      const actions = document.createElement('div');
      actions.className = 'pb-overlay__column-panel-actions';
      const autoBtn = document.createElement('button');
      autoBtn.type = 'button';
      autoBtn.className = 'pb-overlay__btn';
      autoBtn.textContent = resolveText(this.language, 'columnsAutoWidth');
      autoBtn.addEventListener('click', () => { this.autoAdjustColumnWidths(); });

      const resetBtn = document.createElement('button');
      resetBtn.type = 'button';
      resetBtn.className = 'pb-overlay__btn';
      resetBtn.textContent = resolveText(this.language, 'columnsReset');
      resetBtn.addEventListener('click', () => { this.handleColumnReset(); });
      const saveBtn = document.createElement('button');
      saveBtn.type = 'button';
      saveBtn.className = 'pb-overlay__btn pb-overlay__btn--primary';
      saveBtn.textContent = resolveText(this.language, 'columnsSave');
      saveBtn.addEventListener('click', () => { this.handleColumnSave(); });
      actions.appendChild(autoBtn);
      actions.appendChild(resetBtn);
      actions.appendChild(saveBtn);

      panel.appendChild(header);
      panel.appendChild(preview);
      panel.appendChild(actions);
      layer.appendChild(panel);
      this.surface.appendChild(layer);

      this.columnOrderDraft = this.fields.map((field) => field.code);
      this.columnPanelLayer = layer;
      this.columnPreviewRow = previewRow;
      previewRow.addEventListener('dragover', (event) => {
        if (!this.draggingColumnCode) return;
        if (event.target.closest('.pb-overlay__column-chip')) return;
        event.preventDefault();
      });
      previewRow.addEventListener('drop', (event) => {
        const chip = event.target.closest('.pb-overlay__column-chip');
        if (chip) return;
        if (!this.draggingColumnCode) return;
        event.preventDefault();
        this.handleColumnDrop(null, { append: true });
      });
      this.renderColumnPreview();

      setTimeout(() => {
        saveBtn.focus();
      }, 0);
    }

    closeColumnManager(skipFocus = false) {
      if (this.columnPanelLayer) {
        this.columnPanelLayer.remove();
      }
      this.columnPanelLayer = null;
      this.columnPreviewRow = null;
      this.draggingColumnCode = null;
      if (!skipFocus && this.columnButton) {
        this.columnButton.focus();
      }
    }

    renderColumnPreview() {
      if (!this.columnPreviewRow) return;
      this.columnPreviewRow.textContent = '';
      if (!this.columnOrderDraft.length) {
        const empty = document.createElement('div');
        empty.className = 'pb-overlay__column-empty';
        empty.textContent = resolveText(this.language, 'columnsEmpty');
        this.columnPreviewRow.appendChild(empty);
        return;
      }
      this.columnOrderDraft.forEach((code, index) => {
        const field = this.fieldMap.get(code);
        const chip = document.createElement('div');
        chip.className = 'pb-overlay__column-chip';
        chip.draggable = true;
        chip.dataset.code = code;
        chip.dataset.width = String(field?.width || DEFAULT_COLUMN_WIDTH);
        chip.style.width = `${field?.width || DEFAULT_COLUMN_WIDTH}px`;
        chip.setAttribute('aria-label', `${columnLabel(index)} ${field?.label || code}`);

        const letter = document.createElement('span');
        letter.className = 'pb-overlay__column-chip-letter';
        letter.textContent = columnLabel(index);

        const label = document.createElement('span');
        label.className = 'pb-overlay__column-chip-label';
        label.textContent = field?.label || code;

        const widthValue = document.createElement('span');
        widthValue.className = 'pb-overlay__column-chip-width';
        widthValue.textContent = `${field?.width || DEFAULT_COLUMN_WIDTH}px`;

        chip.append(letter, label, widthValue);

        chip.addEventListener('dragstart', (event) => {
          this.draggingColumnCode = code;
          chip.classList.add('is-dragging');
          if (event.dataTransfer) {
            event.dataTransfer.effectAllowed = 'move';
            try {
              event.dataTransfer.setData('text/plain', code);
            } catch (_err) {}
          }
        });
        chip.addEventListener('dragend', () => {
          chip.classList.remove('is-dragging');
          chip.classList.remove('is-drop-target');
          this.draggingColumnCode = null;
        });
        chip.addEventListener('dragover', (event) => {
          if (!this.draggingColumnCode || this.draggingColumnCode === code) return;
          event.preventDefault();
          chip.classList.add('is-drop-target');
        });
        chip.addEventListener('dragleave', () => {
          chip.classList.remove('is-drop-target');
        });
        chip.addEventListener('drop', (event) => {
          if (!this.draggingColumnCode) return;
          event.preventDefault();
          chip.classList.remove('is-drop-target');
          this.handleColumnDrop(code);
        });

        const resizer = document.createElement('div');
        resizer.className = 'pb-overlay__column-chip-resizer';
        resizer.setAttribute('role', 'separator');
        resizer.setAttribute('aria-label', `${field?.label || code} width handle`);
        let startX = 0;
        let startWidth = field?.width || DEFAULT_COLUMN_WIDTH;
        const onMouseMove = (event) => {
          const delta = event.clientX - startX;
          const nextWidth = Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, startWidth + delta));
          chip.style.width = `${nextWidth}px`;
          chip.dataset.width = String(nextWidth);
          widthValue.textContent = `${nextWidth}px`;
        };
        const onMouseUp = () => {
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
          const finalWidth = Number(chip.dataset.width || startWidth);
          this.applyColumnWidth(code, finalWidth);
          this.persistColumnPref(this.columnOrderDraft, this.collectColumnWidths()).catch(() => {});
        };
        resizer.addEventListener('mousedown', (event) => {
          event.preventDefault();
          event.stopPropagation();
          startX = event.clientX;
          startWidth = Number(chip.dataset.width || field?.width || DEFAULT_COLUMN_WIDTH);
          document.addEventListener('mousemove', onMouseMove);
          document.addEventListener('mouseup', onMouseUp);
        });

        chip.appendChild(resizer);
        this.columnPreviewRow.appendChild(chip);
      });
    }

    handleColumnDrop(targetCode, options = {}) {
      if (!this.draggingColumnCode) return;
      if (targetCode && targetCode === this.draggingColumnCode) return;
      const from = this.columnOrderDraft.indexOf(this.draggingColumnCode);
      if (from === -1) return;
      const next = [...this.columnOrderDraft];
      const [moved] = next.splice(from, 1);
      if (options.append || !targetCode) {
        next.push(moved);
      } else {
        const to = next.indexOf(targetCode);
        if (to === -1) return;
        next.splice(to, 0, moved);
      }
      this.columnOrderDraft = next;
      this.draggingColumnCode = null;
      this.renderColumnPreview();
    }

    applyColumnWidth(code, width) {
      const clamped = Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.round(width)));
      const field = this.fieldMap.get(code);
      if (field) {
        field.width = clamped;
      }
      const colIndex = this.fields.findIndex((f) => f.code === code);
      if (colIndex >= 0) {
        this.updateColumnWidth(colIndex, clamped);
      }
      if (this.columnPreviewRow) {
        this.renderColumnPreview();
      }
    }

    updateColumnWidth(columnIndex, width) {
      const headers = this.columnHeaderScroll?.querySelectorAll('.pb-overlay__col-header');
      if (headers && headers[columnIndex]) {
        headers[columnIndex].style.width = `${width}px`;
        headers[columnIndex].style.minWidth = `${width}px`;
      }
      const rows = this.tableContainer?.querySelectorAll('.pb-overlay__row');
      if (rows) {
        rows.forEach((row) => {
          const cell = row.children[columnIndex];
          if (cell) {
            cell.style.width = `${width}px`;
            cell.style.minWidth = `${width}px`;
          }
        });
      }
      this.syncTableWidth();
    }

    syncTableWidth() {
      if (!this.tableContainer) return;
      const totalWidth = this.fields.reduce((sum, field) => sum + (field.width || DEFAULT_COLUMN_WIDTH), 0);
      this.tableContainer.style.width = `${totalWidth}px`;
    }

    collectColumnWidths() {
      const map = {};
      this.fields.forEach((field) => {
        if (field.width && field.width !== DEFAULT_COLUMN_WIDTH) {
          map[field.code] = field.width;
        }
      });
      return map;
    }

    autoAdjustColumnWidths() {
      if (!this.fields.length) return;
      const measureTextWidth = (text) => {
        if (!text) return 0;
        const str = String(text);
        let width = 0;
        for (const ch of str) {
          width += ch.charCodeAt(0) > 255 ? 16 : 9;
        }
        return Math.round(width * 1.15 + 32);
      };
      const sampleRows = this.rows.slice(0, 60);
      const widths = {};
      this.fields.forEach((field) => {
        const headerWidth = measureTextWidth(field.label || field.code);
        let maxWidth = headerWidth;
        sampleRows.forEach((row) => {
          const value = row?.values?.[field.code] ?? '';
          maxWidth = Math.max(maxWidth, measureTextWidth(value));
        });
        widths[field.code] = Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.round(maxWidth)));
      });
      Object.entries(widths).forEach(([code, width]) => {
        this.applyColumnWidth(code, width);
      });
      if (this.columnPreviewRow) {
        this.renderColumnPreview();
      }
      this.persistColumnPref(this.columnOrderDraft, this.collectColumnWidths()).catch(() => {});
      this.notify(resolveText(this.language, 'toastColumnsAutoWidth'));
    }

    async handleColumnSave() {
      if (!this.columnOrderDraft.length) {
        this.closeColumnManager();
        return;
      }
      try {
        const widths = this.collectColumnWidths();
        await this.persistColumnPref(this.columnOrderDraft, widths);
        this.applyColumnOrder(this.columnOrderDraft);
        this.notify(resolveText(this.language, 'toastColumnsSaved'));
      } catch (error) {
        console.error('[kintone-excel-overlay] failed to save column order', error);
        this.notify(resolveText(this.language, 'toastSaveFailed'));
      }
      this.closeColumnManager(true);
      if (this.columnButton) this.columnButton.focus();
    }

    async handleColumnReset() {
      try {
        await this.removeColumnOrder();
        if (this.defaultFieldCodes.length) {
          this.applyColumnOrder(this.defaultFieldCodes, { updateDraft: false });
          this.fields.forEach((field) => {
            field.width = DEFAULT_COLUMN_WIDTH;
          });
          this.renderGrid();
          this.syncTableWidth();
        }
        this.columnOrderDraft = this.fields.map((field) => field.code);
        this.renderColumnPreview();
        this.notify(resolveText(this.language, 'toastColumnsReset'));
      } catch (error) {
        console.error('[kintone-excel-overlay] failed to reset column order', error);
        this.notify(resolveText(this.language, 'toastSaveFailed'));
      }
    }

    transformRecord(record) {
      const id = record?.$id?.value ? String(record.$id.value) : '';
      const revision = record?.$revision?.value ? String(record.$revision.value) : undefined;
      const values = {};
      const original = {};
      this.fields.forEach((field) => {
        const raw = record?.[field.code]?.value;
        const baseValue = raw === undefined || raw === null ? '' : String(raw);
        values[field.code] = baseValue;
        original[field.code] = baseValue;
      });
      return { id, revision, values, original };
    }

    renderNoFields() {
      if (!this.surface) return;
      this.surface.innerHTML = '';
      const message = document.createElement('div');
      message.className = 'pb-overlay__message';
      message.textContent = resolveText(this.language, 'noEditableFields');
      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.className = 'pb-overlay__btn pb-overlay__btn--primary';
      closeBtn.textContent = resolveText(this.language, 'btnClose');
      closeBtn.addEventListener('click', () => { this.close(true); });
      message.appendChild(closeBtn);
      this.surface.appendChild(message);
    }

    renderGrid() {
      if (!this.surface) return;
      this.closeRadioPicker();
      this.inputsByRow.clear();
      const headRow = this.columnHeaderScroll.querySelector('.pb-overlay__col-row');
      headRow.textContent = '';
      const rowWrap = this.rowHeaderScroll.querySelector('.pb-overlay__row-wrap');
      const table = this.bodyScroll.querySelector('.pb-overlay__table');
      rowWrap.textContent = '';
      table.textContent = '';

      const totalHeight = Math.max(0, this.rows.length * this.rowHeight);
      rowWrap.style.height = `${totalHeight}px`;
      table.style.height = `${totalHeight}px`;

      this.fields.forEach((field, index) => {
        const headerCell = document.createElement('div');
        headerCell.className = 'pb-overlay__col-header';
        headerCell.style.width = `${field.width || DEFAULT_COLUMN_WIDTH}px`;
        headerCell.style.minWidth = `${field.width || DEFAULT_COLUMN_WIDTH}px`;
        const codeSpan = document.createElement('span');
        codeSpan.className = 'pb-overlay__col-code';
        codeSpan.textContent = columnLabel(index);
        const labelSpan = document.createElement('span');
        labelSpan.className = 'pb-overlay__col-label';
        labelSpan.textContent = field.label || field.code;
        headerCell.appendChild(codeSpan);
        headerCell.appendChild(labelSpan);
        headRow.appendChild(headerCell);
      });

      const rowContainer = document.createElement('div');
      rowContainer.className = 'pb-overlay__row-virtual';
      rowContainer.style.position = 'absolute';
      rowContainer.style.top = '0';
      rowContainer.style.left = '0';
      rowContainer.style.right = '0';
      rowContainer.style.width = '100%';
      this.rowHeaderContainer = rowContainer;
      rowWrap.appendChild(rowContainer);

      const tableContainer = document.createElement('div');
      tableContainer.className = 'pb-overlay__table-virtual';
      tableContainer.style.position = 'absolute';
      tableContainer.style.top = '0';
      tableContainer.style.left = '0';
      const totalWidth = this.fields.reduce((sum, field) => sum + (field.width || DEFAULT_COLUMN_WIDTH), 0);
      tableContainer.style.width = `${totalWidth}px`;
      this.tableContainer = tableContainer;
      table.appendChild(tableContainer);

      this.virtualStart = 0;
      this.virtualEnd = 0;
      this.updateVirtualRows(true);
    }

    updateVirtualRows(force = false) {
      if (!this.bodyScroll || !this.tableContainer || !this.rowHeaderContainer) return;
      const totalRows = this.rows.length;
      const viewportHeight = this.bodyScroll.clientHeight || (10 * this.rowHeight);
      const scrollTop = this.bodyScroll.scrollTop;
      const buffer = 6;
      const start = Math.max(0, Math.floor(scrollTop / this.rowHeight) - buffer);
      const end = Math.min(totalRows, Math.ceil((scrollTop + viewportHeight) / this.rowHeight) + buffer);
      if (!force && start === this.virtualStart && end === this.virtualEnd) {
        this.restorePendingFocus();
        return;
      }
      this.virtualStart = start;
      this.virtualEnd = end;
      this.inputsByRow.clear();
      this.rowHeaderContainer.textContent = '';
      this.rowHeaderContainer.style.transform = `translateY(${-scrollTop}px)`;
      this.tableContainer.textContent = '';
      const totalWidth = this.fields.length * 160;
      this.tableContainer.style.width = `${totalWidth}px`;

      const headerFrag = document.createDocumentFragment();
      const tableFrag = document.createDocumentFragment();

      for (let rowIndex = start; rowIndex < end; rowIndex += 1) {
        const row = this.rows[rowIndex];
        if (!row) continue;
        const top = rowIndex * this.rowHeight;

        const header = document.createElement('div');
        header.className = 'pb-overlay__row-header';
        header.style.position = 'absolute';
        header.style.top = `${top}px`;
        header.style.left = '0';
        header.textContent = String(rowIndex + 1);
        headerFrag.appendChild(header);

        const rowLine = document.createElement('div');
        rowLine.className = 'pb-overlay__row';
        rowLine.style.position = 'absolute';
        rowLine.style.top = `${top}px`;
        rowLine.style.left = '0';
        rowLine.style.width = 'max-content';

        const inputsRow = [];

        this.fields.forEach((field, colIndex) => {
          const cell = document.createElement('div');
          cell.className = 'pb-overlay__cell';
          const fieldWidth = field.width || DEFAULT_COLUMN_WIDTH;
          cell.style.width = `${fieldWidth}px`;
          cell.style.minWidth = `${fieldWidth}px`;
          const input = this.createInputBase(field);
          this.bindInputToRow(input, field, row, rowIndex, colIndex);
          cell.appendChild(input);
          this.applyCellVisualState(input, row, field.code);
          rowLine.appendChild(cell);
          inputsRow[colIndex] = input;
        });

        tableFrag.appendChild(rowLine);
        this.inputsByRow.set(rowIndex, inputsRow);
      }

      this.rowHeaderContainer.appendChild(headerFrag);
      this.tableContainer.appendChild(tableFrag);
      this.repaintSelection();
      this.restorePendingFocus();
    }

    createInputBase(field) {
      const input = document.createElement('input');
      input.className = 'pb-overlay__input';
      input.spellcheck = false;
      if (field.type === 'NUMBER') input.inputMode = 'decimal';
      if (field.type === 'DATE') input.type = 'date';
      else input.type = 'text';
      if (field.type === 'RADIO_BUTTON') {
        input.placeholder = field.choices?.[0] || '';
      }
      input.readOnly = true;
      input.addEventListener('input', () => { this.onInputChanged(input); });
      input.addEventListener('change', () => { this.onInputChanged(input); });
      input.addEventListener('compositionstart', this.handleCompositionStart);
      input.addEventListener('compositionend', this.handleCompositionEnd);
      input.addEventListener('paste', (event) => { this.handlePaste(event, input); });
      input.addEventListener('mousedown', (event) => { this.handleCellMouseDown(event, input); });
      input.addEventListener('mouseenter', () => { this.handleCellHover(input); });
      input.addEventListener('blur', (event) => {
        this.handleInputBlur(event, input);
      });
      input.addEventListener('click', () => {
        this.handleRadioInputClick(input);
      });
      return input;
    }

    bindInputToRow(input, field, row, rowIndex, colIndex) {
      input.dataset.recordId = row.id;
      input.dataset.fieldCode = field.code;
      input.dataset.rowIndex = String(rowIndex);
      input.dataset.colIndex = String(colIndex);
      input.dataset.originalValue = row.original[field.code];
      input.disabled = false;
      const editing = this.isEditingCell(rowIndex, colIndex);
      input.readOnly = !editing;
      if (editing) {
        input.classList.add('pb-overlay__input--editing');
      } else {
        input.classList.remove('pb-overlay__input--editing');
      }
    }

    applyCellVisualState(input, row, fieldCode) {
      const value = row?.values ? row.values[fieldCode] ?? '' : '';
      input.value = value;
      const cell = input.parentElement;
      if (!cell) return;
      const diffEntry = this.diff.get(row.id);
      if (diffEntry && diffEntry[fieldCode]) {
        cell.classList.add('pb-overlay__cell--dirty');
      } else {
        cell.classList.remove('pb-overlay__cell--dirty');
      }
      const key = `${row.id}:${fieldCode}`;
      if (this.invalidCells.has(key)) {
        cell.classList.add('pb-overlay__cell--error');
        if (this.invalidValueCache.has(key)) {
          input.value = this.invalidValueCache.get(key) ?? '';
        }
      } else {
        cell.classList.remove('pb-overlay__cell--error');
      }
      const rowIndex = Number(input.dataset.rowIndex || '-1');
      const colIndex = Number(input.dataset.colIndex || '-1');
      if (this.isCellSelected(rowIndex, colIndex)) {
        cell.classList.add('pb-overlay__cell--selected');
      } else {
        cell.classList.remove('pb-overlay__cell--selected');
      }
      const editing = this.isEditingCell(rowIndex, colIndex);
      input.readOnly = !editing;
      if (editing) {
        input.classList.add('pb-overlay__input--editing');
      } else {
        input.classList.remove('pb-overlay__input--editing');
      }
    }

    getInput(rowIndex, colIndex) {
      const rowInputs = this.inputsByRow.get(rowIndex);
      if (!rowInputs) return null;
      return rowInputs[colIndex] || null;
    }

    focusCell(rowIndex, colIndex) {
      const maxRow = Math.max(0, this.rows.length - 1);
      const maxCol = Math.max(0, this.fields.length - 1);
      if (maxRow < 0 || maxCol < 0) return;
      const r = Math.max(0, Math.min(maxRow, rowIndex));
      const c = Math.max(0, Math.min(maxCol, colIndex));
      this.pendingFocus = { rowIndex: r, colIndex: c };
      if (r < this.virtualStart || r >= this.virtualEnd) {
        if (this.bodyScroll) {
          this.bodyScroll.scrollTop = r * this.rowHeight;
          this.updateVirtualRows(true);
        }
        return;
      }
      this.restorePendingFocus();
    }

    restorePendingFocus() {
      if (!this.pendingFocus) return;
      const { rowIndex, colIndex } = this.pendingFocus;
      const input = this.getInput(rowIndex, colIndex);
      if (!input) return;
      this.pendingFocus = null;
      requestAnimationFrame(() => {
        input.focus();
        const editing = this.editingCell && this.editingCell.rowIndex === rowIndex && this.editingCell.colIndex === colIndex;
        const shouldSelectAll = editing && this.editingCell.selectAll;
        this.applyCaretPlacement(input, shouldSelectAll);
      });
    }

    canSelectTextInput(input) {
      if (!input) return false;
      const type = (input.type || '').toLowerCase();
      return type === 'text' || type === 'search' || type === 'tel' || type === 'url' || type === 'email' || type === 'password';
    }

    handleInputBlur(event, input) {
      if (!this.editingCell) return;
      if (this.radioPicker && this.radioPicker.input === input) {
        const next = event.relatedTarget;
        if (next && this.radioPicker.panel?.contains(next)) {
          return;
        }
      }
      this.exitEditMode();
    }

    applyCaretPlacement(input, selectAll) {
      if (!this.canSelectTextInput(input)) return;
      try {
        if (selectAll) {
          input.select();
        } else {
          const len = input.value.length;
          input.setSelectionRange(len, len);
        }
      } catch (_e) {
        // Some browsers throw for inputs that disallow selection even if type reports text.
      }
    }

    openRadioPicker(rowIndex, colIndex, input) {
      if (!this.root || !input) return;
      const field = this.fields[colIndex];
      if (!field || field.type !== 'RADIO_BUTTON') return;
      const rawChoices = Array.isArray(field.choices) ? field.choices : [];
      const choices = Array.from(new Set(rawChoices.map((choice) => String(choice || '')).filter((choice) => choice)));
      if (!choices.length) return;

      this.closeRadioPicker();
      const panel = document.createElement('div');
      panel.className = 'pb-overlay__choice-panel';
      panel.setAttribute('role', 'listbox');
      panel.setAttribute('aria-label', field.label || field.code || 'radio choices');
      const list = document.createElement('div');
      list.className = 'pb-overlay__choice-list';
      const currentValue = input.value || '';
      choices.forEach((choice) => {
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'pb-overlay__choice-item';
        button.dataset.value = choice;
        button.textContent = choice;
        button.setAttribute('role', 'option');
        if (choice === currentValue) {
          button.classList.add('is-selected');
          button.setAttribute('aria-selected', 'true');
        } else {
          button.setAttribute('aria-selected', 'false');
        }
        button.addEventListener('click', () => {
          this.handleRadioChoiceSelect(choice, rowIndex, colIndex);
        });
        list.appendChild(button);
      });
      panel.style.visibility = 'hidden';
      panel.style.left = '0px';
      panel.style.top = '0px';
      panel.appendChild(list);
      this.root.appendChild(panel);
      this.positionRadioPicker(panel, input);
      panel.style.visibility = 'visible';
      const focusTarget = panel.querySelector('.pb-overlay__choice-item.is-selected') || panel.querySelector('.pb-overlay__choice-item');
      if (focusTarget) {
        requestAnimationFrame(() => {
          try {
            focusTarget.focus();
          } catch (_e) { /* noop */ }
        });
      }
      this.radioPicker = { panel, input, rowIndex, colIndex };
      if (this.bodyScroll) {
        this.bodyScroll.addEventListener('scroll', this.boundRadioPickerScrollHandler, { passive: true });
      }
      window.addEventListener('resize', this.boundRadioPickerResizeHandler);
      document.addEventListener('mousedown', this.boundRadioPickerOutsideClick, true);
    }

    positionRadioPicker(panel, anchor) {
      if (!this.root) return;
      const rootRect = this.root.getBoundingClientRect();
      const anchorRect = anchor.getBoundingClientRect();
      const margin = 12;
      const anchorWidth = anchorRect.width || panel.getBoundingClientRect().width;
      panel.style.minWidth = `${Math.max(180, anchorWidth)}px`;
      const panelRect = panel.getBoundingClientRect();
      const maxLeft = Math.max(margin, rootRect.width - panelRect.width - margin);
      let left = anchorRect.left - rootRect.left;
      left = Math.max(margin, Math.min(left, maxLeft));
      let top = anchorRect.bottom - rootRect.top + 6;
      const maxTop = Math.max(margin, rootRect.height - panelRect.height - margin);
      if (top > maxTop) {
        top = anchorRect.top - rootRect.top - panelRect.height - 6;
      }
      if (top < margin) top = margin;
      panel.style.left = `${left}px`;
      panel.style.top = `${top}px`;
    }

    closeRadioPicker(focusInput = false) {
      if (!this.radioPicker) return;
      const { panel, input } = this.radioPicker;
      if (panel?.parentElement) {
        panel.remove();
      }
      this.radioPicker = null;
      if (this.bodyScroll) {
        this.bodyScroll.removeEventListener('scroll', this.boundRadioPickerScrollHandler);
      }
      window.removeEventListener('resize', this.boundRadioPickerResizeHandler);
      document.removeEventListener('mousedown', this.boundRadioPickerOutsideClick, true);
      if (focusInput && input) {
        try {
          input.focus();
        } catch (_e) { /* noop */ }
      }
    }

    moveRadioPickerFocus(delta, currentTarget) {
      if (!this.radioPicker?.panel) return;
      const buttons = Array.from(this.radioPicker.panel.querySelectorAll('.pb-overlay__choice-item'));
      if (!buttons.length) return;
      let index = -1;
      if (currentTarget instanceof HTMLButtonElement) {
        index = buttons.indexOf(currentTarget);
      }
      if (index === -1) {
        const active = document.activeElement;
        if (active instanceof HTMLButtonElement) {
          index = buttons.indexOf(active);
        }
      }
      if (index === -1) index = 0;
      let nextIndex = index + delta;
      if (nextIndex < 0) nextIndex = buttons.length - 1;
      if (nextIndex >= buttons.length) nextIndex = 0;
      buttons[nextIndex].focus();
    }

    repositionRadioPicker() {
      if (!this.radioPicker?.panel || !this.radioPicker.input) return;
      if (!this.root || !this.root.contains(this.radioPicker.input)) {
        this.closeRadioPicker();
        return;
      }
      this.positionRadioPicker(this.radioPicker.panel, this.radioPicker.input);
    }

    handleRadioChoiceSelect(value, rowIndex, colIndex) {
      const input = this.getInput(rowIndex, colIndex);
      if (!input) {
        this.closeRadioPicker();
        return;
      }
      input.value = value;
      this.onInputChanged(input);
      this.closeRadioPicker(true);
    }

    handleCellMouseDown(event, input) {
      if (event.button !== 0) return;
      const rowIndex = Number(input.dataset.rowIndex || '0');
      const colIndex = Number(input.dataset.colIndex || '0');
      const editing = this.isEditingCell(rowIndex, colIndex);
      const isDouble = event.detail >= 2;

      if (editing) {
        event.stopPropagation();
        if (this.editingCell) {
          this.editingCell.selectAll = false;
        }
        this.dragSelecting = false;
        return;
      }

      event.preventDefault();
      event.stopPropagation();
      this.startSelection(rowIndex, colIndex, event.shiftKey);
      if (isDouble) {
        this.enterEditMode(rowIndex, colIndex, input, true);
      }
    }

    handleCellHover(input) {
      if (!this.dragSelecting || this.editingCell) return;
      const rowIndex = Number(input.dataset.rowIndex || '0');
      const colIndex = Number(input.dataset.colIndex || '0');
      this.updateSelectionRange(rowIndex, colIndex);
    }

    handleRadioInputClick(input) {
      if (!input) return;
      const fieldCode = input.dataset.fieldCode;
      if (!fieldCode) return;
      const field = this.fieldMap.get(fieldCode);
      if (!field || field.type !== 'RADIO_BUTTON') return;
      const rowIndex = Number(input.dataset.rowIndex || '0');
      const colIndex = Number(input.dataset.colIndex || '0');
      if (!this.isEditingCell(rowIndex, colIndex)) return;
      if (this.radioPicker && this.radioPicker.input === input) return;
      this.openRadioPicker(rowIndex, colIndex, input);
    }

    onInputChanged(input) {
      const recordId = input.dataset.recordId;
      const fieldCode = input.dataset.fieldCode;
      if (!recordId || !fieldCode) return;
      const key = `${recordId}:${fieldCode}`;
      const field = this.fieldMap.get(fieldCode);
      if (!field) return;
      const cell = input.parentElement;
      if (!cell) return;

      const rawValue = input.value;
      const validation = this.validate(rawValue, field);

      if (!validation.ok) {
        cell.classList.add('pb-overlay__cell--error');
        this.invalidCells.add(key);
        this.invalidValueCache.set(key, rawValue);
      } else {
        cell.classList.remove('pb-overlay__cell--error');
        this.invalidCells.delete(key);
        this.invalidValueCache.delete(key);
      }

      if (!validation.ok) {
        this.markDirtyState(recordId, fieldCode, input.dataset.originalValue, null);
        this.updateStats();
        return;
      }

      const normalized = validation.value;
      const originalValue = input.dataset.originalValue;

      const row = this.rowMap.get(recordId);
      if (row) {
        row.values[fieldCode] = normalized;
      }

      if (normalized === originalValue) {
        cell.classList.remove('pb-overlay__cell--dirty');
        this.removeDiff(recordId, fieldCode);
      } else {
        cell.classList.add('pb-overlay__cell--dirty');
        this.addDiff(recordId, fieldCode, normalized, field.type);
      }
      this.updateDirtyBadge();
      this.updateStats();
    }

    handlePaste(event, input) {
      if (!event.clipboardData) return;
      const text = event.clipboardData.getData('text');
      if (!text) return;
      const normalized = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      if (!normalized.includes('\n') && !normalized.includes('\t')) {
        return;
      }
      event.preventDefault();
      this.applyPastedMatrix(normalized, input);
    }

    applyPastedMatrix(text, startInput) {
      const lines = text.split('\n');
      if (lines.length && lines[lines.length - 1] === '') {
        lines.pop();
      }
      if (!lines.length) return;
      const matrix = lines.map((line) => line.split('\t'));
      const startRow = Number(startInput.dataset.rowIndex || '0');
      const startCol = Number(startInput.dataset.colIndex || '0');
      const rowCount = this.rows.length;
      const colCount = this.fields.length;
      for (let r = 0; r < matrix.length; r += 1) {
        const targetRowIndex = startRow + r;
        if (targetRowIndex >= rowCount) break;
        const rowValues = matrix[r];
        for (let c = 0; c < rowValues.length; c += 1) {
          const targetColIndex = startCol + c;
          if (targetColIndex >= colCount) break;
          this.applyPastedValue(targetRowIndex, targetColIndex, rowValues[c]);
        }
      }
    }

    applyPastedValue(rowIndex, colIndex, rawValue) {
      const row = this.rows[rowIndex];
      const field = this.fields[colIndex];
      if (!row || !field) return;
      const input = this.getInput(rowIndex, colIndex);
      if (input) {
        input.value = rawValue;
        this.onInputChanged(input);
        return;
      }
      const recordId = row.id;
      const fieldCode = field.code;
      const key = `${recordId}:${fieldCode}`;
      const validation = this.validate(rawValue, field);
      if (!validation.ok) {
        this.invalidCells.add(key);
        this.invalidValueCache.set(key, rawValue);
        this.markDirtyState(recordId, fieldCode, row.original[fieldCode], null);
        this.updateDirtyBadge();
        this.updateStats();
        return;
      }
      this.invalidCells.delete(key);
      this.invalidValueCache.delete(key);
      const normalized = validation.value;
      row.values[fieldCode] = normalized;
      if (normalized === row.original[fieldCode]) {
        this.removeDiff(recordId, fieldCode);
      } else {
        this.addDiff(recordId, fieldCode, normalized, field.type);
      }
      this.updateDirtyBadge();
      this.updateStats();
    }

    setSelectionSingle(rowIndex, colIndex) {
      const maxRow = Math.max(0, this.rows.length - 1);
      const maxCol = Math.max(0, this.fields.length - 1);
      if (maxRow < 0 || maxCol < 0) {
        this.selection = null;
        this.repaintSelection();
        this.updateStats();
        return;
      }
      const r = Math.max(0, Math.min(maxRow, rowIndex));
      const c = Math.max(0, Math.min(maxCol, colIndex));
      this.selection = {
        anchorRow: r,
        anchorCol: c,
        startRow: r,
        endRow: r,
        startCol: c,
        endCol: c
      };
      this.dragSelecting = false;
      this.repaintSelection();
      this.updateStats();
    }

    startSelection(rowIndex, colIndex, extend = false) {
      this.exitEditMode();
      const maxRow = Math.max(0, this.rows.length - 1);
      const maxCol = Math.max(0, this.fields.length - 1);
      if (maxRow < 0 || maxCol < 0) return;
      const r = Math.max(0, Math.min(maxRow, rowIndex));
      const c = Math.max(0, Math.min(maxCol, colIndex));
      if (!extend || !this.selection) {
        this.selection = {
          anchorRow: r,
          anchorCol: c,
          startRow: r,
          endRow: r,
          startCol: c,
          endCol: c
        };
      } else {
        if (typeof this.selection.anchorRow !== 'number') {
          this.selection.anchorRow = this.selection.startRow;
          this.selection.anchorCol = this.selection.startCol;
        }
        this.selection.startRow = Math.min(this.selection.anchorRow, r);
        this.selection.endRow = Math.max(this.selection.anchorRow, r);
        this.selection.startCol = Math.min(this.selection.anchorCol, c);
        this.selection.endCol = Math.max(this.selection.anchorCol, c);
      }
      this.dragSelecting = true;
      this.repaintSelection();
      this.updateStats();
      this.focusCell(r, c);
    }

    updateSelectionRange(rowIndex, colIndex, focus = false) {
      const maxRow = Math.max(0, this.rows.length - 1);
      const maxCol = Math.max(0, this.fields.length - 1);
      if (maxRow < 0 || maxCol < 0) return;
      const r = Math.max(0, Math.min(maxRow, rowIndex));
      const c = Math.max(0, Math.min(maxCol, colIndex));
      if (!this.selection) {
        this.selection = {
          anchorRow: r,
          anchorCol: c,
          startRow: r,
          endRow: r,
          startCol: c,
          endCol: c
        };
      } else {
        const anchorRow = typeof this.selection.anchorRow === 'number' ? this.selection.anchorRow : r;
        const anchorCol = typeof this.selection.anchorCol === 'number' ? this.selection.anchorCol : c;
        this.selection.startRow = Math.min(anchorRow, r);
        this.selection.endRow = Math.max(anchorRow, r);
        this.selection.startCol = Math.min(anchorCol, c);
        this.selection.endCol = Math.max(anchorCol, c);
      }
      this.repaintSelection();
      this.updateStats();
      if (focus) {
        this.focusCell(r, c);
      }
    }

    endSelection() {
      this.dragSelecting = false;
    }

    isCellSelected(rowIndex, colIndex) {
      if (!this.selection) return false;
      if (rowIndex < 0 || colIndex < 0) return false;
      const s = this.selection;
      return rowIndex >= s.startRow && rowIndex <= s.endRow
        && colIndex >= s.startCol && colIndex <= s.endCol;
    }

    repaintSelection() {
      this.inputsByRow.forEach((rowInputs, rowIndex) => {
        if (!Array.isArray(rowInputs)) return;
        rowInputs.forEach((input, colIndex) => {
          if (!input) return;
          const cell = input.parentElement;
          if (!cell) return;
          if (this.isCellSelected(Number(rowIndex), Number(colIndex))) {
            cell.classList.add('pb-overlay__cell--selected');
          } else {
            cell.classList.remove('pb-overlay__cell--selected');
          }
        });
      });
    }

    isEditingCell(rowIndex, colIndex) {
      if (!this.editingCell) return false;
      return this.editingCell.rowIndex === rowIndex && this.editingCell.colIndex === colIndex;
    }

    enterEditMode(rowIndex, colIndex, input, selectAll = false) {
      const maxRow = Math.max(0, this.rows.length - 1);
      const maxCol = Math.max(0, this.fields.length - 1);
      if (maxRow < 0 || maxCol < 0) return;
      const r = Math.max(0, Math.min(maxRow, rowIndex));
      const c = Math.max(0, Math.min(maxCol, colIndex));
      const field = this.fields[c];
      if (this.editingCell && (this.editingCell.rowIndex !== r || this.editingCell.colIndex !== c)) {
        this.exitEditMode();
      }
      this.dragSelecting = false;
      this.setSelectionSingle(r, c);
      this.editingCell = { rowIndex: r, colIndex: c, selectAll: !!selectAll };
      const targetInput = input || this.getInput(r, c);
      if (targetInput) {
        targetInput.readOnly = false;
        targetInput.classList.add('pb-overlay__input--editing');
        requestAnimationFrame(() => {
          targetInput.focus();
          if (this.editingCell && this.editingCell.rowIndex === r && this.editingCell.colIndex === c) {
            this.applyCaretPlacement(targetInput, !!this.editingCell.selectAll);
          }
          if (field?.type === 'RADIO_BUTTON') {
            this.openRadioPicker(r, c, targetInput);
          }
        });
      }
    }

    exitEditMode() {
      if (!this.editingCell) return;
      const { rowIndex, colIndex } = this.editingCell;
      this.closeRadioPicker();
      const input = this.getInput(rowIndex, colIndex);
      if (input) {
        input.readOnly = true;
        input.classList.remove('pb-overlay__input--editing');
      }
      this.editingCell = null;
      this.setSelectionSingle(rowIndex, colIndex);
    }

    async copySelectionToClipboard() {
      this.exitEditMode();
      let sel = this.selection;
      if (!sel) {
        const active = this.pendingFocus || (() => {
          const el = document.activeElement;
          if (el instanceof HTMLInputElement && el.dataset) {
            return {
              rowIndex: Number(el.dataset.rowIndex || '0'),
              colIndex: Number(el.dataset.colIndex || '0')
            };
          }
          return { rowIndex: 0, colIndex: 0 };
        })();
        this.setSelectionSingle(active.rowIndex, active.colIndex);
        sel = this.selection;
      }
      if (!sel) return;
      const rowsOut = [];
      for (let r = sel.startRow; r <= sel.endRow; r += 1) {
        const rowVals = [];
        const rowData = this.rows[r];
        for (let c = sel.startCol; c <= sel.endCol; c += 1) {
          const field = this.fields[c];
          const value = rowData && field ? rowData.values[field.code] : '';
          rowVals.push(value === undefined || value === null ? '' : String(value));
        }
        rowsOut.push(rowVals.join('\t'));
      }
      const text = rowsOut.join('\n');
      try {
        if (typeof navigator !== 'undefined' && navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(text);
        } else {
          const textarea = document.createElement('textarea');
          textarea.value = text;
          textarea.setAttribute('readonly', 'true');
          textarea.style.position = 'fixed';
          textarea.style.top = '-9999px';
          document.body.appendChild(textarea);
          textarea.focus();
          textarea.select();
          const ok = document.execCommand('copy');
          textarea.remove();
          if (!ok) throw new Error('execCommand copy returned false');
        }
        this.notify(resolveText(this.language, 'toastCopySuccess'));
      } catch (error) {
        this.notify(resolveText(this.language, 'toastCopyFailed'));
        console.error('[kintone-excel-overlay] copy failed', error);
      }
    }

    addDiff(recordId, fieldCode, value, type) {
      let entry = this.diff.get(recordId);
      if (!entry) {
        entry = {};
        this.diff.set(recordId, entry);
      }
      entry[fieldCode] = { value, type };
    }

    removeDiff(recordId, fieldCode) {
      const entry = this.diff.get(recordId);
      if (!entry) return;
      delete entry[fieldCode];
      if (Object.keys(entry).length === 0) {
        this.diff.delete(recordId);
      }
    }

    markDirtyState(recordId, fieldCode, originalValue, value) {
      if (value === originalValue) {
        this.removeDiff(recordId, fieldCode);
        const input = this.findInput(recordId, fieldCode);
        if (input) input.parentElement.classList.remove('pb-overlay__cell--dirty');
      }
      this.updateDirtyBadge();
    }

    findInput(recordId, fieldCode) {
      const rowIndex = this.rowIndexMap.get(recordId);
      if (rowIndex === undefined) return null;
      const colIndex = this.fieldIndexMap.get(fieldCode);
      if (colIndex === undefined) return null;
      return this.getInput(rowIndex, colIndex);
    }

    updateDirtyBadge() {
      if (!this.dirtyBadge) return;
      let count = 0;
      this.diff.forEach((entry) => {
        count += Object.keys(entry).length;
      });
      this.dirtyBadge.textContent = resolveText(this.language, 'dirtyLabel', count);
      if (count > 0) {
        this.dirtyBadge.classList.add('pb-overlay__dirty--active');
      } else {
        this.dirtyBadge.classList.remove('pb-overlay__dirty--active');
      }
    }

    updateStats() {
      const numbers = [];
      if (this.selection) {
        for (let r = this.selection.startRow; r <= this.selection.endRow; r += 1) {
          const row = this.rows[r];
          if (!row) continue;
          for (let c = this.selection.startCol; c <= this.selection.endCol; c += 1) {
            const field = this.fields[c];
            if (!field || field.type !== 'NUMBER') continue;
            const raw = row.values[field.code];
            if (raw === undefined || raw === null || raw === '') continue;
            const parsed = Number(String(raw).replace(/,/g, ''));
            if (!Number.isNaN(parsed)) numbers.push(parsed);
          }
        }
      } else {
        this.diff.forEach((entry) => {
          Object.keys(entry).forEach((fieldCode) => {
            const item = entry[fieldCode];
            if (!item || item.type !== 'NUMBER') return;
            const parsed = Number(String(item.value).replace(/,/g, ''));
            if (!Number.isNaN(parsed)) numbers.push(parsed);
          });
        });
      }

      const count = numbers.length;
      const sum = numbers.reduce((acc, v) => acc + v, 0);
      const avg = count ? (sum / count) : 0;
      const sumDisplay = count ? sum : '-';
      const avgDisplay = count ? avg.toFixed(2) : '-';
      const countDisplay = count || '-';

      if (this.statsElements.sum) {
        this.statsElements.sum.textContent = `${resolveText(this.language, 'statusSum')}: ${sumDisplay}`;
      }
      if (this.statsElements.avg) {
        this.statsElements.avg.textContent = `${resolveText(this.language, 'statusAvg')}: ${avgDisplay}`;
      }
      if (this.statsElements.count) {
        this.statsElements.count.textContent = `${resolveText(this.language, 'statusCount')}: ${countDisplay}`;
      }
    }

    validate(value, field) {
      const type = field?.type;
      const trimmed = String(value ?? '').trim();
      if (type === 'NUMBER') {
        if (!trimmed) return { ok: true, value: '' };
        const num = Number(trimmed);
        if (Number.isNaN(num)) return { ok: false, error: 'NaN' };
        return { ok: true, value: trimmed };
      }
      if (type === 'DATE') {
        if (!trimmed) return { ok: true, value: '' };
        const re = /^\d{4}-\d{2}-\d{2}$/;
        if (!re.test(trimmed)) return { ok: false, error: 'Invalid date' };
        return { ok: true, value: trimmed };
      }
      if (type === 'RADIO_BUTTON') {
        if (!trimmed) return { ok: true, value: '' };
        const choices = Array.isArray(field?.choices) ? field.choices : [];
        if (choices.length && !choices.includes(trimmed)) {
          return { ok: false, error: 'Invalid choice' };
        }
        return { ok: true, value: trimmed };
      }
      return { ok: true, value: value ?? '' };
    }

    applyLimit(baseQuery, limit) {
      const trimmed = String(baseQuery || '').trim();
      if (!limit || limit <= 0) return trimmed;
      const regex = /\blimit\s+\d+/i;
      if (regex.test(trimmed)) return trimmed;
      const clause = `limit ${limit}`;
      if (!trimmed) return clause;
      return `${trimmed} ${clause}`;
    }

    syncScroll() {
      if (!this.bodyScroll) return;
      if (this.columnHeaderScroll) {
        this.columnHeaderScroll.scrollLeft = this.bodyScroll.scrollLeft;
      }
      if (this.rowHeaderScroll) {
        this.rowHeaderScroll.scrollTop = this.bodyScroll.scrollTop;
      }
      this.updateVirtualRows();
    }

    focusFirstCell() {
      if (!this.rows.length || !this.fields.length) return;
      this.setSelectionSingle(0, 0);
      this.focusCell(0, 0);
    }

    handleOverlayKeydown(event) {
      if (!this.isOpen) return;
      const target = event.target;
      if (!this.root?.contains(target)) return;

      if (this.columnPanelLayer && this.columnPanelLayer.contains(target)) {
        if (event.key === 'Escape') {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          this.closeColumnManager();
        }
        return;
      }

      if (this.radioPicker?.panel && this.radioPicker.panel.contains(target)) {
        if (event.key === 'Escape') {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          this.closeRadioPicker(true);
        } else if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          const delta = event.key === 'ArrowUp' ? -1 : 1;
          this.moveRadioPickerFocus(delta, target);
        }
        return;
      }

      const targetInput = target instanceof HTMLInputElement ? target : null;
      const fieldCode = targetInput?.dataset.fieldCode || '';
      const field = fieldCode ? this.fieldMap.get(fieldCode) : null;
      const isRadioField = field?.type === 'RADIO_BUTTON';
      const isSpaceKey = event.key === ' ' || event.key === 'Spacebar';

      if (targetInput && targetInput.readOnly && isRadioField) {
        const wantsDirectOpen = !event.ctrlKey && !event.metaKey && !event.altKey && (event.key === 'Enter' || isSpaceKey);
        const wantsAltDropdown = event.key === 'ArrowDown' && event.altKey;
        if (wantsDirectOpen || wantsAltDropdown) {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          const rowIndex = Number(targetInput.dataset.rowIndex || '0');
          const colIndex = Number(targetInput.dataset.colIndex || '0');
          this.enterEditMode(rowIndex, colIndex, targetInput, false);
          return;
        }
      }

      if (targetInput && !targetInput.readOnly) {
        if (isRadioField) {
          const wantsPicker = (isSpaceKey && !event.ctrlKey && !event.metaKey) || (event.key === 'ArrowDown' && event.altKey);
          if (wantsPicker) {
            event.preventDefault();
            event.stopPropagation();
            event.stopImmediatePropagation();
            const rowIndex = Number(targetInput.dataset.rowIndex || '0');
            const colIndex = Number(targetInput.dataset.colIndex || '0');
            this.openRadioPicker(rowIndex, colIndex, targetInput);
            return;
          }
        }
        if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'c') {
          return;
        }
        if (event.key === 'Escape') {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          this.exitEditMode();
          return;
        }
        if (event.key === 'Enter') {
          if (this.isComposing) return;
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          const delta = event.shiftKey ? -1 : 1;
          const rowIndex = Number(targetInput.dataset.rowIndex || '0');
          const colIndex = Number(targetInput.dataset.colIndex || '0');
          this.exitEditMode();
          const coords = this.moveFocus(targetInput, delta, 0);
          if (coords) {
            if (event.shiftKey) {
              this.updateSelectionRange(coords.row, coords.col, true);
            } else {
              this.setSelectionSingle(coords.row, coords.col);
            }
          }
          return;
        }
        return;
      }

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'c') {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        this.exitEditMode();
        this.copySelectionToClipboard();
        return;
      }

      if (event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        this.tryClose();
        return;
      }

      if (!targetInput) {
        if (event.key === 's' && (event.ctrlKey || event.metaKey)) {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          this.save();
        }
        return;
      }

      const handleMoveResult = (coords, extend = false) => {
        if (!coords) return;
        if (extend) {
          if (!this.selection) {
            this.setSelectionSingle(coords.row, coords.col);
          } else {
            this.updateSelectionRange(coords.row, coords.col, true);
          }
        } else {
          this.setSelectionSingle(coords.row, coords.col);
        }
      };

      if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        const delta = event.key === 'ArrowUp' ? -1 : 1;
        this.exitEditMode();
        const coords = this.moveFocus(targetInput, delta, 0);
        handleMoveResult(coords, event.shiftKey);
        return;
      }
      if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        const delta = event.key === 'ArrowLeft' ? -1 : 1;
        this.exitEditMode();
        const coords = this.moveFocus(targetInput, 0, delta);
        handleMoveResult(coords, event.shiftKey);
        return;
      }
      if (event.key === 'Enter' && !event.ctrlKey && !event.metaKey) {
        if (this.isComposing) return;
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        const delta = event.shiftKey ? -1 : 1;
        this.exitEditMode();
        const coords = this.moveFocus(targetInput, delta, 0);
        handleMoveResult(coords, false);
        return;
      }
      if (event.key === 'Tab') {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        const delta = event.shiftKey ? -1 : 1;
        this.exitEditMode();
        const coords = this.moveFocus(targetInput, 0, delta);
        handleMoveResult(coords, false);
        return;
      }
      if (event.key === 's' && (event.ctrlKey || event.metaKey)) {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        this.save();
        return;
      }
    }

    moveFocus(currentInput, rowDelta, colDelta) {
      const rowIndex = Number(currentInput.dataset.rowIndex || '0');
      const colIndex = Number(currentInput.dataset.colIndex || '0');
      let nextRow = rowIndex + rowDelta;
      let nextCol = colIndex + colDelta;

      const rowCount = this.rows.length;
      const colCount = this.fields.length;
      if (!rowCount || !colCount) return null;

      if (rowDelta !== 0 && colDelta === 0) {
        if (nextRow < 0) nextRow = 0;
        if (nextRow >= rowCount) nextRow = rowCount - 1;
      }
      if (colDelta !== 0 && rowDelta === 0) {
        if (nextCol < 0) {
          nextCol = colCount - 1;
          nextRow = Math.max(0, rowIndex - 1);
        } else if (nextCol >= colCount) {
          nextCol = 0;
          nextRow = Math.min(rowCount - 1, rowIndex + 1);
        }
      }

      nextRow = Math.max(0, Math.min(rowCount - 1, nextRow));
      nextCol = Math.max(0, Math.min(colCount - 1, nextCol));
      this.focusCell(nextRow, nextCol);
      return { row: nextRow, col: nextCol };
    }

    hasUnsavedChanges() {
      let has = false;
      this.diff.forEach((entry) => {
        if (Object.keys(entry).length) has = true;
      });
      return has;
    }

    isConflictResponse(response) {
      if (!response) return false;
      const status = Number(response.status || 0);
      if (status === 409 || status === 412) return true;
      const code = String(response.errorCode || '');
      return code === 'GAIA_CO02' || code === 'GAIA_CO03' || code === 'CB_WA01';
    }

    buildEntriesForIds(ids) {
      const entries = [];
      ids.forEach((recordId) => {
        const fields = this.diff.get(recordId);
        if (!fields) return;
        const record = {};
        Object.keys(fields).forEach((fieldCode) => {
          record[fieldCode] = { value: fields[fieldCode].value };
        });
        if (Object.keys(record).length) {
          const revision = this.revisionMap.get(recordId);
          const item = { id: recordId, record };
          if (revision !== undefined) item.revision = revision;
          entries.push(item);
        }
      });
      return entries;
    }

    createBatches() {
      const allEntries = this.buildEntriesForIds(Array.from(this.diff.keys()));
      if (!allEntries.length) return [];
      const batches = [];
      for (let i = 0; i < allEntries.length; i += 50) {
        batches.push(allEntries.slice(i, i + 50));
      }
      return batches;
    }

    async save() {
      if (this.saving) return;
      if (this.invalidCells.size > 0) {
        this.notify(resolveText(this.language, 'toastInvalidCells'));
        return;
      }
      let batches = this.createBatches();
      if (!batches.length) {
        this.notify(resolveText(this.language, 'toastNoChanges'));
        return;
      }
      this.saving = true;
      this.saveButton.disabled = true;
      let anySaved = false;
      try {
        for (let index = 0; index < batches.length;) {
          let batch = batches[index];
          let attempt = 0;
          while (attempt < 2) {
            attempt += 1;
            const response = await this.postFn('EXCEL_PUT_RECORDS', {
              appId: this.appId,
              records: batch
            });
            if (response?.ok) {
              this.applySaveResult(batch, response.result);
              anySaved = true;
              break;
            }
            if (attempt === 1 && this.isConflictResponse(response)) {
              const handling = await this.handleConflict(batch);
              if (handling === 'retry') {
                this.notify(resolveText(this.language, 'conflictRetry'));
                batches = this.createBatches();
                if (!batches.length) {
                  this.updateDirtyBadge();
                  this.updateStats();
                  this.notify(resolveText(this.language, 'conflictNoChanges'));
                  this.saving = false;
                  this.saveButton.disabled = false;
                  return;
                }
                index = 0;
                batch = batches[index];
                attempt = 0;
                continue;
              }
              if (handling === 'cleared') {
                this.updateDirtyBadge();
                this.updateStats();
                this.notify(resolveText(this.language, 'conflictNoChanges'));
                this.saving = false;
                this.saveButton.disabled = false;
                return;
              }
            }
            const err = new Error(response?.error || 'save failed');
            if (this.isConflictResponse(response)) {
              err.conflict = true;
            }
            throw err;
          }
          index += 1;
        }
        if (anySaved) {
          this.notify(resolveText(this.language, 'toastSaveSuccess'));
          this.diff.clear();
          this.updateDirtyBadge();
          this.updateStats();
          this.requireReload = true;
        }
      } catch (error) {
        const key = error?.conflict ? 'conflictFailed' : 'toastSaveFailed';
        this.notify(resolveText(this.language, key));
        console.error('[kintone-excel-overlay] save failed', error);
      } finally {
        this.saving = false;
        this.saveButton.disabled = false;
      }
    }

    applySaveResult(batch, result) {
      const records = result?.records;
      const revisionById = new Map();
      if (Array.isArray(records)) {
        records.forEach((item, index) => {
          const id = item?.id || batch[index]?.id;
          if (id && item?.revision !== undefined) {
            revisionById.set(String(id), String(item.revision));
          }
        });
      }
      batch.forEach((entry) => {
        const row = this.rowMap.get(entry.id);
        if (!row) return;
        Object.keys(entry.record).forEach((fieldCode) => {
          const value = entry.record[fieldCode].value ?? '';
          row.original[fieldCode] = value;
          row.values[fieldCode] = value;
          const input = this.findInput(entry.id, fieldCode);
          if (input) {
            input.dataset.originalValue = value;
            input.value = value;
            this.applyCellVisualState(input, row, fieldCode);
          }
        });
        const newRevision = revisionById.get(entry.id);
        if (newRevision) {
          row.revision = newRevision;
          this.revisionMap.set(entry.id, newRevision);
        }
        this.diff.delete(entry.id);
        this.clearInvalidStateForRecord(entry.id);
      });
    }

    tryClose() {
      if (this.hasUnsavedChanges()) {
        const ok = window.confirm(resolveText(this.language, 'confirmClose'));
        if (!ok) return;
      }
      this.close(true);
    }

    close(force = false) {
      if (!force && this.saving) return;
      this.isOpen = false;
      this.closeColumnManager(true);
      this.closeRadioPicker();
      if (this.requireReload) {
        this.requireReload = false;
        this.reloading = true;
        this.detachEvents();
        if (this.surface) {
          this.surface.classList.add('pb-overlay__surface--reloading');
        }
        this.showLoading(resolveText(this.language, 'reloading'));
        if (document.body) {
          document.body.style.overflow = this.bodyOverflowBackup || '';
        }
        this.bodyOverflowBackup = '';
        this.previousFocus = null;
        setTimeout(() => {
          try {
            window.location.reload();
          } catch (_e) {
            window.location.href = window.location.href;
          }
        }, 50);
        return;
      }
      this.detachEvents();
      if (this.root) {
        this.root.remove();
      }
      this.root = null;
      this.surface = null;
      this.bodyScroll = null;
      this.columnHeaderScroll = null;
      this.rowHeaderScroll = null;
      this.toastHost = null;
      this.loadingEl = null;
      this.dirtyBadge = null;
      this.saveButton = null;
      this.closeButton = null;
      this.titleElement = null;
      this.appName = '';
      this.statsElements = { sum: null, avg: null, count: null };
      this.fields = [];
      this.fieldMap.clear();
      this.fieldIndexMap.clear();
      this.rows = [];
      this.rowMap.clear();
      this.rowIndexMap.clear();
      this.inputsByRow.clear();
      this.diff.clear();
      this.invalidCells.clear();
      this.revisionMap.clear();
      this.invalidValueCache.clear();
      this.rowHeaderContainer = null;
      this.tableContainer = null;
      this.virtualStart = 0;
      this.virtualEnd = 0;
      this.pendingFocus = null;
      this.editingCell = null;
      this.selection = null;
      this.dragSelecting = false;
      if (document.body) {
        document.body.style.overflow = this.bodyOverflowBackup || '';
      }
      this.bodyOverflowBackup = '';
      if (this.previousFocus && document.contains(this.previousFocus)) {
        const target = this.previousFocus;
        setTimeout(() => {
          try {
            target.focus({ preventScroll: true });
          } catch (_e) { /* noop */ }
        }, 0);
      }
      this.previousFocus = null;
      if (this.requireReload) {
        setTimeout(() => {
          try {
            if (typeof window.kintone?.app?.getId === 'function') {
              window.location.reload();
            } else {
              window.location.reload();
            }
          } catch (_e) {
            window.location.reload();
          }
        }, 50);
      }
      this.requireReload = false;
    }

    showLoading(text) {
      if (!this.surface) return;
      if (!this.loadingEl) {
        const loader = document.createElement('div');
        loader.className = 'pb-overlay__loading';
        this.loadingEl = loader;
        this.surface.appendChild(loader);
      }
      this.loadingEl.textContent = text;
      this.loadingEl.classList.add('is-active');
    }

    hideLoading() {
      if (this.loadingEl) {
        this.loadingEl.classList.remove('is-active');
      }
    }

    notify(message) {
      if (!this.toastHost) return;
      const toast = document.createElement('div');
      toast.className = 'pb-overlay__toast';
      toast.textContent = message;
      this.toastHost.appendChild(toast);
      requestAnimationFrame(() => {
        toast.classList.add('is-visible');
      });
      const remove = () => {
        toast.classList.remove('is-visible');
        setTimeout(() => toast.remove(), 250);
      };
      let timer = setTimeout(remove, 2200);
      toast.addEventListener('mouseenter', () => {
        clearTimeout(timer);
      });
      toast.addEventListener('mouseleave', () => {
        timer = setTimeout(remove, 800);
      });
    }

    clearInvalidStateForRecord(recordId) {
      const prefix = `${recordId}:`;
      const toDelete = [];
      this.invalidCells.forEach((key) => {
        if (key.startsWith(prefix)) toDelete.push(key);
      });
      toDelete.forEach((key) => {
        this.invalidCells.delete(key);
        this.invalidValueCache.delete(key);
      });
    }

    refreshRowFromRemote(record) {
      const latest = this.transformRecord(record);
      if (!latest.id) return;
      let existing = this.rowMap.get(latest.id);
      if (!existing) {
        existing = latest;
        this.rows.push(existing);
        const newIndex = this.rows.length - 1;
        this.rowMap.set(latest.id, existing);
        this.rowIndexMap.set(latest.id, newIndex);
      } else {
        existing.original = latest.original;
        existing.values = { ...latest.values };
        existing.revision = latest.revision;
        const rowIndex = this.rowIndexMap.has(latest.id)
          ? this.rowIndexMap.get(latest.id)
          : this.rows.findIndex((r) => r.id === latest.id);
        if (rowIndex !== undefined && rowIndex >= 0) {
          this.rowIndexMap.set(latest.id, rowIndex);
        }
      }
      if (latest.revision !== undefined) {
        this.revisionMap.set(latest.id, latest.revision);
      }
      const diffEntry = this.diff.get(latest.id) || {};
      this.fields.forEach((field) => {
        const input = this.findInput(latest.id, field.code);
        const key = `${latest.id}:${field.code}`;
        this.invalidCells.delete(key);
        this.invalidValueCache.delete(key);
        if (!input) return;
        input.dataset.originalValue = existing.original[field.code];
        if (!diffEntry[field.code]) {
          input.value = existing.values[field.code];
        }
        this.applyCellVisualState(input, existing, field.code);
      });
      this.updateVirtualRows(true);
    }

    reapplyDiffForRecord(recordId) {
      const entry = this.diff.get(recordId);
      if (!entry) return;
      const row = this.rowMap.get(recordId);
      if (!row) return;
      const pairs = Object.entries(entry);
      pairs.forEach(([fieldCode, info]) => {
        const input = this.findInput(recordId, fieldCode);
        if (!input) return;
        input.dataset.originalValue = row.original[fieldCode];
        input.value = info.value ?? '';
        this.onInputChanged(input);
      });
    }

    async handleConflict(batch) {
      const ids = Array.from(new Set(batch.map((entry) => entry.id).filter(Boolean))).map(String);
      if (!ids.length) return 'fail';
      const fieldCodes = this.fields.map((f) => f.code);
      const response = await this.postFn('EXCEL_GET_RECORDS_BY_IDS', {
        appId: this.appId,
        ids,
        fields: fieldCodes
      });
      if (!response?.ok) {
        return 'fail';
      }
      const records = Array.isArray(response.records) ? response.records : [];
      if (!records.length) return 'fail';
      records.forEach((record) => {
        this.refreshRowFromRemote(record);
      });
      ids.forEach((id) => {
        this.clearInvalidStateForRecord(id);
        this.reapplyDiffForRecord(id);
      });
      if (!this.diff.size) {
        return 'cleared';
      }
      return 'retry';
    }
  }

  const overlayController = new ExcelOverlayController(postToPage);

  document.addEventListener('keydown', (event) => {
    if (overlayController.isOpen) return;
    if (event.key.toLowerCase() === 'e' && event.shiftKey && (event.ctrlKey || event.metaKey)) {
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      overlayController.open();
    }
  }, true);

})();













