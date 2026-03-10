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

  const devContextState = {
    target: null,
    lastContextInfo: null
  };

  function normalizeFieldCode(value) {
    const code = String(value || '').trim();
    return code ? code : '';
  }

  function pickContextTarget(eventTarget) {
    if (!eventTarget) return null;
    if (eventTarget instanceof Element) return eventTarget;
    if (eventTarget.parentElement) return eventTarget.parentElement;
    return null;
  }

  function normalizeContextText(value) {
    return String(value || '').replace(/\s+/g, ' ').trim().slice(0, 200);
  }

  function computeCellLikeIndex(cell) {
    if (!cell) return null;
    if (typeof cell.cellIndex === 'number' && cell.cellIndex >= 0) return cell.cellIndex;
    const row = cell.parentElement;
    if (!row) return null;
    const siblings = Array.from(row.children).filter((node) => {
      if (!(node instanceof Element)) return false;
      return node.matches('th,td,[role="columnheader"],[role="gridcell"]');
    });
    const index = siblings.indexOf(cell);
    return index >= 0 ? index : null;
  }

  function createLastContextInfo(target, clientX, clientY) {
    const overlayNode = target?.closest?.('[data-field-code]') || null;
    const overlayFieldCode = normalizeFieldCode(overlayNode?.getAttribute?.('data-field-code') || '');
    const cellNode = target?.closest?.('th,td,[role="columnheader"],[role="gridcell"]') || null;
    const role = String(cellNode?.getAttribute?.('role') || '').toLowerCase();
    const tag = String(cellNode?.tagName || '').toUpperCase();
    const isHeader = tag === 'TH' || role === 'columnheader';
    const cellIndex = computeCellLikeIndex(cellNode);
    const targetTag = target?.tagName ? String(target.tagName).toLowerCase() : null;
    const targetText = normalizeContextText(target?.textContent || '');
    return {
      x: Number(clientX || 0),
      y: Number(clientY || 0),
      timestamp: Date.now(),
      targetTag,
      targetText,
      overlayFieldCode: overlayFieldCode || null,
      cellIndex: typeof cellIndex === 'number' ? cellIndex : null,
      isHeader,
      isListLike: Boolean(cellNode || overlayFieldCode),
      hasTarget: Boolean(target)
    };
  }

  function debugDevField(stage, payload) {
    try {
      console.log(`[kfav][DEV] field resolve ${stage}`, payload);
    } catch (_err) {
      // ignore
    }
  }

  document.addEventListener('contextmenu', (event) => {
    const target = pickContextTarget(event.target);
    devContextState.target = target;
    devContextState.lastContextInfo = createLastContextInfo(target, event.clientX, event.clientY);
  }, true);

  let currentPageContext = {
    href: location.href,
    timestamp: Date.now()
  };

  function notifyPageContextUpdated(trigger) {
    currentPageContext = {
      href: location.href,
      timestamp: Date.now()
    };
    try {
      chrome.runtime.sendMessage({
        type: 'PAGE_CONTEXT_UPDATED',
        payload: {
          href: currentPageContext.href,
          timestamp: currentPageContext.timestamp,
          trigger: String(trigger || 'unknown')
        }
      }).catch(() => {});
    } catch (_err) {
      // ignore
    }
  }

  document.addEventListener('click', (event) => {
    if (event.button !== 0) return;
    notifyPageContextUpdated('click');
  }, true);
  window.addEventListener('hashchange', () => notifyPageContextUpdated('hashchange'), true);
  window.addEventListener('popstate', () => notifyPageContextUpdated('popstate'), true);

  function wrapHistoryMethodForContext(name) {
    const original = history[name];
    if (typeof original !== 'function') return;
    history[name] = function patchedHistoryMethod(...args) {
      const result = original.apply(this, args);
      notifyPageContextUpdated(name);
      return result;
    };
  }
  wrapHistoryMethodForContext('pushState');
  wrapHistoryMethodForContext('replaceState');
  notifyPageContextUpdated('init');

  async function writeClipboardText(text) {
    const value = String(text || '');
    if (!value) return false;
    try {
      if (navigator?.clipboard?.writeText) {
        await navigator.clipboard.writeText(value);
        return true;
      }
    } catch (_err) {
      // fallback below
    }
    try {
      const ta = document.createElement('textarea');
      ta.value = value;
      ta.setAttribute('readonly', 'readonly');
      ta.style.position = 'fixed';
      ta.style.left = '-9999px';
      ta.style.top = '-9999px';
      document.body.appendChild(ta);
      ta.select();
      ta.setSelectionRange(0, ta.value.length);
      const ok = document.execCommand('copy');
      ta.remove();
      return Boolean(ok);
    } catch (_err) {
      return false;
    }
  }

  async function resolveDevQuery() {
    const context = await postToPage('DEV_GET_LIST_CONTEXT', {});
    if (!context?.ok) return { ok: false, reason: context?.error || 'context_unavailable' };
    const query = String(context.query || '').trim();
    if (!query) return { ok: false, reason: 'query_not_found' };
    return { ok: true, value: query };
  }

  function resolveFieldCodeFromTarget(target) {
    if (!target || typeof target.closest !== 'function') return '';
    const direct = target.closest('[data-field-code]');
    if (direct?.getAttribute) {
      const code = normalizeFieldCode(direct.getAttribute('data-field-code'));
      if (code) return code;
    }
    return '';
  }

  async function resolveFieldCodeByViewMeta(lastContextInfo) {
    const colIndex = Number(lastContextInfo?.cellIndex);
    if (!Number.isInteger(colIndex) || colIndex < 0) {
      return { ok: false, reason: 'cell_index_not_found' };
    }
    const context = await postToPage('DEV_GET_LIST_CONTEXT', {});
    if (!context?.ok) {
      return { ok: false, reason: 'view_meta_missing' };
    }
    const meta = await postToPage('DEV_GET_VIEW_META', {
      appId: context?.appId || '',
      viewId: context?.viewId || ''
    });
    if (!meta?.ok) {
      return { ok: false, reason: 'view_meta_missing' };
    }
    const columns = Array.isArray(meta.columns) ? meta.columns : [];
    debugDevField('candidates', {
      lastContextInfo,
      detectedCellIndex: colIndex,
      isHeader: Boolean(lastContextInfo?.isHeader),
      viewColumns: columns
    });
    const candidateIndexes = [colIndex, colIndex - 1, colIndex + 1, colIndex - 2, colIndex + 2]
      .filter((idx) => idx >= 0);
    for (const idx of candidateIndexes) {
      const hit = columns.find((column) => Number(column?.index) === idx) || columns[idx];
      const code = normalizeFieldCode(hit?.fieldCode);
      if (code) {
        return { ok: true, value: code };
      }
    }
    return { ok: false, reason: 'field_code_not_resolved' };
  }

  async function resolveDevFieldCode() {
    const target = devContextState.target;
    const lastContextInfo = devContextState.lastContextInfo;
    if (!target) {
      debugDevField('result', { ok: false, reason: 'no_context_target' });
      return { ok: false, reason: 'no_context_target' };
    }
    debugDevField('start', {
      lastContextInfo
    });

    const overlayFieldCode = normalizeFieldCode(lastContextInfo?.overlayFieldCode || '');
    if (overlayFieldCode) {
      debugDevField('result', { ok: true, source: 'overlay', fieldCode: overlayFieldCode });
      return { ok: true, value: overlayFieldCode };
    }

    const viaViewMeta = await resolveFieldCodeByViewMeta(lastContextInfo);
    if (viaViewMeta?.ok && viaViewMeta?.value) {
      debugDevField('result', { ok: true, source: 'view_meta', fieldCode: viaViewMeta.value });
      return viaViewMeta;
    }

    const fallbackCode = resolveFieldCodeFromTarget(target);
    if (fallbackCode) {
      debugDevField('result', { ok: true, source: 'target_dataset', fieldCode: fallbackCode });
      return { ok: true, value: fallbackCode };
    }

    const reason = viaViewMeta?.reason || 'field_code_not_resolved';
    debugDevField('result', {
      ok: false,
      reason,
      overlayFieldCode: lastContextInfo?.overlayFieldCode || null,
      detectedCellIndex: lastContextInfo?.cellIndex ?? null,
      isHeader: Boolean(lastContextInfo?.isHeader)
    });
    return { ok: false, reason };
  }

  async function copyDevValue(handlerName) {
    const resolver = handlerName === 'query' ? resolveDevQuery : resolveDevFieldCode;
    const resolved = await resolver();
    if (!resolved?.ok || !resolved?.value) {
      return { ok: false, reason: resolved?.reason || 'not_found' };
    }
    const copied = await writeClipboardText(resolved.value);
    if (!copied) {
      if (handlerName === 'field_code') {
        debugDevField('result', { ok: false, reason: 'clipboard_failed' });
      }
      return { ok: false, reason: 'clipboard_failed' };
    }
    console.log('[kfav][DEV] copied', { type: handlerName, value: resolved.value });
    return { ok: true, value: resolved.value };
  }

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    (async () => {
      try {
        if (msg?.type === 'DEV_COPY_QUERY') {
          const result = await copyDevValue('query');
          sendResponse(result); return;
        }
        if (msg?.type === 'DEV_COPY_TEXT') {
          const text = String(msg?.text || '');
          const label = String(msg?.label || 'text');
          if (!text) {
            sendResponse({ ok: false, reason: 'empty_text' }); return;
          }
          const copied = await writeClipboardText(text);
          if (!copied) {
            console.warn('[kfav][DEV] copy failed', { reason: 'clipboard_failed', label });
            sendResponse({ ok: false, reason: 'clipboard_failed' }); return;
          }
          console.log('[kfav][DEV] copied', { type: label });
          sendResponse({ ok: true }); return;
        }
        if (msg?.type === 'DEV_COPY_FIELD_CODE') {
          const result = await copyDevValue('field_code');
          sendResponse(result); return;
        }
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
        if (msg?.type === 'GET_APP_NAME') {
          const res = await postToPage('GET_APP_NAME', msg.payload);
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

  const RECENT_TRACK_INTERVAL_MS = 500;
  const RECENT_UPSERT_COOLDOWN_MS = 2000;
  let lastObservedHref = '';
  let lastRecentId = '';
  let lastRecentAt = 0;

  function isKintoneHostName(hostname) {
    return /\.kintone(?:-dev)?\.com$/.test(hostname) || /\.cybozu\.com$/.test(hostname);
  }

  function parseRecentRecordFromUrl(rawUrl) {
    try {
      const url = new URL(rawUrl);
      if (!isKintoneHostName(url.hostname)) return null;
      const appMatch = url.pathname.match(/\/k\/(\d+)\/show(?:\/|$)/);
      if (!appMatch) return null;
      const appId = appMatch[1];
      const hash = String(url.hash || '');
      const hashMatch = hash.match(/(?:^#|[?&])record=(\d+)/);
      const searchMatch = String(url.search || '').match(/[?&]record=(\d+)/);
      const recordId = (hashMatch && hashMatch[1]) || (searchMatch && searchMatch[1]) || '';
      if (!recordId) return null;
      const host = url.origin.replace(/\/$/, '');
      const normalizedUrl = `${host}/k/${encodeURIComponent(appId)}/show#record=${encodeURIComponent(recordId)}`;
      let appName = '';
      try {
        if (typeof window.kintone?.app?.getName === 'function') {
          appName = String(window.kintone.app.getName() || '').trim();
        }
      } catch (_err) {
        appName = '';
      }
      const now = Date.now();
      return {
        id: `${host}|${appId}|${recordId}`,
        host,
        appId,
        recordId,
        appName,
        url: normalizedUrl,
        visitedAt: now
      };
    } catch (_err) {
      return null;
    }
  }

  function tryUpsertRecentRecord() {
    const record = parseRecentRecordFromUrl(location.href);
    if (!record) return;
    const now = Date.now();
    if (record.id === lastRecentId && (now - lastRecentAt) < RECENT_UPSERT_COOLDOWN_MS) return;
    lastRecentId = record.id;
    lastRecentAt = now;
    chrome.runtime.sendMessage({ type: 'RECENT_UPSERT', payload: record }).catch(() => {});
  }

  function startRecentRecordWatcher() {
    lastObservedHref = location.href;
    tryUpsertRecentRecord();
    setInterval(() => {
      const currentHref = location.href;
      if (currentHref === lastObservedHref) return;
      lastObservedHref = currentHref;
      tryUpsertRecentRecord();
    }, RECENT_TRACK_INTERVAL_MS);
  }

  startRecentRecordWatcher();

  const overlayTexts = {
    ja: {
      title: "Excelモード",
      loading: "読み込み中…",
      loadFailed: "データの取得に失敗しました",
      noEditableFields: "編集可能なフィールドがありません",
      btnSave: "保存",
      btnSaving: "保存中...",
      btnUndo: "↶ Undo",
      btnRedo: "↷ Redo",
      titleUndo: "元に戻す (Ctrl+Z)",
      titleRedo: "やり直し (Ctrl+Y)",
      btnAddRow: "＋ 行追加",
      btnPrev: "◀ 前へ",
      btnNext: "次へ ▶",
      btnClose: "閉じる",
      modeViewOnly: "Standard",
      toastNoChanges: "変更はありません",
      toastInvalidCells: "エラーを解消してから保存してください",
      toastRequiredMissing: "必須項目を入力してください",
      toastViewOnlyBlocked: "Proモードで利用できます",
      overlayProOnly: "Proモードで利用できます",
      overlayStandardReadonly: "Standardでは閲覧のみ利用できます",
      overlayProComingSoon: "Proモードは近日公開予定です",
      lookupAutoReadonly: "LOOKUPにより自動入力されるため編集できません",
      toastSaveSuccess: "保存しました",
      toastSaveFailed: "保存に失敗しました",
      confirmClose: "未保存の変更があります。閉じますか？",
      confirmPageMoveUnsaved: "未保存の変更があります。ページを移動しますか？",
      reloading: "一覧を更新しています…",
      conflictRetry: "最新のデータを反映して再保存します…",
      conflictNoChanges: "最新のデータと一致しました",
      conflictFailed: "他の変更と競合し保存できませんでした",
      btnSubtableEdit: "編集",
      subtableTitle: "サブテーブル編集",
      subtableClose: "閉じる",
      subtableAddRow: "＋ 行追加",
      subtableRemoveRow: "削除",
      subtableSave: "保存",
      subtableCancel: "キャンセル",
      subtableRows: (n) => `${n} 行`,
      subtableEmpty: "行がありません",
      subtableActionsAria: "操作",
      permNoAdd: "追加権限がありません",
      permNoEdit: "編集権限がありません",
      permNoFieldEdit: "このフィールドは編集できません",
      permNoDelete: "削除権限がありません",
      permUnknown: "権限未判定",
      permWarningUnknown: "権限を確認できないため、一部編集を無効化しています",
      permViewOnly: "閲覧のみ",
      permEditable: "編集可",
      permDeletable: "削除可",
      rowDeletePending: "削除予定",
      rowDeleteToggle: "削除予定にする",
      rowDeleteUndo: "削除予定を解除",
      rowDeleteNewUndo: "新規行を取消",
      rowDeletePendingTitle: "削除予定（保存時に削除）",
      toastPasteSkipped: (n) => `${n}件のセルを権限によりスキップしました`,
      toastCopySuccess: "コピーしました",
      toastCopyFailed: "コピーに失敗しました",
      dirtyLabel: (n) => `未保存: ${n}`,
      statusSum: "合計",
      statusAvg: "平均",
      statusCount: "件数",
      statusSelectedCells: "選択セル",
      statusFilteredRows: "フィルタ後",
      statusPendingDeletes: "削除予定",
      statusNewRows: "新規行",
      btnColumns: "列順",
      filterButton: "フィルター",
      filterTitle: "フィルター",
      filterContains: "含む",
      filterEquals: "一致",
      filterStartsWith: "前方一致",
      filterFrom: "開始",
      filterTo: "終了",
      filterClear: "クリア",
      filterApply: "適用",
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
      btnSaving: "Saving...",
      btnUndo: "↶ Undo",
      btnRedo: "↷ Redo",
      titleUndo: "Undo (Ctrl+Z)",
      titleRedo: "Redo (Ctrl+Y)",
      btnAddRow: "+ Add row",
      btnPrev: "◀ Prev",
      btnNext: "Next ▶",
      btnClose: "Close",
      modeViewOnly: "Standard",
      toastNoChanges: "No changes",
      toastInvalidCells: "Fix errors before saving",
      toastRequiredMissing: "Please fill required fields",
      toastViewOnlyBlocked: "Available in Pro mode",
      overlayProOnly: "Available in Pro mode",
      overlayStandardReadonly: "Editing is disabled in Standard mode",
      overlayProComingSoon: "Pro mode is coming soon",
      lookupAutoReadonly: "This field is auto-populated by LOOKUP and cannot be edited",
      toastSaveSuccess: "Changes saved",
      toastSaveFailed: "Failed to save changes",
      confirmClose: "You have unsaved changes. Close anyway?",
      confirmPageMoveUnsaved: "You have unsaved changes. Move to another page?",
      reloading: "Refreshing list…",
      conflictRetry: "Data updated, retrying save…",
      conflictNoChanges: "Your edits now match the latest data",
      conflictFailed: "Conflicted with other changes; save aborted",
      btnSubtableEdit: "Edit",
      subtableTitle: "Edit subtable",
      subtableClose: "Close",
      subtableAddRow: "+ Add row",
      subtableRemoveRow: "Delete",
      subtableSave: "Save",
      subtableCancel: "Cancel",
      subtableRows: (n) => (n === 1 ? '1 row' : `${n} rows`),
      subtableEmpty: "No rows",
      subtableActionsAria: "Actions",
      permNoAdd: "No add permission",
      permNoEdit: "No edit permission",
      permNoFieldEdit: "This field cannot be edited",
      permNoDelete: "No delete permission",
      permUnknown: "Permission unknown",
      permWarningUnknown: "Permissions could not be verified, so some editing is disabled",
      permViewOnly: "View only",
      permEditable: "Editable",
      permDeletable: "Deletable",
      rowDeletePending: "Pending delete",
      rowDeleteToggle: "Mark row for delete",
      rowDeleteUndo: "Unmark delete",
      rowDeleteNewUndo: "Discard new row",
      rowDeletePendingTitle: "Pending delete (will delete on save)",
      toastPasteSkipped: (n) => `${n} cells were skipped due to permissions`,
      toastCopySuccess: "Copied",
      toastCopyFailed: "Copy failed",
      dirtyLabel: (n) => `Unsaved: ${n}`,
      statusSum: "Sum",
      statusAvg: "Avg",
      statusCount: "Count",
      statusSelectedCells: "Selected",
      statusFilteredRows: "Filtered",
      statusPendingDeletes: "Pending delete",
      statusNewRows: "New rows",
      btnColumns: "Columns",
      filterButton: "Filter",
      filterTitle: "Filter",
      filterContains: "contains",
      filterEquals: "equals",
      filterStartsWith: "starts with",
      filterFrom: "from",
      filterTo: "to",
      filterClear: "Clear",
      filterApply: "Apply",
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
    allowedTypes: new Set(['SINGLE_LINE_TEXT', 'NUMBER', 'DATE', 'RADIO_BUTTON', 'DROP_DOWN', 'SUBTABLE', 'LOOKUP'])
  };

  const UI_LANGUAGE_KEY = 'uiLanguage';
  const UI_LANGUAGE_VALUES = ['auto', 'ja', 'en'];
  const DEFAULT_UI_LANGUAGE = 'auto';
  const DEFAULT_OVERLAY_LANGUAGE = 'ja';

  const COLUMN_PREF_STORAGE_KEY = 'kfavExcelColumns';
  const EXCEL_OVERLAY_MODE_STANDARD = 'standard';
  const EXCEL_OVERLAY_MODE_PRO = 'pro';
  const DEFAULT_EXCEL_OVERLAY_MODE = EXCEL_OVERLAY_MODE_STANDARD;
  const EXCEL_OVERLAY_MODE_VALUES = new Set([EXCEL_OVERLAY_MODE_STANDARD, EXCEL_OVERLAY_MODE_PRO]);
  const MAX_COLUMN_PREF_ENTRIES = 80;
  const DEFAULT_COLUMN_WIDTH = 160;
  const MIN_COLUMN_WIDTH = 96;
  const MAX_COLUMN_WIDTH = 320;

  let overlayCssLoaded = false;

  function normalizeUiLanguage(raw) {
    const value = String(raw || '').trim().toLowerCase();
    if (!value) return DEFAULT_OVERLAY_LANGUAGE;
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
      return DEFAULT_OVERLAY_LANGUAGE;
    }
  }

  function resolveEffectiveUiLanguage(setting) {
    if (setting === 'ja' || setting === 'en') return setting;
    return getBrowserUiLanguage();
  }

  async function resolveOverlayUiLanguage() {
    let setting = DEFAULT_UI_LANGUAGE;
    try {
      const stored = await chrome.storage.local.get(UI_LANGUAGE_KEY);
      setting = normalizeUiLanguageSetting(stored?.[UI_LANGUAGE_KEY]);
    } catch (_err) {
      setting = DEFAULT_UI_LANGUAGE;
    }
    return {
      setting,
      language: resolveEffectiveUiLanguage(setting)
    };
  }

  function resolveText(language, key, ...args) {
    const primary = overlayTexts[language] || {};
    const value = primary[key];
    if (typeof value === 'function') return value(...args);
    if (typeof value === 'string') return value;
    const fallbackJa = overlayTexts.ja?.[key];
    if (typeof fallbackJa === 'function') return fallbackJa(...args);
    if (typeof fallbackJa === 'string') return fallbackJa;
    return key;
  }

  function normalizeExcelOverlayMode(value) {
    if (value === 'edit') return EXCEL_OVERLAY_MODE_PRO;
    if (value === 'view' || value === 'off') return EXCEL_OVERLAY_MODE_STANDARD;
    return EXCEL_OVERLAY_MODE_VALUES.has(value) ? value : DEFAULT_EXCEL_OVERLAY_MODE;
  }

  function createPermissionServiceSafe() {
    if (typeof createPermissionService === 'function') {
      return createPermissionService();
    }
    return {
      init() {},
      setFields() {},
      setAppPermission() {},
      setPendingDeletes() {},
      refreshRecordAcl() {},
      upsertRecordAcl() {},
      removeRecordAcl() {},
      clearRecordAcl() {},
      canEditRecord() { return false; },
      canDeleteRecord() { return false; },
      canAddRow() { return false; },
      canEditCell() { return false; },
      getCellPermission() { return { editable: false, reason: 'no_acl' }; },
      filterPasteTargets(targets) {
        return { applicable: [], skipped: Array.isArray(targets) ? targets : [] };
      },
      isSystemField() { return true; },
      isLookupField() { return false; },
      isFieldTypeEditable() { return false; }
    };
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
      this.permissionService = createPermissionServiceSafe();
      this.isOpen = false;
      this.root = null;
      this.surface = null;
      this.bodyScroll = null;
      this.columnHeaderScroll = null;
      this.rowHeaderScroll = null;
      this.toastHost = null;
      this.loadingEl = null;
      this.permissionWarningEl = null;
      this.dirtyBadge = null;
      this.saveButton = null;
      this.saveButtonSpinner = null;
      this.saveButtonLabel = null;
      this.undoButton = null;
      this.redoButton = null;
      this.closeButton = null;
      this.addRowButton = null;
      this.titleElement = null;
      this.prevPageButton = null;
      this.nextPageButton = null;
      this.pageLabelElement = null;
      this.appName = '';
      this.statsElements = {
        sum: null,
        avg: null,
        count: null,
        selected: null,
        filtered: null,
        pendingDelete: null,
        newRows: null
      };
      this.language = 'ja';
      this.uiLanguageSetting = DEFAULT_UI_LANGUAGE;
      this.appName = '';
      this.overlayMode = DEFAULT_EXCEL_OVERLAY_MODE;
      this.proEntitlement = false;
      this.modeBadge = null;
      this.viewOnlyNoticeShown = false;
      this.appId = null;
      this.baseQuery = '';
      this.query = '';
      this.pageSize = 100;
      this.pageOffset = 0;
      this.totalCount = 0;
      this.paging = false;
      this.fields = [];
      this.fieldsMeta = [];
      this.fieldMap = new Map();
      this.fieldsMetaMap = new Map();
      this.fieldIndexMap = new Map();
      this.rows = [];
      this.rowMap = new Map();
      this.rowIndexMap = new Map();
      this.aclStatus = { loaded: false, failed: false };
      this.pendingDeletes = new Set();
      this.filters = new Map();
      this.filteredRowIds = null;
      this.filterPanel = null;
      this.needsFilterReapply = false;
      this.sortState = { fieldCode: '', direction: '' };
      this.undoStack = [];
      this.redoStack = [];
      this.isReplayingHistory = false;
      this.maxHistory = 100;
      this.diff = new Map();
      this.invalidCells = new Set();
      this.revisionMap = new Map();
      this.isComposing = false;
      this.saving = false;
      this.handleScroll = this.syncScroll.bind(this);
      this.handleKeyOnOverlay = this.handleOverlayKeydown.bind(this);
      this.handleGlobalArrowKeyCapture = this.handleOverlayGlobalArrowKeydown.bind(this);
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
      this.pendingEdit = null;
      this.invalidValueCache = new Map();
      this.selection = null;
      this.armedCell = null;
      this.navigating = false;
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
      this.newRowSeq = 1;
      this.subtableEditor = null;
      this.pasteFlashTimer = null;
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
      this.boundFilterPanelOutsideClick = (event) => {
        if (!this.filterPanel) return;
        const panel = this.filterPanel.panel;
        const button = this.filterPanel.button;
        if (!panel) return;
        if (panel.contains(event.target)) return;
        if (button && (button === event.target || button.contains(event.target))) return;
        this.closeFilterPanel();
      };
      this.boundUiLanguageStorageChanged = this.handleUiLanguageStorageChanged.bind(this);
      this.uiLanguageListenerAttached = false;
    }

    async open() {
      if (this.isOpen) return;
      await this.loadOverlayMode();
      console.log('[overlay] mode flags', this.getOverlayModeFlags());
      this.isOpen = true;
      try {
        this.previousFocus = document.activeElement instanceof HTMLElement ? document.activeElement : null;
        this.bodyOverflowBackup = document.body.style.overflow || '';
        document.body.style.overflow = 'hidden';
        this.requireReload = false;
        this.viewOnlyNoticeShown = false;
        await this.ensureStyles();
        await this.loadLanguage();
        this.attachUiLanguageListener();
        this.mountShell();
        this.showLoading(resolveText(this.language, 'loading'));
        await this.fetchData();
        this.hideLoading();
        if (!this.fields.length) {
          this.renderNoFields();
          return;
        }
        console.log('[overlay] render grid start');
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
      const permissionWarning = document.createElement('div');
      permissionWarning.className = 'pb-overlay__perm-warning pb-overlay__perm-warning--hidden';
      permissionWarning.textContent = resolveText(this.language, 'permWarningUnknown');
      this.permissionWarningEl = permissionWarning;
      const grid = this.buildGridShell();
      const status = this.buildStatusBar();
      const toastHost = document.createElement('div');
      toastHost.className = 'pb-overlay__toast-host toast-container';
      this.toastHost = toastHost;

      surface.appendChild(toolbar);
      surface.appendChild(permissionWarning);
      surface.appendChild(grid);
      surface.appendChild(status);
      this.root.appendChild(surface);
      this.root.appendChild(toastHost);
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
      if (this.isOverlayViewOnly()) {
        const badge = document.createElement('span');
        badge.className = 'pb-overlay__mode-badge';
        badge.textContent = resolveText(this.language, 'modeViewOnly');
        this.modeBadge = badge;
        titleWrap.appendChild(badge);
      }

      const actions = document.createElement('div');
      actions.className = 'pb-overlay__toolbar-actions';

      const pager = document.createElement('div');
      pager.className = 'pb-overlay__pager';
      const prevBtn = document.createElement('button');
      prevBtn.type = 'button';
      prevBtn.className = 'pb-overlay__btn pb-overlay__btn--pager';
      prevBtn.textContent = resolveText(this.language, 'btnPrev');
      this.prevPageButton = prevBtn;
      const pageLabel = document.createElement('span');
      pageLabel.className = 'pb-overlay__pager-label';
      pageLabel.textContent = '0-0 / 0';
      this.pageLabelElement = pageLabel;
      const nextBtn = document.createElement('button');
      nextBtn.type = 'button';
      nextBtn.className = 'pb-overlay__btn pb-overlay__btn--pager';
      nextBtn.textContent = resolveText(this.language, 'btnNext');
      this.nextPageButton = nextBtn;
      pager.appendChild(prevBtn);
      pager.appendChild(pageLabel);
      pager.appendChild(nextBtn);

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
      const saveSpinner = document.createElement('span');
      saveSpinner.className = 'pb-overlay__spinner pb-overlay__spinner--hidden';
      saveSpinner.setAttribute('aria-hidden', 'true');
      const saveLabel = document.createElement('span');
      saveLabel.className = 'pb-overlay__save-label';
      saveLabel.textContent = resolveText(this.language, 'btnSave');
      saveBtn.appendChild(saveSpinner);
      saveBtn.appendChild(saveLabel);
      this.saveButton = saveBtn;
      this.saveButtonSpinner = saveSpinner;
      this.saveButtonLabel = saveLabel;

      const addRowBtn = document.createElement('button');
      addRowBtn.type = 'button';
      addRowBtn.className = 'pb-overlay__btn';
      addRowBtn.textContent = resolveText(this.language, 'btnAddRow');
      this.addRowButton = addRowBtn;

      const undoBtn = document.createElement('button');
      undoBtn.type = 'button';
      undoBtn.className = 'pb-overlay__btn';
      undoBtn.textContent = resolveText(this.language, 'btnUndo');
      undoBtn.title = resolveText(this.language, 'titleUndo');
      undoBtn.disabled = true;
      this.undoButton = undoBtn;

      const redoBtn = document.createElement('button');
      redoBtn.type = 'button';
      redoBtn.className = 'pb-overlay__btn';
      redoBtn.textContent = resolveText(this.language, 'btnRedo');
      redoBtn.title = resolveText(this.language, 'titleRedo');
      redoBtn.disabled = true;
      this.redoButton = redoBtn;

      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.className = 'pb-overlay__btn';
      closeBtn.textContent = resolveText(this.language, 'btnClose');
      this.closeButton = closeBtn;

      actions.appendChild(pager);
      actions.appendChild(columnsBtn);
      actions.appendChild(dirty);
      actions.appendChild(addRowBtn);
      actions.appendChild(undoBtn);
      actions.appendChild(redoBtn);
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

      const selected = document.createElement('span');
      selected.className = 'pb-overlay__status-item';
      this.statsElements.selected = selected;

      const filtered = document.createElement('span');
      filtered.className = 'pb-overlay__status-item';
      this.statsElements.filtered = filtered;

      const pendingDelete = document.createElement('span');
      pendingDelete.className = 'pb-overlay__status-item';
      this.statsElements.pendingDelete = pendingDelete;

      const newRows = document.createElement('span');
      newRows.className = 'pb-overlay__status-item';
      this.statsElements.newRows = newRows;

      const sum = document.createElement('span');
      sum.className = 'pb-overlay__status-item';
      this.statsElements.sum = sum;

      const avg = document.createElement('span');
      avg.className = 'pb-overlay__status-item';
      this.statsElements.avg = avg;

      const count = document.createElement('span');
      count.className = 'pb-overlay__status-item';
      this.statsElements.count = count;

      status.appendChild(selected);
      status.appendChild(filtered);
      status.appendChild(pendingDelete);
      status.appendChild(newRows);
      status.appendChild(sum);
      status.appendChild(avg);
      status.appendChild(count);

      return status;
    }

    attachEvents() {
      if (!this.root) return;
      this.bodyScroll.addEventListener('scroll', this.handleScroll, { passive: true });
      this.saveButton.addEventListener('click', () => { this.save(); });
      this.addRowButton?.addEventListener('click', () => { this.addNewRow(); });
      this.undoButton?.addEventListener('click', () => { this.undo(); });
      this.redoButton?.addEventListener('click', () => { this.redo(); });
      this.prevPageButton?.addEventListener('click', () => { this.prevPage(); });
      this.nextPageButton?.addEventListener('click', () => { this.nextPage(); });
      this.closeButton.addEventListener('click', () => { this.tryClose(); });
      this.root.addEventListener('keydown', this.handleKeyOnOverlay, true);
      document.addEventListener('keydown', this.handleGlobalArrowKeyCapture, true);
      document.addEventListener('mouseup', this.handleMouseUp, true);
    }

    detachEvents() {
      if (!this.root) return;
      if (this.bodyScroll) {
        this.bodyScroll.removeEventListener('scroll', this.handleScroll);
      }
      this.root.removeEventListener('keydown', this.handleKeyOnOverlay, true);
      document.removeEventListener('keydown', this.handleGlobalArrowKeyCapture, true);
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
      const resolved = await resolveOverlayUiLanguage();
      this.uiLanguageSetting = resolved.setting;
      this.language = resolved.language;
    }

    attachUiLanguageListener() {
      if (this.uiLanguageListenerAttached) return;
      if (!chrome?.storage?.onChanged?.addListener) return;
      chrome.storage.onChanged.addListener(this.boundUiLanguageStorageChanged);
      this.uiLanguageListenerAttached = true;
    }

    detachUiLanguageListener() {
      if (!this.uiLanguageListenerAttached) return;
      if (!chrome?.storage?.onChanged?.removeListener) return;
      chrome.storage.onChanged.removeListener(this.boundUiLanguageStorageChanged);
      this.uiLanguageListenerAttached = false;
    }

    handleUiLanguageStorageChanged(changes, areaName) {
      if (areaName !== 'local' || !changes || !Object.prototype.hasOwnProperty.call(changes, UI_LANGUAGE_KEY)) {
        return;
      }
      const nextSetting = normalizeUiLanguageSetting(changes[UI_LANGUAGE_KEY]?.newValue);
      const nextLanguage = resolveEffectiveUiLanguage(nextSetting);
      if (nextSetting === this.uiLanguageSetting && nextLanguage === this.language) return;
      this.uiLanguageSetting = nextSetting;
      this.language = nextLanguage;
      if (!this.isOpen || !this.root) return;
      this.applyLanguageToOverlayUi();
    }

    applyLanguageToOverlayUi() {
      if (this.filterPanel?.panel) {
        this.closeFilterPanel();
      }
      if (this.titleElement) {
        this.titleElement.textContent = this.appName || resolveText(this.language, 'title');
      }
      if (this.modeBadge) {
        this.modeBadge.textContent = resolveText(this.language, 'modeViewOnly');
      }
      if (this.prevPageButton) {
        this.prevPageButton.textContent = resolveText(this.language, 'btnPrev');
      }
      if (this.nextPageButton) {
        this.nextPageButton.textContent = resolveText(this.language, 'btnNext');
      }
      if (this.columnButton) {
        this.columnButton.textContent = resolveText(this.language, 'btnColumns');
      }
      if (this.addRowButton) {
        this.addRowButton.textContent = resolveText(this.language, 'btnAddRow');
      }
      if (this.undoButton) {
        this.undoButton.textContent = resolveText(this.language, 'btnUndo');
        this.undoButton.title = resolveText(this.language, 'titleUndo');
      }
      if (this.redoButton) {
        this.redoButton.textContent = resolveText(this.language, 'btnRedo');
        this.redoButton.title = resolveText(this.language, 'titleRedo');
      }
      if (this.closeButton) {
        this.closeButton.textContent = resolveText(this.language, 'btnClose');
      }
      if (this.saveButtonLabel) {
        this.saveButtonLabel.textContent = resolveText(this.language, this.saving ? 'btnSaving' : 'btnSave');
      }
      if (this.permissionWarningEl) {
        this.permissionWarningEl.textContent = resolveText(this.language, 'permWarningUnknown');
      }
      this.refreshHeaderFilterButtonsText();
      this.refreshSubtableEditorTexts();
      if (this.tableContainer && this.rowHeaderContainer) {
        this.updateVirtualRows(true);
      }
      this.updatePermissionUiState();
      this.updateDirtyBadge();
      this.updateStats();
    }

    refreshHeaderFilterButtonsText() {
      const headRow = this.columnHeaderScroll?.querySelector('.pb-overlay__col-row');
      if (!headRow) return;
      const filterLabel = resolveText(this.language, 'filterButton');
      const headerCells = Array.from(headRow.querySelectorAll('.pb-overlay__col-header'));
      headerCells.forEach((cell, index) => {
        const btn = cell.querySelector('.pb-overlay__filter-btn');
        if (!btn) return;
        const field = this.fields[index];
        const fieldLabel = field ? (field.label || field.code) : '';
        btn.title = filterLabel;
        btn.setAttribute('aria-label', fieldLabel ? `${filterLabel}: ${fieldLabel}` : filterLabel);
      });
    }

    refreshSubtableEditorTexts() {
      if (!this.subtableEditor?.panel) return;
      const panel = this.subtableEditor.panel;
      const field = this.fields[this.subtableEditor.colIndex];
      const title = panel.querySelector('.pb-overlay__subtable-title');
      if (title) {
        title.textContent = `${resolveText(this.language, 'subtableTitle')}: ${field?.label || field?.code || ''}`;
      }
      const closeBtn = panel.querySelector('.pb-overlay__subtable-close');
      if (closeBtn) {
        closeBtn.title = resolveText(this.language, 'subtableClose');
        closeBtn.setAttribute('aria-label', resolveText(this.language, 'subtableClose'));
      }
      const addBtn = panel.querySelector('[data-subtable-action="add"]');
      if (addBtn) addBtn.textContent = resolveText(this.language, 'subtableAddRow');
      const cancelBtn = panel.querySelector('[data-subtable-action="cancel"]');
      if (cancelBtn) cancelBtn.textContent = resolveText(this.language, 'subtableCancel');
      const saveBtn = panel.querySelector('[data-subtable-action="save"]');
      if (saveBtn) saveBtn.textContent = resolveText(this.language, 'subtableSave');
      const opHeader = panel.querySelector('th.pb-subtable__op');
      if (opHeader) {
        opHeader.setAttribute('aria-label', resolveText(this.language, 'subtableActionsAria'));
      }
      panel.querySelectorAll('.pb-subtable__remove').forEach((btn) => {
        btn.title = resolveText(this.language, 'subtableRemoveRow');
        btn.setAttribute('aria-label', resolveText(this.language, 'subtableRemoveRow'));
      });
      const emptyCell = panel.querySelector('.pb-subtable__empty');
      if (emptyCell) {
        emptyCell.textContent = resolveText(this.language, 'subtableEmpty');
      }
    }

    async loadOverlayMode() {
      // Public build is fixed to Standard mode for CWS release.
      // Pro mode plumbing stays in place for future entitlement integration.
      this.overlayMode = EXCEL_OVERLAY_MODE_STANDARD;
      this.proEntitlement = false;
    }

    isProMode() {
      return this.overlayMode === EXCEL_OVERLAY_MODE_PRO;
    }

    canUseEditFeatures() {
      // Future expansion:
      // return this.isProMode() && this.proEntitlement && permissionService.canEdit...
      return this.isProMode() && this.proEntitlement;
    }

    canViewOverlay() {
      return true;
    }

    canEditOverlay() {
      return this.canUseEditFeatures();
    }

    canSaveOverlay() {
      return this.canEditOverlay();
    }

    getOverlayModeFlags() {
      return {
        overlayMode: this.overlayMode,
        canViewOverlay: this.canViewOverlay(),
        canEditOverlay: this.canEditOverlay(),
        canSaveOverlay: this.canSaveOverlay()
      };
    }

    isOverlayEditable() {
      return this.canEditOverlay();
    }

    isOverlayViewOnly() {
      return this.canViewOverlay() && !this.canEditOverlay();
    }

    canMutateOverlay(showNotice = false) {
      const allowed = this.canEditOverlay();
      if (!allowed && showNotice) {
        this.notifyViewOnlyBlocked();
      }
      return allowed;
    }

    notifyViewOnlyBlocked() {
      if (this.viewOnlyNoticeShown) return;
      this.viewOnlyNoticeShown = true;
      this.notify(resolveText(this.language, 'toastViewOnlyBlocked'));
      setTimeout(() => {
        this.viewOnlyNoticeShown = false;
      }, 1200);
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
      this.baseQuery = String(resolvedQuery || '').trim();
      this.query = this.baseQuery;
      this.pageOffset = 0;
      this.totalCount = 0;

      const fieldsMetaRes = await this.postFn('EXCEL_GET_FIELDS_META', { appId: this.appId });
      let rawFields = [];
      if (fieldsMetaRes?.ok && Array.isArray(fieldsMetaRes.fieldsMeta)) {
        const allMetaFields = fieldsMetaRes.fieldsMeta
          .map((f) => ({
            code: f.code,
            label: f.label || '',
            type: f.type || '',
            required: Boolean(f.required),
            choices: Array.isArray(f.choices) ? f.choices : [],
            lookupAuto: Boolean(f.lookupAuto),
            lookup: f.lookup ? { ...f.lookup } : undefined,
            subtable: f.subtable && Array.isArray(f.subtable.fields)
              ? {
                fields: f.subtable.fields.map((child) => ({
                  ...child,
                  required: Boolean(child?.required),
                  choices: Array.isArray(child.choices) ? child.choices : []
                }))
              }
              : undefined
          }))
          .filter((f) => this.shouldIncludeField(f));

        if (this.canEditOverlay()) {
          rawFields = allMetaFields.filter((f) => overlayConfig.allowedTypes.has(f.type));
        } else {
          const viewOrdered = [];
          const viewSeen = new Set();
          const viewMap = new Map(allMetaFields.map((field) => [String(field.code), field]));
          (Array.isArray(viewFieldOrder) ? viewFieldOrder : []).forEach((code) => {
            const normalizedCode = String(code || '').trim();
            if (!normalizedCode || viewSeen.has(normalizedCode)) return;
            const field = viewMap.get(normalizedCode);
            if (!field) return;
            viewOrdered.push(field);
            viewSeen.add(normalizedCode);
          });
          rawFields = viewOrdered.length ? viewOrdered : allMetaFields.slice();
        }
      } else {
        const fieldsRes = await this.postFn('EXCEL_GET_FIELDS', { appId: this.appId });
        if (!fieldsRes?.ok) throw new Error('Failed to get fields');
        rawFields = Array.isArray(fieldsRes.fields)
          ? fieldsRes.fields.filter((f) => overlayConfig.allowedTypes.has(f.type))
          : [];
        if (!rawFields.length && !this.canEditOverlay()) {
          const fallbackCodes = Array.from(new Set(
            (Array.isArray(viewFieldOrder) ? viewFieldOrder : [])
              .map((code) => String(code || '').trim())
              .filter(Boolean)
          ));
          rawFields = fallbackCodes.map((code) => ({
            code,
            label: code,
            type: 'SINGLE_LINE_TEXT',
            required: false,
            choices: []
          }));
        }
      }
      this.fieldsMeta = rawFields.map((field) => ({ ...field }));
      this.fieldsMetaMap.clear();
      this.fieldsMeta.forEach((field) => {
        if (field?.code) this.fieldsMetaMap.set(field.code, field);
      });
      this.syncPermissionServiceFields();

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
        required: Boolean(field.required),
        choices: Array.isArray(field.choices) ? field.choices.map((choice) => String(choice)) : [],
        editable: true,
        width: widthMap.get(field.code) || DEFAULT_COLUMN_WIDTH
      }));
      this.rebuildFieldMaps();
      await this.loadPermissions();
      await this.loadCurrentPageRecords();
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

    normalizeOverlayQuery(baseQuery, pageSize, pageOffset) {
      let raw = String(baseQuery || '').trim();
      raw = raw.replace(/(^|\s+)limit\s+\d+\b/ig, ' ');
      raw = raw.replace(/(^|\s+)offset\s+\d+\b/ig, ' ');
      raw = raw.replace(/\s+/g, ' ').trim();
      const paging = `limit ${pageSize} offset ${pageOffset}`;
      return raw ? `${raw} ${paging}` : paging;
    }

    buildPagedQuery() {
      return this.normalizeOverlayQuery(this.baseQuery || this.query || '', this.pageSize, this.pageOffset);
    }

    syncPermissionServicePendingDeletes() {
      this.permissionService.setPendingDeletes(Array.from(this.pendingDeletes));
    }

    syncPermissionServiceFields() {
      this.permissionService.setFields(this.fieldsMeta);
    }

    getVisibleRows() {
      let visible;
      if (!Array.isArray(this.filteredRowIds)) {
        visible = this.rows.slice();
      } else {
        visible = [];
        this.filteredRowIds.forEach((id) => {
          const row = this.rowMap.get(String(id || '').trim());
          if (row) visible.push(row);
        });
      }
      return this.sortVisibleRows(visible);
    }

    getVisibleRowCount() {
      return this.getVisibleRows().length;
    }

    getVisibleRowAt(index) {
      const rows = this.getVisibleRows();
      return rows[index] || null;
    }

    getVisibleRowIndexById(recordId) {
      const key = String(recordId || '').trim();
      if (!key) return -1;
      const rows = this.getVisibleRows();
      for (let i = 0; i < rows.length; i += 1) {
        if (String(rows[i]?.id || '').trim() === key) return i;
      }
      return -1;
    }

    isSortActive() {
      return Boolean(this.sortState?.fieldCode && this.sortState?.direction);
    }

    getSortDirectionForField(fieldCode) {
      const code = String(fieldCode || '').trim();
      if (!code) return '';
      if (this.sortState.fieldCode !== code) return '';
      if (this.sortState.direction === 'asc' || this.sortState.direction === 'desc') {
        return this.sortState.direction;
      }
      return '';
    }

    toggleSortForField(fieldCode) {
      const code = String(fieldCode || '').trim();
      if (!code) return;
      const current = this.getSortDirectionForField(code);
      if (!current) {
        this.sortState = { fieldCode: code, direction: 'asc' };
      } else if (current === 'asc') {
        this.sortState = { fieldCode: code, direction: 'desc' };
      } else {
        this.sortState = { fieldCode: '', direction: '' };
      }
      this.resetViewAfterSort();
    }

    resetViewAfterSort() {
      this.closeFilterPanel();
      this.closeRadioPicker();
      this.selection = null;
      this.armedCell = null;
      this.editingCell = null;
      this.pendingFocus = null;
      this.pendingEdit = null;
      if (this.bodyScroll) this.bodyScroll.scrollTop = 0;
      this.virtualStart = 0;
      this.virtualEnd = 0;
      this.renderGrid();
      this.updateVirtualRows(true);
      this.updateStats();
      this.updateDirtyBadge();
    }

    isEmptySortValue(value) {
      if (value == null) return true;
      if (Array.isArray(value)) return value.length === 0;
      return String(value).trim() === '';
    }

    parseSortNumber(value) {
      const str = String(value ?? '').trim();
      if (!str) return null;
      const num = Number(str);
      return Number.isNaN(num) ? null : num;
    }

    parseSortDate(value) {
      const str = String(value ?? '').trim();
      if (!str) return null;
      const time = Date.parse(str);
      return Number.isNaN(time) ? null : time;
    }

    compareSortValues(aValue, bValue, field) {
      const aEmpty = this.isEmptySortValue(aValue);
      const bEmpty = this.isEmptySortValue(bValue);
      if (aEmpty && bEmpty) return 0;
      if (aEmpty && !bEmpty) return 1;
      if (!aEmpty && bEmpty) return -1;

      const type = String(field?.type || '');
      if (type === 'NUMBER') {
        const aNum = this.parseSortNumber(aValue);
        const bNum = this.parseSortNumber(bValue);
        if (aNum != null && bNum != null) {
          if (aNum < bNum) return -1;
          if (aNum > bNum) return 1;
          return 0;
        }
      } else if (type === 'DATE') {
        const aDate = this.parseSortDate(aValue);
        const bDate = this.parseSortDate(bValue);
        if (aDate != null && bDate != null) {
          if (aDate < bDate) return -1;
          if (aDate > bDate) return 1;
          return 0;
        }
      }

      const aStr = String(aValue ?? '').toLowerCase();
      const bStr = String(bValue ?? '').toLowerCase();
      return aStr.localeCompare(bStr, undefined, { numeric: true, sensitivity: 'base' });
    }

    sortVisibleRows(rows) {
      if (!Array.isArray(rows) || rows.length <= 1) return Array.isArray(rows) ? rows : [];
      if (!this.isSortActive()) return rows;
      const fieldCode = String(this.sortState.fieldCode || '').trim();
      const direction = this.sortState.direction === 'desc' ? 'desc' : 'asc';
      const field = this.fieldMap.get(fieldCode);
      if (!field) return rows;
      const factor = direction === 'desc' ? -1 : 1;
      const sorted = rows.slice();
      sorted.sort((a, b) => {
        const aValue = a?.values?.[fieldCode];
        const bValue = b?.values?.[fieldCode];
        const base = this.compareSortValues(aValue, bValue, field);
        return base * factor;
      });
      return sorted;
    }

    getFilterTypeForField(field) {
      if (!field) return 'text';
      if (field.type === 'NUMBER') return 'number';
      if (field.type === 'DATE') return 'date';
      if (field.type === 'RADIO_BUTTON' || field.type === 'DROP_DOWN') return 'choice';
      return 'text';
    }

    createDefaultFilter(field) {
      const type = this.getFilterTypeForField(field);
      if (type === 'number' || type === 'date') {
        return { type, mode: 'range', from: '', to: '', value: '', values: [] };
      }
      if (type === 'choice') {
        return { type, mode: 'in', values: [], value: '', from: '', to: '' };
      }
      return { type, mode: 'contains', value: '', from: '', to: '', values: [] };
    }

    isFilterActive(filter) {
      if (!filter || typeof filter !== 'object') return false;
      if (filter.type === 'number' || filter.type === 'date') {
        return String(filter.from || '').trim().length > 0 || String(filter.to || '').trim().length > 0;
      }
      if (filter.type === 'choice') {
        return Array.isArray(filter.values) && filter.values.length > 0;
      }
      return String(filter.value || '').trim().length > 0;
    }

    hasActiveFilters() {
      for (const [, filter] of this.filters.entries()) {
        if (this.isFilterActive(filter)) return true;
      }
      return false;
    }

    getActiveFilterCount() {
      let count = 0;
      this.filters.forEach((filter) => {
        if (this.isFilterActive(filter)) count += 1;
      });
      return count;
    }

    parseFilterNumber(value) {
      const str = String(value ?? '').trim();
      if (!str) return null;
      const num = Number(str);
      return Number.isNaN(num) ? null : num;
    }

    parseFilterDate(value) {
      const str = String(value ?? '').trim();
      if (!str) return null;
      const time = Date.parse(str);
      return Number.isNaN(time) ? null : time;
    }

    matchesFilter(value, filter) {
      if (!this.isFilterActive(filter)) return true;
      if (filter.type === 'number') {
        const current = this.parseFilterNumber(value);
        if (current == null) return false;
        const from = this.parseFilterNumber(filter.from);
        const to = this.parseFilterNumber(filter.to);
        if (from != null && current < from) return false;
        if (to != null && current > to) return false;
        return true;
      }
      if (filter.type === 'date') {
        const current = this.parseFilterDate(value);
        if (current == null) return false;
        const from = this.parseFilterDate(filter.from);
        const to = this.parseFilterDate(filter.to);
        if (from != null && current < from) return false;
        if (to != null && current > to) return false;
        return true;
      }
      if (filter.type === 'choice') {
        const selected = Array.isArray(filter.values) ? filter.values.map((item) => String(item ?? '')) : [];
        if (!selected.length) return true;
        return selected.includes(String(value ?? ''));
      }
      const needle = String(filter.value || '').trim().toLowerCase();
      const hay = String(value ?? '').toLowerCase();
      const mode = String(filter.mode || 'contains');
      if (mode === 'equals') return hay === needle;
      if (mode === 'startsWith') return hay.startsWith(needle);
      return hay.includes(needle);
    }

    recomputeFilteredRowIds() {
      const active = Array.from(this.filters.entries()).filter(([, filter]) => this.isFilterActive(filter));
      if (!active.length) {
        this.filteredRowIds = null;
        return;
      }
      const visible = this.rows.filter((row) => {
        return active.every(([fieldCode, filter]) => {
          const value = row?.values?.[fieldCode] ?? '';
          return this.matchesFilter(value, filter);
        });
      });
      this.filteredRowIds = visible.map((row) => String(row.id || '').trim()).filter(Boolean);
    }

    applyFilters() {
      this.needsFilterReapply = false;
      this.recomputeFilteredRowIds();
      this.resetViewAfterFilter();
    }

    resetViewAfterFilter() {
      this.closeFilterPanel();
      this.closeRadioPicker();
      this.selection = null;
      this.armedCell = null;
      this.editingCell = null;
      this.pendingFocus = null;
      this.pendingEdit = null;
      if (this.bodyScroll) this.bodyScroll.scrollTop = 0;
      this.virtualStart = 0;
      this.virtualEnd = 0;
      this.renderGrid();
      this.updateVirtualRows(true);
      this.updateStats();
      this.updateDirtyBadge();
    }

    clearFieldFilter(fieldCode) {
      const code = String(fieldCode || '').trim();
      if (!code) return;
      this.filters.delete(code);
      this.applyFilters();
    }

    isFieldFilterActive(fieldCode) {
      const filter = this.filters.get(String(fieldCode || '').trim());
      return this.isFilterActive(filter);
    }

    openFilterPanel(field, anchorButton) {
      if (!this.root || !field || !anchorButton) return;
      const fieldCode = String(field.code || '').trim();
      if (!fieldCode) return;
      this.closeFilterPanel();

      const existing = this.filters.get(fieldCode);
      const working = existing
        ? { ...existing, values: Array.isArray(existing.values) ? existing.values.slice() : [] }
        : this.createDefaultFilter(field);

      const panel = document.createElement('div');
      panel.className = 'pb-overlay__filter-panel';
      panel.setAttribute('role', 'dialog');
      panel.addEventListener('click', (event) => event.stopPropagation());

      const title = document.createElement('div');
      title.className = 'pb-overlay__filter-title';
      title.textContent = `${resolveText(this.language, 'filterTitle')}: ${field.label || field.code}`;
      panel.appendChild(title);

      const body = document.createElement('div');
      body.className = 'pb-overlay__filter-body';
      panel.appendChild(body);

      let modeSelect = null;
      let valueInput = null;
      let fromInput = null;
      let toInput = null;
      const choiceInputs = [];

      if (working.type === 'text') {
        modeSelect = document.createElement('select');
        modeSelect.className = 'pb-overlay__filter-select';
        [
          ['contains', resolveText(this.language, 'filterContains')],
          ['equals', resolveText(this.language, 'filterEquals')],
          ['startsWith', resolveText(this.language, 'filterStartsWith')]
        ].forEach(([value, label]) => {
          const option = document.createElement('option');
          option.value = value;
          option.textContent = label;
          if (working.mode === value) option.selected = true;
          modeSelect.appendChild(option);
        });
        valueInput = document.createElement('input');
        valueInput.type = 'text';
        valueInput.className = 'pb-overlay__filter-input';
        valueInput.value = String(working.value || '');
        body.appendChild(modeSelect);
        body.appendChild(valueInput);
      } else if (working.type === 'number') {
        fromInput = document.createElement('input');
        fromInput.type = 'text';
        fromInput.inputMode = 'decimal';
        fromInput.className = 'pb-overlay__filter-input';
        fromInput.placeholder = resolveText(this.language, 'filterFrom');
        fromInput.value = String(working.from || '');
        toInput = document.createElement('input');
        toInput.type = 'text';
        toInput.inputMode = 'decimal';
        toInput.className = 'pb-overlay__filter-input';
        toInput.placeholder = resolveText(this.language, 'filterTo');
        toInput.value = String(working.to || '');
        body.appendChild(fromInput);
        body.appendChild(toInput);
      } else if (working.type === 'date') {
        fromInput = document.createElement('input');
        fromInput.type = 'date';
        fromInput.className = 'pb-overlay__filter-input';
        fromInput.value = String(working.from || '');
        toInput = document.createElement('input');
        toInput.type = 'date';
        toInput.className = 'pb-overlay__filter-input';
        toInput.value = String(working.to || '');
        body.appendChild(fromInput);
        body.appendChild(toInput);
      } else if (working.type === 'choice') {
        const list = document.createElement('div');
        list.className = 'pb-overlay__filter-choices';
        const selected = new Set(Array.isArray(working.values) ? working.values.map((item) => String(item)) : []);
        const choices = Array.isArray(field.choices) ? field.choices : [];
        choices.forEach((choice) => {
          const label = document.createElement('label');
          label.className = 'pb-overlay__filter-choice';
          const check = document.createElement('input');
          check.type = 'checkbox';
          check.value = String(choice);
          check.checked = selected.has(String(choice));
          choiceInputs.push(check);
          const text = document.createElement('span');
          text.textContent = String(choice);
          label.appendChild(check);
          label.appendChild(text);
          list.appendChild(label);
        });
        body.appendChild(list);
      }

      const actions = document.createElement('div');
      actions.className = 'pb-overlay__filter-actions';
      const clearBtn = document.createElement('button');
      clearBtn.type = 'button';
      clearBtn.className = 'pb-overlay__btn';
      clearBtn.textContent = resolveText(this.language, 'filterClear');
      const applyBtn = document.createElement('button');
      applyBtn.type = 'button';
      applyBtn.className = 'pb-overlay__btn pb-overlay__btn--primary';
      applyBtn.textContent = resolveText(this.language, 'filterApply');
      actions.appendChild(clearBtn);
      actions.appendChild(applyBtn);
      panel.appendChild(actions);

      clearBtn.addEventListener('click', () => {
        this.clearFieldFilter(fieldCode);
      });
      applyBtn.addEventListener('click', () => {
        const next = this.createDefaultFilter(field);
        if (working.type === 'text') {
          next.mode = modeSelect ? modeSelect.value : 'contains';
          next.value = valueInput ? String(valueInput.value || '').trim() : '';
        } else if (working.type === 'number' || working.type === 'date') {
          next.from = fromInput ? String(fromInput.value || '').trim() : '';
          next.to = toInput ? String(toInput.value || '').trim() : '';
        } else if (working.type === 'choice') {
          next.values = choiceInputs.filter((input) => input.checked).map((input) => String(input.value || ''));
        }
        if (this.isFilterActive(next)) {
          this.filters.set(fieldCode, next);
        } else {
          this.filters.delete(fieldCode);
        }
        this.applyFilters();
      });

      this.root.appendChild(panel);
      const rootRect = this.root.getBoundingClientRect();
      const anchorRect = anchorButton.getBoundingClientRect();
      const panelRect = panel.getBoundingClientRect();
      const margin = 12;
      let left = anchorRect.left - rootRect.left - panelRect.width + anchorRect.width;
      if (left < margin) left = margin;
      const maxLeft = Math.max(margin, rootRect.width - panelRect.width - margin);
      if (left > maxLeft) left = maxLeft;
      let top = anchorRect.bottom - rootRect.top + 6;
      const maxTop = Math.max(margin, rootRect.height - panelRect.height - margin);
      if (top > maxTop) {
        top = anchorRect.top - rootRect.top - panelRect.height - 6;
      }
      if (top < margin) top = margin;
      panel.style.left = `${left}px`;
      panel.style.top = `${top}px`;

      this.filterPanel = { panel, fieldCode, button: anchorButton };
      document.addEventListener('mousedown', this.boundFilterPanelOutsideClick, true);
      requestAnimationFrame(() => {
        const focusable = panel.querySelector('input,select,button');
        if (focusable instanceof HTMLElement) focusable.focus();
      });
    }

    closeFilterPanel() {
      if (!this.filterPanel) return;
      const { panel } = this.filterPanel;
      if (panel?.parentElement) panel.remove();
      this.filterPanel = null;
      document.removeEventListener('mousedown', this.boundFilterPanelOutsideClick, true);
    }

    updatePagerUi() {
      if (this.pageLabelElement) {
        const total = Math.max(0, Number(this.totalCount || 0));
        if (!total) {
          this.pageLabelElement.textContent = '0-0 / 0';
        } else {
          const start = this.pageOffset + 1;
          const visible = this.rows.length || this.pageSize;
          const end = Math.min(this.pageOffset + visible, total);
          this.pageLabelElement.textContent = `${start}-${end} / ${total}`;
        }
      }
      if (this.prevPageButton) {
        this.prevPageButton.disabled = this.saving || this.paging || this.pageOffset <= 0;
      }
      if (this.nextPageButton) {
        const total = Math.max(0, Number(this.totalCount || 0));
        this.nextPageButton.disabled = this.saving || this.paging || total <= 0 || (this.pageOffset + this.pageSize >= total);
      }
    }

    async loadCurrentPageRecords() {
      const fieldCodes = this.fields.map((f) => f.code);
      const pagedQuery = this.buildPagedQuery();
      console.debug('[excel-query]', {
        baseQuery: this.baseQuery || this.query || '',
        pagedQuery
      });
      const recordsRes = await this.postFn('EXCEL_GET_RECORDS', {
        appId: this.appId,
        fields: fieldCodes,
        query: pagedQuery
      });
      if (!recordsRes?.ok) {
        const errorCode = recordsRes?.errorCode || '';
        if (errorCode === 'GAIA_IL02' || /GAIA_IL02/.test(recordsRes?.error || '')) {
          throw new Error(resolveText(this.language, 'errorUnsupportedFilter'));
        }
        throw new Error(recordsRes?.error || 'Failed to get records');
      }
      this.totalCount = Number(recordsRes.totalCount || 0);
      const records = Array.isArray(recordsRes.records) ? recordsRes.records : [];
      this.rows = records.map((record) => this.transformRecord(record));
      console.log('[overlay] records loaded', records.length);
      this.rebuildRowMaps();
      this.recomputeFilteredRowIds();
      await this.loadEffectiveRecordAcl();
      this.syncPermissionServicePendingDeletes();
      this.updatePermissionUiState();
      this.updatePagerUi();
    }

    discardUnsavedChangesForPaging() {
      this.exitEditMode();
      this.closeRadioPicker();
      this.diff.clear();
      this.invalidCells.clear();
      this.invalidValueCache.clear();
      this.pendingDeletes.clear();
      this.syncPermissionServicePendingDeletes();
      this.updateDirtyBadge();
      this.updateStats();
    }

    confirmPageMoveIfNeeded() {
      if (!this.hasUnsavedChanges()) return true;
      const ok = window.confirm(resolveText(this.language, 'confirmPageMoveUnsaved'));
      if (!ok) return false;
      this.discardUnsavedChangesForPaging();
      return true;
    }

    async reloadPage() {
      if (this.paging || !this.isOpen) return;
      this.paging = true;
      this.updatePagerUi();
      try {
        this.showLoading(resolveText(this.language, 'loading'));
        await this.loadCurrentPageRecords();
        if (this.bodyScroll) this.bodyScroll.scrollTop = 0;
        this.renderGrid();
        this.updateDirtyBadge();
        this.updateStats();
        this.focusFirstCell();
      } catch (error) {
        this.notify(resolveText(this.language, 'loadFailed'));
        console.error('[kintone-excel-overlay] failed to reload page', error);
      } finally {
        this.hideLoading();
        this.paging = false;
        this.updatePermissionUiState();
        this.updatePagerUi();
      }
    }

    async prevPage() {
      if (this.pageOffset <= 0 || this.paging) return;
      if (!this.confirmPageMoveIfNeeded()) return;
      this.pageOffset = Math.max(0, this.pageOffset - this.pageSize);
      await this.reloadPage();
    }

    async nextPage() {
      if (this.paging) return;
      if (this.pageOffset + this.pageSize >= this.totalCount) return;
      if (!this.confirmPageMoveIfNeeded()) return;
      this.pageOffset += this.pageSize;
      await this.reloadPage();
    }

    rebuildFieldMaps() {
      this.fieldMap.clear();
      this.fieldIndexMap.clear();
      this.fields.forEach((field, index) => {
        this.fieldMap.set(field.code, field);
        this.fieldIndexMap.set(field.code, index);
      });
    }

    rebuildRowMaps() {
      this.rowMap.clear();
      this.rowIndexMap.clear();
      this.revisionMap.clear();
      this.rows.forEach((row, index) => {
        const id = String(row?.id || '').trim();
        if (!id) return;
        this.rowMap.set(id, row);
        this.rowIndexMap.set(id, index);
        if (row?.revision !== undefined) {
          this.revisionMap.set(id, row.revision);
        }
      });
    }

    isChoiceField(field) {
      const type = field?.type || '';
      return type === 'RADIO_BUTTON' || type === 'DROP_DOWN';
    }

    isSubtableField(field) {
      return (field?.type || '') === 'SUBTABLE';
    }

    shouldIncludeField(field) {
      if (!field || !field.code || !field.type) return false;
      return true;
    }

    getCellPermissionInfo(recordId, fieldCode) {
      const key = String(recordId || '').trim();
      const code = String(fieldCode || '').trim();
      if (!key || !code) return { editable: false, reason: 'unknown_field' };
      return this.permissionService.getCellPermission(key, code);
    }

    async loadEffectiveRecordAcl() {
      this.permissionService.clearRecordAcl();
      this.aclStatus = { loaded: false, failed: false };
      const appId = String(this.appId || '').trim();
      if (!appId) return;
      const ids = this.rows
        .map((row) => String(row?.id || '').trim())
        .filter((id) => id && !this.isNewRowId(id));
      if (!ids.length) return;

      const chunkSize = 100;
      let succeeded = false;
      let failed = false;
      const allRights = [];
      for (let i = 0; i < ids.length; i += chunkSize) {
        const chunk = ids.slice(i, i + chunkSize);
        try {
          const res = await this.postFn('EXCEL_EVALUATE_RECORD_ACL', { appId, ids: chunk });
          if (!res?.ok) {
            failed = true;
            console.warn('[kintone-excel-overlay] failed to evaluate record acl', {
              appId,
              ids: chunk,
              error: res?.error || 'unknown error'
            });
            continue;
          }
          succeeded = true;
          const rights = Array.isArray(res.rights)
            ? res.rights
            : Array.isArray(res.records)
              ? res.records
              : [];
          allRights.push(...rights);
        } catch (error) {
          failed = true;
          console.warn('[kintone-excel-overlay] record acl evaluation request failed', {
            appId,
            ids: chunk,
            error: String(error?.message || error)
          });
        }
      }
      if (allRights.length) {
        this.permissionService.refreshRecordAcl(allRights);
      }
      this.aclStatus = {
        loaded: succeeded,
        failed: !succeeded && failed
      };
      this.syncPermissionServicePendingDeletes();
      console.debug('[excel-acl-status]', this.aclStatus);
    }

    getRowPermissionStatusText(recordId) {
      const editable = this.permissionService.canEditRecord(recordId);
      const deletable = this.permissionService.canDeleteRecord(recordId);
      const parts = [];
      if (editable) {
        parts.push(resolveText(this.language, 'permEditable'));
      } else {
        parts.push(resolveText(this.language, 'permViewOnly'));
      }
      if (deletable) {
        parts.push(resolveText(this.language, 'permDeletable'));
      }
      return parts.join(' / ');
    }

    async loadPermissions() {
      // Keep default behavior: do not call app-level ACL APIs in normal flow.
      // Effective per-record restrictions are evaluated via EXCEL_EVALUATE_RECORD_ACL.
      this.syncPermissionServiceFields();
      this.permissionService.setAppPermission({
        editable: true,
        addable: true,
        deletable: true,
        source: 'local'
      });
      this.syncPermissionServicePendingDeletes();
      console.debug('[excel-permission-state]', {
        aclStatus: this.aclStatus,
        rows: this.rows.length
      });
    }

    updatePermissionUiState() {
      const canEdit = this.canEditOverlay();
      const canSave = this.canSaveOverlay();
      if (this.addRowButton) {
        const allowed = canEdit && this.permissionService.canAddRow();
        this.addRowButton.disabled = this.saving || !allowed;
        this.addRowButton.title = allowed
          ? ''
          : (canEdit ? resolveText(this.language, 'permNoAdd') : resolveText(this.language, 'toastViewOnlyBlocked'));
      }
      if (this.saveButton && !this.saving) {
        let dirtyCount = 0;
        this.diff.forEach((entry) => {
          dirtyCount += Object.keys(entry).length;
        });
        dirtyCount += this.pendingDeletes.size;
        this.saveButton.disabled = !canSave || dirtyCount === 0;
        this.saveButton.title = canSave ? '' : resolveText(this.language, 'toastViewOnlyBlocked');
      }
      if (this.permissionWarningEl) {
        const showWarning = this.aclStatus.failed === true && this.aclStatus.loaded === false;
        this.permissionWarningEl.classList.toggle('pb-overlay__perm-warning--hidden', !showWarning);
      }
      this.updateHistoryButtons();
      this.updatePagerUi();
    }

    cloneFieldValue(field, value) {
      if (this.isSubtableField(field)) {
        if (Array.isArray(value)) {
          return value.map((row) => ({
            id: row?.id || '',
            value: row && typeof row.value === 'object' ? JSON.parse(JSON.stringify(row.value)) : {}
          }));
        }
        return [];
      }
      return value === undefined || value === null ? '' : String(value);
    }

    formatCellDisplayValue(field, value) {
      if (this.isSubtableField(field)) {
        const count = Array.isArray(value) ? value.length : 0;
        return count ? resolveText(this.language, 'subtableRows', count) : '';
      }
      return value === undefined || value === null ? '' : String(value);
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
      this.armedCell = null;
      this.editingCell = null;
      this.pendingFocus = null;
      this.pendingEdit = null;
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
        const baseValue = this.cloneFieldValue(field, raw);
        values[field.code] = baseValue;
        original[field.code] = this.cloneFieldValue(field, baseValue);
      });
      return { id, revision, values, original, isNew: false };
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
      this.closeFilterPanel();
      this.closeRadioPicker();
      this.inputsByRow.clear();
      if (this.columnHeaderScroll) {
        this.columnHeaderScroll.style.transform = 'translateX(0px)';
      }
      const headRow = this.columnHeaderScroll.querySelector('.pb-overlay__col-row');
      headRow.textContent = '';
      const rowWrap = this.rowHeaderScroll.querySelector('.pb-overlay__row-wrap');
      const table = this.bodyScroll.querySelector('.pb-overlay__table');
      rowWrap.textContent = '';
      table.textContent = '';

      const visibleRows = this.getVisibleRows();
      const totalHeight = Math.max(0, visibleRows.length * this.rowHeight);
      rowWrap.style.height = `${totalHeight}px`;
      table.style.height = `${totalHeight}px`;

      this.fields.forEach((field, index) => {
        const headerCell = document.createElement('div');
        headerCell.className = 'pb-overlay__col-header';
        headerCell.style.width = `${field.width || DEFAULT_COLUMN_WIDTH}px`;
        headerCell.style.minWidth = `${field.width || DEFAULT_COLUMN_WIDTH}px`;
        const headerMain = document.createElement('div');
        headerMain.className = 'pb-overlay__col-header-main';
        headerMain.style.cursor = 'pointer';
        const codeSpan = document.createElement('span');
        codeSpan.className = 'pb-overlay__col-code';
        codeSpan.textContent = columnLabel(index);
        const labelWrap = document.createElement('span');
        labelWrap.className = 'pb-overlay__col-label-wrap';
        const labelSpan = document.createElement('span');
        labelSpan.className = 'pb-overlay__col-label';
        labelSpan.textContent = field.label || field.code;
        const sortIndicator = document.createElement('span');
        sortIndicator.className = 'pb-overlay__col-sort-indicator';
        const sortDirection = this.getSortDirectionForField(field.code);
        sortIndicator.textContent = sortDirection === 'asc' ? '▲' : (sortDirection === 'desc' ? '▼' : '');
        labelWrap.appendChild(labelSpan);
        labelWrap.appendChild(sortIndicator);
        headerMain.appendChild(codeSpan);
        headerMain.appendChild(labelWrap);
        const filterBtn = document.createElement('button');
        filterBtn.type = 'button';
        filterBtn.className = 'pb-overlay__filter-btn';
        filterBtn.textContent = '▼';
        const filterLabel = resolveText(this.language, 'filterButton');
        filterBtn.title = filterLabel;
        filterBtn.setAttribute('aria-label', `${filterLabel}: ${field.label || field.code}`);
        if (this.isFieldFilterActive(field.code)) {
          filterBtn.classList.add('pb-overlay__filter-btn--active');
        }
        filterBtn.addEventListener('click', (event) => {
          event.preventDefault();
          event.stopPropagation();
          this.openFilterPanel(field, filterBtn);
        });
        headerMain.addEventListener('click', (event) => {
          event.preventDefault();
          this.toggleSortForField(field.code);
        });
        headerCell.appendChild(headerMain);
        headerCell.appendChild(filterBtn);
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
      const visibleRows = this.getVisibleRows();
      const totalRows = visibleRows.length;
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
      this.rowHeaderContainer.style.height = `${totalRows * this.rowHeight}px`;
      this.tableContainer.textContent = '';
      const totalWidth = this.fields.length * 160;
      this.tableContainer.style.width = `${totalWidth}px`;

      const headerFrag = document.createDocumentFragment();
      const tableFrag = document.createDocumentFragment();

      for (let rowIndex = start; rowIndex < end; rowIndex += 1) {
        const row = visibleRows[rowIndex];
        if (!row) continue;
        const recordId = String(row.id || '').trim();
        const pendingDelete = this.pendingDeletes.has(recordId);
        const rowEditable = this.isOverlayEditable() && this.permissionService.canEditRecord(recordId);
        const rowEditableVisual = this.isOverlayViewOnly() ? true : rowEditable;
        const rowDeletable = this.isOverlayEditable() && this.permissionService.canDeleteRecord(recordId);
        const top = rowIndex * this.rowHeight;

        const header = document.createElement('div');
        header.className = 'pb-overlay__row-header';
        header.style.position = 'absolute';
        header.style.top = `${top}px`;
        header.style.left = '0';
        if (!rowEditableVisual) header.classList.add('pb-overlay__row-header--readonly');
        if (pendingDelete) header.classList.add('pb-overlay__row-header--pending-delete');
        if (pendingDelete) {
          header.title = resolveText(this.language, 'rowDeletePendingTitle');
        } else if (!rowEditableVisual) {
          header.title = resolveText(this.language, 'permNoEdit');
        } else {
          header.removeAttribute('title');
        }

        const indexEl = document.createElement('span');
        indexEl.className = 'pb-overlay__row-index';
        indexEl.textContent = String(rowIndex + 1);
        header.appendChild(indexEl);

        const sepEl = document.createElement('span');
        sepEl.className = 'pb-overlay__row-sep';
        sepEl.setAttribute('aria-hidden', 'true');
        header.appendChild(sepEl);

        const deleteBtn = document.createElement('button');
        deleteBtn.type = 'button';
        deleteBtn.className = 'pb-overlay__row-delete pb-overlay__delete-btn';
        deleteBtn.textContent = '🗑';
        if (pendingDelete) {
          deleteBtn.classList.add('is-pending');
        }
        const isNew = this.isNewRowId(recordId);
        if (isNew) {
          deleteBtn.title = resolveText(this.language, 'rowDeleteNewUndo');
          deleteBtn.setAttribute('aria-label', resolveText(this.language, 'rowDeleteNewUndo'));
        } else if (pendingDelete) {
          deleteBtn.title = resolveText(this.language, 'rowDeletePendingTitle');
          deleteBtn.setAttribute('aria-label', resolveText(this.language, 'rowDeletePendingTitle'));
        } else {
          deleteBtn.title = rowDeletable
            ? resolveText(this.language, 'rowDeleteToggle')
            : resolveText(this.language, 'permNoDelete');
          deleteBtn.setAttribute('aria-label', deleteBtn.title);
        }
        if (!isNew && !rowDeletable) {
          deleteBtn.disabled = true;
        }
        deleteBtn.addEventListener('click', (event) => {
          event.preventDefault();
          event.stopPropagation();
          this.togglePendingDelete(recordId);
        });
        header.appendChild(deleteBtn);

        const permEl = document.createElement('span');
        permEl.className = 'pb-overlay__row-perm';
        permEl.textContent = rowEditableVisual ? '✎' : '👁';
        permEl.title = this.getRowPermissionStatusText(recordId);
        header.appendChild(permEl);
        headerFrag.appendChild(header);

        const rowLine = document.createElement('div');
        rowLine.className = 'pb-overlay__row';
        rowLine.style.position = 'absolute';
        rowLine.style.top = `${top}px`;
        rowLine.style.left = '0';
        rowLine.style.width = 'max-content';
        if (!rowEditableVisual) rowLine.classList.add('pb-overlay__row--readonly');
        if (pendingDelete) rowLine.classList.add('pb-overlay__row--pending-delete');

        const inputsRow = [];

        this.fields.forEach((field, colIndex) => {
          const cell = document.createElement('div');
          cell.className = 'pb-overlay__cell';
          cell.dataset.fieldCode = String(field.code || '');
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
      if (this.isChoiceField(field)) {
        input.placeholder = field.choices?.[0] || '';
      }
      if (this.isSubtableField(field)) {
        input.placeholder = resolveText(this.language, 'btnSubtableEdit');
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
        this.handleChoiceInputClick(input);
      });
      return input;
    }

    bindInputToRow(input, field, row, rowIndex, colIndex) {
      input.dataset.recordId = row.id;
      input.dataset.fieldCode = field.code;
      input.dataset.rowIndex = String(rowIndex);
      input.dataset.colIndex = String(colIndex);
      input.dataset.originalValue = this.formatCellDisplayValue(field, row.original[field.code]);
      const permission = this.getCellPermissionInfo(row.id, field.code);
      const editable = this.isOverlayEditable() && permission.editable;
      input.dataset.editable = editable ? 'true' : 'false';
      input.disabled = false;
      const editableVisual = this.isOverlayViewOnly() ? true : editable;
      const isPending = Boolean(
        this.pendingEdit
        && this.pendingEdit.rowIndex === rowIndex
        && this.pendingEdit.colIndex === colIndex
      );
      const editing = (this.isEditingCell(rowIndex, colIndex) || isPending) && editable;
      const lookupReadonly = permission.reason === 'lookup_readonly';
      const isDropdown = field?.type === 'DROP_DOWN';
      const isRadio = field?.type === 'RADIO_BUTTON';
      const isLookupKey = field?.type === 'SINGLE_LINE_TEXT' && Boolean(field?.lookup);
      const fieldDeniedByAcl = permission.reason === 'field_readonly';
      this.setInputEditingVisual(input, editing);
      input.classList.toggle('pb-overlay__input--readonly', !editableVisual);
      input.classList.toggle('pb-overlay__input--subtable', this.isSubtableField(field));
      if (lookupReadonly) {
        input.title = resolveText(this.language, 'lookupAutoReadonly');
      } else if (!this.isSubtableField(field) && !editableVisual) {
        input.title = resolveText(this.language, fieldDeniedByAcl ? 'permNoFieldEdit' : 'permNoEdit');
      } else if (!this.isSubtableField(field)) {
        input.removeAttribute('title');
      }
      if (input.parentElement) {
        input.parentElement.classList.toggle('pb-overlay__cell--lookup-auto', lookupReadonly);
        input.parentElement.classList.toggle('pb-overlay__cell--readonly', !editableVisual);
        input.parentElement.classList.toggle('pb-overlay__cell--dropdown', isDropdown);
        input.parentElement.classList.toggle('pb-overlay__cell--radio', isRadio);
        input.parentElement.classList.toggle('pb-overlay__cell--lookup', isLookupKey);
      }
    }

    applyCellVisualState(input, row, fieldCode) {
      const field = this.fieldMap.get(fieldCode);
      const value = row?.values ? row.values[fieldCode] ?? '' : '';
      input.value = this.formatCellDisplayValue(field, value);
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
        if (cell.dataset.errorReason) {
          delete cell.dataset.errorReason;
        }
      }
      const rowIndex = Number(input.dataset.rowIndex || '-1');
      const colIndex = Number(input.dataset.colIndex || '-1');
      if (this.isCellSelected(rowIndex, colIndex)) {
        cell.classList.add('pb-overlay__cell--selected');
      } else {
        cell.classList.remove('pb-overlay__cell--selected');
      }
      if (this.isCellArmed(rowIndex, colIndex)) {
        cell.classList.add('pb-overlay__cell--armed');
      } else {
        cell.classList.remove('pb-overlay__cell--armed');
      }
      if (this.isCellActive(rowIndex, colIndex)) {
        cell.classList.add('pb-overlay__cell--active');
      } else {
        cell.classList.remove('pb-overlay__cell--active');
      }
      const permission = this.getCellPermissionInfo(row.id, fieldCode);
      const editable = this.isOverlayEditable() && permission.editable;
      const editableVisual = this.isOverlayViewOnly() ? true : editable;
      const isPending = Boolean(
        this.pendingEdit
        && this.pendingEdit.rowIndex === rowIndex
        && this.pendingEdit.colIndex === colIndex
      );
      const editing = (this.isEditingCell(rowIndex, colIndex) || isPending) && editable;
      const lookupReadonly = permission.reason === 'lookup_readonly';
      const isDropdown = field?.type === 'DROP_DOWN';
      const isRadio = field?.type === 'RADIO_BUTTON';
      const isLookupKey = field?.type === 'SINGLE_LINE_TEXT' && Boolean(field?.lookup);
      const fieldDeniedByAcl = permission.reason === 'field_readonly';
      this.setInputEditingVisual(input, editing);
      input.classList.toggle('pb-overlay__input--readonly', !editableVisual);
      input.classList.toggle('pb-overlay__input--subtable', this.isSubtableField(field));
      if (lookupReadonly) {
        input.title = resolveText(this.language, 'lookupAutoReadonly');
      } else if (!this.isSubtableField(field) && !editableVisual) {
        input.title = resolveText(this.language, fieldDeniedByAcl ? 'permNoFieldEdit' : 'permNoEdit');
      } else if (!this.isSubtableField(field)) {
        input.removeAttribute('title');
      }
      cell.classList.toggle('pb-overlay__cell--readonly', !editableVisual);
      cell.classList.toggle('pb-overlay__cell--lookup-auto', lookupReadonly);
      cell.classList.toggle('pb-overlay__cell--dropdown', isDropdown);
      cell.classList.toggle('pb-overlay__cell--radio', isRadio);
      cell.classList.toggle('pb-overlay__cell--lookup', isLookupKey);
      cell.classList.toggle('pb-overlay__cell--editing', editing);
    }

    getInput(rowIndex, colIndex) {
      const rowInputs = this.inputsByRow.get(rowIndex);
      if (!rowInputs) return null;
      return rowInputs[colIndex] || null;
    }

    focusCell(rowIndex, colIndex, options = {}) {
      const ensureVisible = options?.ensureVisible !== false;
      const maxRow = Math.max(0, this.getVisibleRowCount() - 1);
      const maxCol = Math.max(0, this.fields.length - 1);
      if (maxRow < 0 || maxCol < 0) return;
      const r = Math.max(0, Math.min(maxRow, rowIndex));
      const c = Math.max(0, Math.min(maxCol, colIndex));
      if (ensureVisible) {
        this.ensureCellVisible(r, c);
      }
      this.pendingFocus = { rowIndex: r, colIndex: c };
      if (r < this.virtualStart || r >= this.virtualEnd) {
        if (this.bodyScroll) {
          this.updateVirtualRows(true);
        }
        return;
      }
      this.restorePendingFocus();
    }

    getColumnWidth(colIndex) {
      const field = this.fields[colIndex];
      return Number(field?.width || DEFAULT_COLUMN_WIDTH);
    }

    getColumnLeft(colIndex) {
      let left = 0;
      for (let i = 0; i < colIndex; i += 1) {
        left += this.getColumnWidth(i);
      }
      return left;
    }

    ensureCellVisible(rowIndex, colIndex) {
      if (!this.bodyScroll) return;
      const rowTop = rowIndex * this.rowHeight;
      const rowBottom = rowTop + this.rowHeight;
      const viewTop = this.bodyScroll.scrollTop;
      const viewBottom = viewTop + this.bodyScroll.clientHeight;
      if (rowTop < viewTop) {
        this.bodyScroll.scrollTop = rowTop;
      } else if (rowBottom > viewBottom) {
        this.bodyScroll.scrollTop = rowBottom - this.bodyScroll.clientHeight;
      }

      const colLeft = this.getColumnLeft(colIndex);
      const colRight = colLeft + this.getColumnWidth(colIndex);
      const viewLeft = this.bodyScroll.scrollLeft;
      const viewRight = viewLeft + this.bodyScroll.clientWidth;
      if (colLeft < viewLeft) {
        this.bodyScroll.scrollLeft = colLeft;
      } else if (colRight > viewRight) {
        this.bodyScroll.scrollLeft = colRight - this.bodyScroll.clientWidth;
      }
    }

    getGridScrollElement() {
      return this.bodyScroll || null;
    }

    syncSelectionScroll(options = {}) {
      const scroller = this.getGridScrollElement();
      if (!scroller) return;
      const rowIndex = Number(options.rowIndex ?? this.selection?.endRow ?? 0);
      const colIndex = Number(options.colIndex ?? this.selection?.endCol ?? 0);
      const rowHeight = Number(this.rowHeight || 32);
      const verticalMargin = Number.isFinite(options.verticalMargin) ? Number(options.verticalMargin) : rowHeight;
      const horizontalMargin = Number.isFinite(options.horizontalMargin) ? Number(options.horizontalMargin) : 24;

      const rowTop = rowIndex * rowHeight;
      const rowBottom = rowTop + rowHeight;
      const viewTop = scroller.scrollTop;
      const viewBottom = viewTop + scroller.clientHeight;

      let nextScrollTop = viewTop;
      if (rowTop < viewTop + verticalMargin) {
        nextScrollTop = Math.max(0, rowTop - verticalMargin);
      } else if (rowBottom > viewBottom - verticalMargin) {
        nextScrollTop = Math.max(0, rowBottom - scroller.clientHeight + verticalMargin);
      }

      const colLeft = this.getColumnLeft(colIndex);
      const colWidth = this.getColumnWidth(colIndex);
      const colRight = colLeft + colWidth;
      const viewLeft = scroller.scrollLeft;
      const viewRight = viewLeft + scroller.clientWidth;

      let nextScrollLeft = viewLeft;
      if (colLeft < viewLeft + horizontalMargin) {
        nextScrollLeft = Math.max(0, colLeft - horizontalMargin);
      } else if (colRight > viewRight - horizontalMargin) {
        nextScrollLeft = Math.max(0, colRight - scroller.clientWidth + horizontalMargin);
      }

      let changed = false;
      if (nextScrollTop !== scroller.scrollTop) {
        scroller.scrollTop = nextScrollTop;
        changed = true;
      }
      if (nextScrollLeft !== scroller.scrollLeft) {
        scroller.scrollLeft = nextScrollLeft;
        changed = true;
      }
      if (changed) {
        this.syncScroll();
      }
    }

    restorePendingFocus() {
      if (!this.pendingFocus) return;
      const { rowIndex, colIndex } = this.pendingFocus;
      const input = this.getInput(rowIndex, colIndex);
      if (!input) return;
      this.pendingFocus = null;
      requestAnimationFrame(() => {
        input.focus();
        const shouldForceEdit = Boolean(
          this.pendingEdit
          && this.pendingEdit.rowIndex === rowIndex
          && this.pendingEdit.colIndex === colIndex
        );
        if (shouldForceEdit) {
          this.forceApplyEditState(input, rowIndex, colIndex);
          this.pendingEdit = null;
        }
        const editing = this.isEditingCell(rowIndex, colIndex);
        const rowInputs = this.inputsByRow.get(rowIndex) || [];
        rowInputs.forEach((candidate) => {
          if (candidate && candidate !== input && candidate.classList?.contains('pb-overlay__input--editing')) {
            this.setInputEditingVisual(candidate, false);
          }
        });
        this.setInputEditingVisual(input, editing);
        const shouldSelectAll = editing && this.editingCell && this.editingCell.selectAll;
        this.applyCaretPlacement(input, shouldSelectAll);
      });
    }

    setInputEditingVisual(input, isEditing) {
      if (!input) return;
      const editable = input.dataset?.editable !== 'false';
      const nextEditing = Boolean(isEditing) && editable && !input.disabled;
      input.readOnly = !nextEditing;
      input.classList.toggle('pb-overlay__input--editing', nextEditing);
      if (input.parentElement) {
        input.parentElement.classList.toggle('pb-overlay__cell--editing', nextEditing);
      }
    }

    forceApplyEditState(input, rowIndex, colIndex) {
      const isSame = this.editingCell
        && this.editingCell.rowIndex === rowIndex
        && this.editingCell.colIndex === colIndex;
      this.editingCell = isSame
        ? { ...this.editingCell, selectAll: false }
        : this.createEditingState(rowIndex, colIndex, false);
      this.setInputEditingVisual(input, true);
      const cell = input.parentElement;
      if (cell) cell.classList.add('pb-overlay__cell--selected');
    }

    canSelectTextInput(input) {
      if (!input) return false;
      const type = (input.type || '').toLowerCase();
      return type === 'text' || type === 'search' || type === 'tel' || type === 'url' || type === 'email' || type === 'password';
    }

    handleInputBlur(event, input) {
      if (this.navigating) return;
      if (!this.editingCell) return;
      if (this.radioPicker && this.radioPicker.input === input) {
        const next = event.relatedTarget;
        if (next && this.radioPicker.panel?.contains(next)) {
          return;
        }
      }
      this.exitEditMode();
      if (this.needsFilterReapply && this.hasActiveFilters()) {
        this.needsFilterReapply = false;
        this.applyFilters();
      }
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
      if (!field || !this.isChoiceField(field)) return;
      const rawChoices = Array.isArray(field.choices) ? field.choices : [];
      const choices = Array.from(new Set(rawChoices.map((choice) => String(choice || '')).filter((choice) => choice)));
      if (!choices.length) return;

      this.closeRadioPicker();
      const panel = document.createElement('div');
      panel.className = 'pb-overlay__choice-panel';
      panel.setAttribute('role', 'listbox');
      panel.setAttribute('aria-label', field.label || field.code || 'choices');
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

    createSubtableDefaultValue(type) {
      if (type === 'NUMBER') return '';
      if (type === 'DATE') return '';
      return '';
    }

    openSubtableEditor(rowIndex, colIndex, anchorInput = null) {
      if (this.saving) return;
      if (!this.root) return;
      const row = this.getVisibleRowAt(rowIndex);
      const field = this.fields[colIndex];
      if (!row || !field || !this.isSubtableField(field)) return;
      const readOnlyMode = !this.isOverlayEditable() || !this.permissionService.canEditRecord(row.id);
      this.closeSubtableEditor(false);
      this.exitEditMode();

      const childFields = Array.isArray(field.subtable?.fields)
        ? field.subtable.fields.filter((child) => (
          child
          && child.code
          && ['SINGLE_LINE_TEXT', 'NUMBER', 'DATE', 'RADIO_BUTTON', 'DROP_DOWN'].includes(child.type)
        ))
        : [];
      const sourceRows = Array.isArray(row.values[field.code]) ? row.values[field.code] : [];
      const editorRows = sourceRows.map((item, idx) => ({
        localId: `row:${idx}:${Date.now()}`,
        id: item?.id ? String(item.id) : '',
        value: item && typeof item.value === 'object' ? JSON.parse(JSON.stringify(item.value)) : {}
      }));

      const layer = document.createElement('div');
      layer.className = 'pb-overlay__subtable-layer';
      const panel = document.createElement('div');
      panel.className = 'pb-overlay__subtable-panel';
      panel.addEventListener('click', (event) => event.stopPropagation());

      const head = document.createElement('div');
      head.className = 'pb-overlay__subtable-head';
      const title = document.createElement('div');
      title.className = 'pb-overlay__subtable-title';
      title.textContent = `${resolveText(this.language, 'subtableTitle')}: ${field.label || field.code}`;
      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.className = 'pb-overlay__subtable-close';
      closeBtn.textContent = '×';
      closeBtn.title = resolveText(this.language, 'subtableClose');
      closeBtn.setAttribute('aria-label', resolveText(this.language, 'subtableClose'));
      closeBtn.addEventListener('click', () => this.closeSubtableEditor(true));
      head.appendChild(title);
      head.appendChild(closeBtn);

      const body = document.createElement('div');
      body.className = 'pb-overlay__subtable-body';

      const tableWrap = document.createElement('div');
      tableWrap.className = 'pb-subtable-wrap';
      const table = document.createElement('table');
      table.className = 'pb-subtable';
      const thead = document.createElement('thead');
      const theadRow = document.createElement('tr');
      childFields.forEach((child) => {
        const th = document.createElement('th');
        th.className = 'pb-subtable__head';
        th.textContent = child.label || child.code;
        theadRow.appendChild(th);
      });
      const opTh = document.createElement('th');
      opTh.className = 'pb-subtable__head pb-subtable__op';
      opTh.setAttribute('aria-label', resolveText(this.language, 'subtableActionsAria'));
      theadRow.appendChild(opTh);
      thead.appendChild(theadRow);
      const tbody = document.createElement('tbody');
      tbody.className = 'pb-subtable__body';
      table.appendChild(thead);
      table.appendChild(tbody);
      tableWrap.appendChild(table);

      const footer = document.createElement('div');
      footer.className = 'pb-overlay__subtable-foot';
      const addBtn = document.createElement('button');
      addBtn.type = 'button';
      addBtn.className = 'pb-overlay__btn';
      addBtn.dataset.subtableAction = 'add';
      addBtn.textContent = resolveText(this.language, 'subtableAddRow');
      const cancelBtn = document.createElement('button');
      cancelBtn.type = 'button';
      cancelBtn.className = 'pb-overlay__btn';
      cancelBtn.dataset.subtableAction = 'cancel';
      cancelBtn.textContent = resolveText(this.language, 'subtableCancel');
      const saveBtn = document.createElement('button');
      saveBtn.type = 'button';
      saveBtn.className = 'pb-overlay__btn pb-overlay__btn--primary';
      saveBtn.dataset.subtableAction = 'save';
      saveBtn.textContent = resolveText(this.language, 'subtableSave');
      footer.appendChild(addBtn);
      footer.appendChild(cancelBtn);
      footer.appendChild(saveBtn);

      const renderRows = () => {
        tbody.textContent = '';
        if (!editorRows.length) {
          const emptyRow = document.createElement('tr');
          emptyRow.className = 'pb-subtable__empty-row';
          const emptyCell = document.createElement('td');
          emptyCell.className = 'pb-subtable__empty';
          emptyCell.colSpan = Math.max(1, childFields.length + 1);
          emptyCell.textContent = resolveText(this.language, 'subtableEmpty');
          emptyRow.appendChild(emptyCell);
          tbody.appendChild(emptyRow);
          return;
        }
        editorRows.forEach((item) => {
          const rowEl = document.createElement('tr');
          rowEl.className = 'pb-subtable__row';

          childFields.forEach((child) => {
            const cell = document.createElement('td');
            cell.className = 'pb-subtable__cell';
            const input = document.createElement('input');
            input.className = 'pb-overlay__subtable-input';
            input.type = child.type === 'DATE' ? 'date' : 'text';
            if (child.type === 'NUMBER') input.inputMode = 'decimal';
            if (child.type === 'RADIO_BUTTON' || child.type === 'DROP_DOWN') {
              input.placeholder = Array.isArray(child.choices) ? (child.choices[0] || '') : '';
            }
            const rawCell = item.value?.[child.code]?.value;
            input.value = rawCell === undefined || rawCell === null ? '' : String(rawCell);
            input.disabled = readOnlyMode;
            input.addEventListener('input', () => {
              if (readOnlyMode) return;
              if (!item.value || typeof item.value !== 'object') item.value = {};
              item.value[child.code] = { value: input.value };
            });
            cell.appendChild(input);
            rowEl.appendChild(cell);
          });

          const opCell = document.createElement('td');
          opCell.className = 'pb-subtable__cell pb-subtable__op';
          const removeBtn = document.createElement('button');
          removeBtn.type = 'button';
          removeBtn.className = 'pb-overlay__btn pb-subtable__remove';
          removeBtn.textContent = '×';
          removeBtn.title = resolveText(this.language, 'subtableRemoveRow');
          removeBtn.setAttribute('aria-label', resolveText(this.language, 'subtableRemoveRow'));
          removeBtn.disabled = readOnlyMode;
          removeBtn.addEventListener('click', () => {
            if (readOnlyMode) return;
            const idx = editorRows.indexOf(item);
            if (idx >= 0) {
              editorRows.splice(idx, 1);
              renderRows();
            }
          });
          opCell.appendChild(removeBtn);
          rowEl.appendChild(opCell);
          tbody.appendChild(rowEl);
        });
      };

      addBtn.addEventListener('click', () => {
        if (readOnlyMode) return;
        const value = {};
        childFields.forEach((child) => {
          value[child.code] = { value: this.createSubtableDefaultValue(child.type) };
        });
        editorRows.push({
          localId: `row:new:${Date.now()}:${Math.random().toString(16).slice(2)}`,
          id: '',
          value
        });
        renderRows();
      });

      cancelBtn.addEventListener('click', () => this.closeSubtableEditor(true));
      saveBtn.addEventListener('click', () => {
        if (readOnlyMode) {
          this.notifyViewOnlyBlocked();
          return;
        }
        const beforeValueSnapshot = this.deepClone(row.values[field.code]);
        for (const item of editorRows) {
          for (const child of childFields) {
            const raw = item?.value?.[child.code]?.value ?? '';
            const validation = this.validate(raw, child);
            if (!validation.ok) {
              this.notify(resolveText(this.language, 'toastInvalidCells'));
              return;
            }
            if (!item.value || typeof item.value !== 'object') item.value = {};
            item.value[child.code] = { value: validation.value };
          }
        }
        const nextValue = editorRows.map((item) => {
          const out = { value: item.value && typeof item.value === 'object' ? JSON.parse(JSON.stringify(item.value)) : {} };
          if (item.id) out.id = item.id;
          return out;
        });
        const currentRow = this.rowMap.get(row.id);
        if (!currentRow) {
          this.closeSubtableEditor(true);
          return;
        }
        currentRow.values[field.code] = nextValue;
        const originalValue = currentRow.original[field.code];
        const changed = JSON.stringify(nextValue) !== JSON.stringify(originalValue);
        if (changed) {
          this.addDiff(currentRow.id, field.code, nextValue, 'SUBTABLE');
          this.pushHistory({
            type: 'subtable-edit',
            payload: {
              recordId: currentRow.id,
              fieldCode: field.code,
              before: beforeValueSnapshot,
              after: this.deepClone(nextValue)
            }
          });
        } else {
          this.removeDiff(currentRow.id, field.code);
        }
        const input = this.findInput(currentRow.id, field.code);
        if (input) {
          this.applyCellVisualState(input, currentRow, field.code);
        }
        this.updateDirtyBadge();
        this.updateStats();
        this.closeSubtableEditor(true);
      });
      addBtn.disabled = readOnlyMode;
      saveBtn.disabled = readOnlyMode;

      layer.addEventListener('click', () => this.closeSubtableEditor(true));
      layer.appendChild(panel);
      body.appendChild(tableWrap);
      panel.appendChild(head);
      panel.appendChild(body);
      panel.appendChild(footer);
      this.root.appendChild(layer);
      renderRows();
      this.subtableEditor = {
        layer,
        panel,
        rowIndex,
        colIndex,
        anchorInput: anchorInput || this.getInput(rowIndex, colIndex)
      };
    }

    closeSubtableEditor(focusInput = false) {
      if (!this.subtableEditor) return;
      const { layer, anchorInput } = this.subtableEditor;
      if (layer?.parentElement) layer.remove();
      this.subtableEditor = null;
      if (focusInput && anchorInput) {
        try {
          anchorInput.focus();
        } catch (_e) { /* noop */ }
      }
    }

    handleCellMouseDown(event, input) {
      if (this.saving) return;
      if (event.button !== 0) return;
      const rowIndex = Number(input.dataset.rowIndex || '0');
      const colIndex = Number(input.dataset.colIndex || '0');
      const editing = this.isEditingCell(rowIndex, colIndex);

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
      if (event.shiftKey) {
        this.startSelection(rowIndex, colIndex, true);
        this.armedCell = { rowIndex, colIndex };
        return;
      }
      this.startSelection(rowIndex, colIndex, false);
      this.armedCell = { rowIndex, colIndex };
      if (event.detail >= 2) {
        this.enterEditMode(rowIndex, colIndex, input, true);
      }
    }

    handleCellHover(input) {
      if (!this.dragSelecting || this.editingCell) return;
      const rowIndex = Number(input.dataset.rowIndex || '0');
      const colIndex = Number(input.dataset.colIndex || '0');
      this.updateSelectionRange(rowIndex, colIndex);
    }

    handleChoiceInputClick(input) {
      if (this.saving) return;
      if (!input) return;
      const fieldCode = input.dataset.fieldCode;
      const recordId = input.dataset.recordId || '';
      if (!fieldCode) return;
      const field = this.fieldMap.get(fieldCode);
      const rowIndex = Number(input.dataset.rowIndex || '0');
      const colIndex = Number(input.dataset.colIndex || '0');
      if (!field) return;
      if (this.isSubtableField(field)) {
        this.openSubtableEditor(rowIndex, colIndex, input);
        return;
      }
      if (!this.isChoiceField(field)) return;
      if (!this.permissionService.canEditCell(recordId, fieldCode)) return;
      if (!this.isEditingCell(rowIndex, colIndex)) return;
      if (this.radioPicker && this.radioPicker.input === input) return;
      this.openRadioPicker(rowIndex, colIndex, input);
    }

    onInputChanged(input) {
      if (!this.canMutateOverlay(false)) return;
      const recordId = input.dataset.recordId;
      const fieldCode = input.dataset.fieldCode;
      if (!recordId || !fieldCode) return;
      const key = `${recordId}:${fieldCode}`;
      const field = this.fieldMap.get(fieldCode);
      if (!field) return;
      const cellPermission = this.getCellPermissionInfo(recordId, fieldCode);
      if (!cellPermission.editable || !this.isOverlayEditable()) return;
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
      const row = this.rowMap.get(recordId);
      const originalValue = row ? row.original[fieldCode] : input.dataset.originalValue;
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
      if (this.hasActiveFilters()) {
        const rowIndex = Number(input.dataset.rowIndex || '-1');
        const colIndex = Number(input.dataset.colIndex || '-1');
        const editingNow = this.isEditingCell(rowIndex, colIndex) && !input.readOnly;
        if (editingNow) {
          this.needsFilterReapply = true;
        } else {
          this.applyFilters();
          return;
        }
      }
      this.updateStats();
    }

    handlePaste(event, input) {
      if (!this.canMutateOverlay(true)) return;
      if (!event.clipboardData) return;
      const text = event.clipboardData.getData('text');
      if (!text) return;
      event.preventDefault();
      const { matrix, rowCount, colCount } = this.parseClipboardMatrix(text);
      if (!rowCount || !colCount) return;
      const startRow = Number(input?.dataset?.rowIndex || '0');
      const startCol = Number(input?.dataset?.colIndex || '0');
      this.applyPastedMatrixAt(matrix, startRow, startCol);
    }

    parseClipboardMatrix(text) {
      const normalized = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      const lines = normalized.split('\n');
      if (lines.length && lines[lines.length - 1] === '') {
        lines.pop();
      }
      if (!lines.length) return { matrix: [], rowCount: 0, colCount: 0 };
      const matrix = lines.map((line) => line.split('\t'));
      const rowCount = matrix.length;
      const colCount = matrix.reduce((max, row) => Math.max(max, row.length), 0);
      return { matrix, rowCount, colCount };
    }

    getPasteStartCell() {
      if (this.selection) {
        return {
          rowIndex: this.selection.startRow,
          colIndex: this.selection.startCol
        };
      }
      const active = document.activeElement;
      if (active instanceof HTMLInputElement) {
        return {
          rowIndex: Number(active.dataset.rowIndex || '0'),
          colIndex: Number(active.dataset.colIndex || '0')
        };
      }
      return { rowIndex: 0, colIndex: 0 };
    }

    async pasteClipboard() {
      if (!this.canMutateOverlay(true)) return;
      if (this.saving) return;
      let text = '';
      try {
        if (!navigator?.clipboard?.readText) return;
        text = await navigator.clipboard.readText();
      } catch (error) {
        console.warn('[kintone-excel-overlay] clipboard read failed', error);
        return;
      }
      const { matrix, rowCount, colCount } = this.parseClipboardMatrix(text);
      if (!rowCount || !colCount) return;
      const start = this.getPasteStartCell();
      this.exitEditMode();
      this.applyPastedMatrixAt(matrix, start.rowIndex, start.colIndex);
    }

    applyPastedMatrixAt(matrix, startRow, startCol) {
      if (!this.canMutateOverlay(true)) return;
      if (!Array.isArray(matrix) || !matrix.length) return;
      const rowCount = this.getVisibleRowCount();
      const colCount = this.fields.length;
      let attemptedRows = 0;
      let attemptedCols = 0;
      const targets = [];
      for (let r = 0; r < matrix.length; r += 1) {
        const targetRowIndex = startRow + r;
        if (targetRowIndex >= rowCount) break;
        const rowValues = matrix[r];
        for (let c = 0; c < rowValues.length; c += 1) {
          const targetColIndex = startCol + c;
          if (targetColIndex >= colCount) break;
          const row = this.getVisibleRowAt(targetRowIndex);
          const field = this.fields[targetColIndex];
          if (!row || !field) continue;
          targets.push({
            recordId: String(row.id || '').trim(),
            fieldCode: String(field.code || '').trim(),
            rowIndex: targetRowIndex,
            colIndex: targetColIndex,
            value: rowValues[c]
          });
          attemptedCols = Math.max(attemptedCols, c + 1);
        }
        attemptedRows = r + 1;
      }

      const filteredTargets = this.permissionService.filterPasteTargets(targets);
      const applicableTargets = Array.isArray(filteredTargets?.applicable) ? filteredTargets.applicable : [];
      const skippedTargets = Array.isArray(filteredTargets?.skipped) ? filteredTargets.skipped : [];
      const pasteChanges = [];
      applicableTargets.forEach((target) => {
        const change = this.applyPastedValue(target.rowIndex, target.colIndex, target.value);
        if (change) pasteChanges.push(change);
      });

      if (skippedTargets.length) {
        this.notify(resolveText(this.language, 'toastPasteSkipped', skippedTargets.length));
      }
      if (!attemptedRows || !attemptedCols) return;
      if (!applicableTargets.length) return;
      if (pasteChanges.length) {
        this.pushHistory({
          type: 'paste',
          payload: { changes: pasteChanges }
        });
      }
      const endRow = Math.min(rowCount - 1, startRow + attemptedRows - 1);
      const endCol = Math.min(colCount - 1, startCol + attemptedCols - 1);
      this.selection = {
        startRow,
        endRow,
        startCol,
        endCol,
        anchorRow: startRow,
        anchorCol: startCol
      };
      this.repaintSelection();
      this.updateStats();
      this.flashPastedRange(startRow, endRow, startCol, endCol);
      if (this.hasActiveFilters()) {
        this.needsFilterReapply = false;
        this.applyFilters();
      }
    }

    flashPastedRange(startRow, endRow, startCol, endCol) {
      const flashedCells = [];
      for (let r = startRow; r <= endRow; r += 1) {
        for (let c = startCol; c <= endCol; c += 1) {
          const input = this.getInput(r, c);
          const cell = input?.parentElement;
          if (!cell) continue;
          cell.classList.add('pb-overlay__cell--paste-flash');
          flashedCells.push(cell);
        }
      }
      if (this.pasteFlashTimer) {
        clearTimeout(this.pasteFlashTimer);
      }
      this.pasteFlashTimer = setTimeout(() => {
        flashedCells.forEach((cell) => cell.classList.remove('pb-overlay__cell--paste-flash'));
        this.pasteFlashTimer = null;
      }, 220);
    }

    applyPastedValue(rowIndex, colIndex, rawValue) {
      const row = this.getVisibleRowAt(rowIndex);
      const field = this.fields[colIndex];
      if (!row || !field) return null;
      if (this.isSubtableField(field)) return null;
      const cellPermission = this.getCellPermissionInfo(row.id, field.code);
      if (!this.isOverlayEditable() || !cellPermission.editable) return null;
      const beforeValue = this.deepClone(row.values[field.code]);
      const input = this.getInput(rowIndex, colIndex);
      if (input) {
        input.value = rawValue;
        this.onInputChanged(input);
        const afterValue = this.deepClone(row.values[field.code]);
        if (this.valuesEqual(beforeValue, afterValue)) return null;
        return {
          recordId: row.id,
          fieldCode: field.code,
          before: beforeValue,
          after: afterValue
        };
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
        return null;
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
      const afterValue = this.deepClone(row.values[fieldCode]);
      if (this.valuesEqual(beforeValue, afterValue)) return null;
      return {
        recordId,
        fieldCode,
        before: beforeValue,
        after: afterValue
      };
    }

    setSelectionSingle(rowIndex, colIndex) {
      const maxRow = Math.max(0, this.getVisibleRowCount() - 1);
      const maxCol = Math.max(0, this.fields.length - 1);
      if (maxRow < 0 || maxCol < 0) {
        this.selection = null;
        this.armedCell = null;
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
      this.armedCell = { rowIndex: r, colIndex: c };
      this.dragSelecting = false;
      this.repaintSelection();
      this.updateStats();
    }

    setSelection(rowIndex, colIndex, options = {}) {
      const maxRow = Math.max(0, this.getVisibleRowCount() - 1);
      const maxCol = Math.max(0, this.fields.length - 1);
      if (maxRow < 0 || maxCol < 0) {
        this.selection = null;
        this.armedCell = null;
        this.repaintSelection();
        this.updateStats();
        return null;
      }
      const r = Math.max(0, Math.min(maxRow, rowIndex));
      const c = Math.max(0, Math.min(maxCol, colIndex));
      const extend = options?.extend === true;

      if (extend) {
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
          const anchorRow = typeof this.selection.anchorRow === 'number' ? this.selection.anchorRow : this.selection.startRow;
          const anchorCol = typeof this.selection.anchorCol === 'number' ? this.selection.anchorCol : this.selection.startCol;
          this.selection.startRow = Math.min(anchorRow, r);
          this.selection.endRow = Math.max(anchorRow, r);
          this.selection.startCol = Math.min(anchorCol, c);
          this.selection.endCol = Math.max(anchorCol, c);
        }
      } else {
        this.selection = {
          anchorRow: r,
          anchorCol: c,
          startRow: r,
          endRow: r,
          startCol: c,
          endCol: c
        };
      }

      this.armedCell = { rowIndex: r, colIndex: c };
      this.dragSelecting = false;
      this.repaintSelection();
      this.updateStats();
      this.syncSelectionScroll({ rowIndex: r, colIndex: c });
      if (options?.focus !== false) {
        this.focusCell(r, c, { ensureVisible: false });
      }
      return { row: r, col: c };
    }

    startSelection(rowIndex, colIndex, extend = false) {
      this.exitEditMode();
      const maxRow = Math.max(0, this.getVisibleRowCount() - 1);
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
      this.armedCell = { rowIndex: r, colIndex: c };
      this.dragSelecting = true;
      this.repaintSelection();
      this.updateStats();
      this.focusCell(r, c);
    }

    updateSelectionRange(rowIndex, colIndex, focus = false) {
      const maxRow = Math.max(0, this.getVisibleRowCount() - 1);
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

    isCellArmed(rowIndex, colIndex) {
      if (!this.armedCell) return false;
      if (this.editingCell) return false;
      return this.armedCell.rowIndex === rowIndex && this.armedCell.colIndex === colIndex;
    }

    isCellActive(rowIndex, colIndex) {
      if (this.isEditingCell(rowIndex, colIndex)) return true;
      if (!this.armedCell) return false;
      return this.armedCell.rowIndex === rowIndex && this.armedCell.colIndex === colIndex;
    }

    repaintSelection() {
      this.inputsByRow.forEach((rowInputs, rowIndex) => {
        if (!Array.isArray(rowInputs)) return;
        rowInputs.forEach((input, colIndex) => {
          if (!input) return;
          const cell = input.parentElement;
          if (!cell) return;
          const r = Number(rowIndex);
          const c = Number(colIndex);
          if (this.isCellSelected(Number(rowIndex), Number(colIndex))) {
            cell.classList.add('pb-overlay__cell--selected');
          } else {
            cell.classList.remove('pb-overlay__cell--selected');
          }
          if (this.isCellArmed(r, c)) {
            cell.classList.add('pb-overlay__cell--armed');
          } else {
            cell.classList.remove('pb-overlay__cell--armed');
          }
          if (this.isCellActive(r, c)) {
            cell.classList.add('pb-overlay__cell--active');
          } else {
            cell.classList.remove('pb-overlay__cell--active');
          }
          if (this.isEditingCell(r, c) && input.classList.contains('pb-overlay__input--editing')) {
            cell.classList.add('pb-overlay__cell--editing');
          } else {
            cell.classList.remove('pb-overlay__cell--editing');
          }
        });
      });
    }

    isEditingCell(rowIndex, colIndex) {
      if (!this.editingCell) return false;
      return this.editingCell.rowIndex === rowIndex && this.editingCell.colIndex === colIndex;
    }

    enterEditMode(rowIndex, colIndex, input, selectAll = false) {
      if (this.saving) return;
      const maxRow = Math.max(0, this.getVisibleRowCount() - 1);
      const maxCol = Math.max(0, this.fields.length - 1);
      if (maxRow < 0 || maxCol < 0) return;
      const r = Math.max(0, Math.min(maxRow, rowIndex));
      const c = Math.max(0, Math.min(maxCol, colIndex));
      const field = this.fields[c];
      const targetRow = this.getVisibleRowAt(r);
      const targetRecordId = String(targetRow?.id || '').trim();
      const cellPermission = this.getCellPermissionInfo(targetRecordId, String(field?.code || ''));
      if (!this.isOverlayEditable() || !cellPermission.editable) {
        if (this.isSubtableField(field)) {
          if (targetRecordId) {
            this.openSubtableEditor(r, c, input || this.getInput(r, c));
          }
        }
        return;
      }
      if (this.editingCell && (this.editingCell.rowIndex !== r || this.editingCell.colIndex !== c)) {
        this.exitEditMode();
      }
      this.dragSelecting = false;
      this.setSelectionSingle(r, c);
      this.armedCell = null;
      this.pendingEdit = null;
      this.editingCell = this.createEditingState(r, c, !!selectAll);
      const targetInput = input || this.getInput(r, c);
      if (targetInput) {
        this.setInputEditingVisual(targetInput, true);
        requestAnimationFrame(() => {
          targetInput.focus();
          if (this.editingCell && this.editingCell.rowIndex === r && this.editingCell.colIndex === c) {
            this.applyCaretPlacement(targetInput, !!this.editingCell.selectAll);
          }
          if (this.isChoiceField(field)) {
            this.openRadioPicker(r, c, targetInput);
          }
        });
      }
    }

    exitEditMode() {
      if (!this.editingCell) return;
      const previousEditingCell = this.editingCell;
      const { rowIndex, colIndex } = previousEditingCell;
      this.closeRadioPicker();
      const input = this.getInput(rowIndex, colIndex);
      if (input) {
        this.setInputEditingVisual(input, false);
      }
      this.commitCellEditHistory(previousEditingCell);
      this.editingCell = null;
      this.pendingEdit = null;
      this.setSelectionSingle(rowIndex, colIndex);
    }

    createEditingState(rowIndex, colIndex, selectAll = false) {
      const row = this.getVisibleRowAt(rowIndex);
      const field = this.fields[colIndex];
      const recordId = String(row?.id || '').trim();
      const fieldCode = String(field?.code || '').trim();
      const beforeValue = row && fieldCode ? this.deepClone(row.values[fieldCode]) : '';
      return {
        rowIndex,
        colIndex,
        selectAll: Boolean(selectAll),
        recordId,
        fieldCode,
        beforeValue
      };
    }

    commitCellEditHistory(editState) {
      if (!editState || this.isReplayingHistory) return;
      const recordId = String(editState.recordId || '').trim();
      const fieldCode = String(editState.fieldCode || '').trim();
      if (!recordId || !fieldCode) return;
      const row = this.rowMap.get(recordId);
      if (!row) return;
      const before = this.deepClone(editState.beforeValue);
      const after = this.deepClone(row.values[fieldCode]);
      if (this.valuesEqual(before, after)) return;
      this.pushHistory({
        type: 'cell-edit',
        payload: {
          recordId,
          fieldCode,
          before,
          after
        }
      });
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
        const rowData = this.getVisibleRowAt(r);
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
      const rowIndex = this.getVisibleRowIndexById(recordId);
      if (rowIndex < 0) return null;
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
      count += this.pendingDeletes.size;
      this.dirtyBadge.textContent = resolveText(this.language, 'dirtyLabel', count);
      if (count > 0) {
        this.dirtyBadge.classList.add('pb-overlay__dirty--active');
      } else {
        this.dirtyBadge.classList.remove('pb-overlay__dirty--active');
      }
      if (this.saveButton && !this.saving) {
        const canSave = this.canSaveOverlay();
        this.saveButton.disabled = !canSave || count === 0;
        this.saveButton.title = canSave ? '' : resolveText(this.language, 'toastViewOnlyBlocked');
      }
    }

    updateStats() {
      const numbers = [];
      const visibleRows = this.getVisibleRows();
      let selectedCells = 0;
      if (this.selection) {
        const selectedRows = Math.max(0, this.selection.endRow - this.selection.startRow + 1);
        const selectedCols = Math.max(0, this.selection.endCol - this.selection.startCol + 1);
        selectedCells = selectedRows * selectedCols;
        for (let r = this.selection.startRow; r <= this.selection.endRow; r += 1) {
          const row = visibleRows[r];
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
      const filteredRowsCount = visibleRows.length;
      const pendingDeleteCount = this.pendingDeletes.size;
      const newRowsCount = this.rows.reduce((acc, row) => acc + (this.isNewRowId(row?.id) ? 1 : 0), 0);

      if (this.statsElements.selected) {
        this.statsElements.selected.textContent = `${resolveText(this.language, 'statusSelectedCells')}: ${selectedCells}`;
      }
      if (this.statsElements.filtered) {
        this.statsElements.filtered.textContent = `${resolveText(this.language, 'statusFilteredRows')}: ${filteredRowsCount}`;
      }
      if (this.statsElements.pendingDelete) {
        this.statsElements.pendingDelete.textContent = `${resolveText(this.language, 'statusPendingDeletes')}: ${pendingDeleteCount}`;
      }
      if (this.statsElements.newRows) {
        this.statsElements.newRows.textContent = `${resolveText(this.language, 'statusNewRows')}: ${newRowsCount}`;
      }
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
      if (type === 'RADIO_BUTTON' || type === 'DROP_DOWN') {
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
      const size = Number(limit);
      if (!Number.isFinite(size) || size <= 0) return String(baseQuery || '').trim();
      return this.normalizeOverlayQuery(baseQuery, Math.floor(size), 0);
    }

    deepClone(value) {
      if (value === undefined) return undefined;
      if (typeof structuredClone === 'function') {
        try {
          return structuredClone(value);
        } catch (_err) {
          // fallback below
        }
      }
      return JSON.parse(JSON.stringify(value));
    }

    valuesEqual(a, b) {
      return JSON.stringify(a) === JSON.stringify(b);
    }

    pushHistory(entry) {
      if (!entry || this.isReplayingHistory) return;
      this.undoStack.push({
        ...entry,
        timestamp: Date.now()
      });
      if (this.undoStack.length > this.maxHistory) {
        this.undoStack.shift();
      }
      this.redoStack = [];
      this.updateHistoryButtons();
    }

    clearHistory() {
      this.undoStack = [];
      this.redoStack = [];
      this.updateHistoryButtons();
    }

    updateHistoryButtons() {
      if (this.undoButton) {
        this.undoButton.disabled = this.saving || !this.canEditOverlay() || this.undoStack.length === 0;
      }
      if (this.redoButton) {
        this.redoButton.disabled = this.saving || !this.canEditOverlay() || this.redoStack.length === 0;
      }
    }

    setRowFieldValueFromHistory(recordId, fieldCode, nextValue) {
      const key = String(recordId || '').trim();
      const code = String(fieldCode || '').trim();
      if (!key || !code) return;
      const row = this.rowMap.get(key);
      const field = this.fieldMap.get(code);
      if (!row || !field) return;
      const value = this.deepClone(nextValue);
      row.values[code] = value;
      const original = row.original[code];
      if (this.valuesEqual(value, original)) {
        this.removeDiff(key, code);
      } else {
        this.addDiff(key, code, value, field.type);
      }
      const invalidKey = `${key}:${code}`;
      this.invalidCells.delete(invalidKey);
      this.invalidValueCache.delete(invalidKey);
      const input = this.findInput(key, code);
      if (input) {
        input.value = this.formatCellDisplayValue(field, row.values[code]);
        this.applyCellVisualState(input, row, code);
      }
    }

    applyDeleteToggleState(recordId, nextState) {
      const key = String(recordId || '').trim();
      if (!key || this.isNewRowId(key)) return;
      const shouldSet = Boolean(nextState);
      if (shouldSet) this.pendingDeletes.add(key);
      else this.pendingDeletes.delete(key);
      this.syncPermissionServicePendingDeletes();
      const editingRowId = this.editingCell ? this.getVisibleRowAt(this.editingCell.rowIndex)?.id : '';
      if (editingRowId && String(editingRowId) === key) {
        this.exitEditMode();
      }
      this.updateVirtualRows(true);
      this.updateDirtyBadge();
      this.updateStats();
    }

    restoreHistoryAddRow(rowSnapshot, index) {
      const row = this.deepClone(rowSnapshot);
      const id = String(row?.id || '').trim();
      if (!id) return;
      if (this.rowMap.has(id)) return;
      const insertAt = Math.max(0, Math.min(Number(index) || 0, this.rows.length));
      this.rows.splice(insertAt, 0, row);
      this.rebuildRowMaps();
      this.recomputeFilteredRowIds();
      this.updateVirtualRows(true);
      this.updateDirtyBadge();
      this.updateStats();
    }

    removeHistoryAddRow(rowId) {
      const key = String(rowId || '').trim();
      if (!key) return;
      const index = this.rows.findIndex((row) => String(row?.id || '').trim() === key);
      if (index < 0) return;
      if (this.editingCell && this.getVisibleRowAt(this.editingCell.rowIndex)?.id === key) {
        this.exitEditMode();
      }
      this.rows.splice(index, 1);
      this.pendingDeletes.delete(key);
      this.syncPermissionServicePendingDeletes();
      this.diff.delete(key);
      this.rowMap.delete(key);
      this.rowIndexMap.delete(key);
      this.revisionMap.delete(key);
      this.permissionService.removeRecordAcl(key);
      this.clearInvalidStateForRecord(key);
      this.rebuildRowMaps();
      this.recomputeFilteredRowIds();
      const visibleCount = this.getVisibleRowCount();
      if (!visibleCount) {
        this.selection = null;
        this.armedCell = null;
      } else if (this.selection) {
        const nextRow = Math.max(0, Math.min(visibleCount - 1, this.selection.startRow || 0));
        const nextCol = Math.max(0, Math.min(this.fields.length - 1, this.selection.startCol || 0));
        this.setSelectionSingle(nextRow, nextCol);
      }
      this.updateVirtualRows(true);
      this.updateDirtyBadge();
      this.updateStats();
    }

    applyHistoryEntry(entry, reverse = false) {
      if (!entry || !entry.type) return;
      const payload = entry.payload || {};
      let shouldRecomputeFilters = false;
      if (entry.type === 'cell-edit') {
        const nextValue = reverse ? payload.before : payload.after;
        this.setRowFieldValueFromHistory(payload.recordId, payload.fieldCode, nextValue);
        shouldRecomputeFilters = true;
      } else if (entry.type === 'paste') {
        const changes = Array.isArray(payload.changes) ? payload.changes.slice() : [];
        if (reverse) changes.reverse();
        changes.forEach((change) => {
          const nextValue = reverse ? change.before : change.after;
          this.setRowFieldValueFromHistory(change.recordId, change.fieldCode, nextValue);
        });
        shouldRecomputeFilters = true;
      } else if (entry.type === 'add-row') {
        if (reverse) {
          this.removeHistoryAddRow(payload.row?.id);
        } else {
          this.restoreHistoryAddRow(payload.row, payload.index);
        }
      } else if (entry.type === 'toggle-delete') {
        const nextValue = reverse ? payload.before : payload.after;
        this.applyDeleteToggleState(payload.recordId, nextValue);
      } else if (entry.type === 'subtable-edit') {
        const nextValue = reverse ? payload.before : payload.after;
        this.setRowFieldValueFromHistory(payload.recordId, payload.fieldCode, nextValue);
        shouldRecomputeFilters = true;
      }
      if (shouldRecomputeFilters && this.hasActiveFilters()) {
        this.recomputeFilteredRowIds();
      }
    }

    undo() {
      if (!this.canMutateOverlay(true)) return;
      if (this.saving || !this.undoStack.length) return;
      const entry = this.undoStack.pop();
      this.isReplayingHistory = true;
      try {
        this.applyHistoryEntry(entry, true);
        this.redoStack.push(entry);
      } finally {
        this.isReplayingHistory = false;
      }
      this.updateHistoryButtons();
      this.updateVirtualRows(true);
      this.repaintSelection();
      this.updateDirtyBadge();
      this.updateStats();
    }

    redo() {
      if (!this.canMutateOverlay(true)) return;
      if (this.saving || !this.redoStack.length) return;
      const entry = this.redoStack.pop();
      this.isReplayingHistory = true;
      try {
        this.applyHistoryEntry(entry, false);
        this.undoStack.push(entry);
      } finally {
        this.isReplayingHistory = false;
      }
      this.updateHistoryButtons();
      this.updateVirtualRows(true);
      this.repaintSelection();
      this.updateDirtyBadge();
      this.updateStats();
    }

    syncScroll() {
      if (!this.bodyScroll) return;
      const scrollLeft = this.bodyScroll.scrollLeft;
      const scrollTop = this.bodyScroll.scrollTop;
      if (this.columnHeaderScroll) {
        this.columnHeaderScroll.style.transform = `translateX(${-scrollLeft}px)`;
      }
      if (this.rowHeaderScroll) {
        this.rowHeaderScroll.scrollTop = scrollTop;
      }
      this.updateVirtualRows();
    }

    focusFirstCell() {
      if (!this.getVisibleRowCount() || !this.fields.length) return;
      this.setSelectionSingle(0, 0);
      this.focusCell(0, 0);
    }

    findFirstEditableColIndex(recordId = '') {
      const index = this.fields.findIndex((field) => this.permissionService.canEditCell(recordId, field.code));
      return index >= 0 ? index : 0;
    }

    addNewRow() {
      if (!this.canMutateOverlay(true)) return;
      if (this.saving) return;
      if (!this.permissionService.canAddRow()) return;
      if (!this.fields.length) return;
      const tempId = `NEW:${Date.now()}:${this.newRowSeq++}`;
      const values = {};
      const original = {};
      this.fields.forEach((field) => {
        const base = this.isSubtableField(field) ? [] : '';
        values[field.code] = this.cloneFieldValue(field, base);
        original[field.code] = this.cloneFieldValue(field, base);
      });
      const row = {
        id: tempId,
        revision: undefined,
        values,
        original,
        isNew: true
      };
      this.rows.unshift(row);
      this.pushHistory({
        type: 'add-row',
        payload: {
          row: this.deepClone(row),
          index: 0
        }
      });
      this.rebuildRowMaps();
      this.recomputeFilteredRowIds();
      if (this.bodyScroll) {
        this.bodyScroll.scrollTop = 0;
      }
      this.updateVirtualRows(true);
      const colIndex = this.findFirstEditableColIndex(tempId);
      this.setSelectionSingle(0, colIndex);
      this.focusCell(0, colIndex);
      const targetInput = this.getInput(0, colIndex);
      if (targetInput) {
        this.enterEditMode(0, colIndex, targetInput, true);
      }
    }

    removeNewRowById(recordId) {
      const key = String(recordId || '').trim();
      if (!key || !this.isNewRowId(key)) return;
      const index = this.rows.findIndex((row) => String(row?.id || '').trim() === key);
      if (index < 0) return;
      if (this.editingCell && this.getVisibleRowAt(this.editingCell.rowIndex)?.id === key) {
        this.exitEditMode();
      }
      this.rows.splice(index, 1);
      this.pendingDeletes.delete(key);
      this.syncPermissionServicePendingDeletes();
      this.diff.delete(key);
      this.permissionService.removeRecordAcl(key);
      this.clearInvalidStateForRecord(key);
      this.rebuildRowMaps();
      this.recomputeFilteredRowIds();
      const visibleCount = this.getVisibleRowCount();
      if (!visibleCount) {
        this.selection = null;
        this.armedCell = null;
      } else if (this.selection) {
        const nextRow = Math.max(0, Math.min(visibleCount - 1, this.selection.startRow || 0));
        const nextCol = Math.max(0, Math.min(this.fields.length - 1, this.selection.startCol || 0));
        this.setSelectionSingle(nextRow, nextCol);
      }
      this.updateVirtualRows(true);
      this.updateDirtyBadge();
      this.updateStats();
    }

    togglePendingDelete(recordId) {
      if (!this.canMutateOverlay(true)) return;
      const key = String(recordId || '').trim();
      if (!key || this.saving) return;
      if (this.isNewRowId(key)) {
        this.removeNewRowById(key);
        return;
      }
      if (!this.permissionService.canDeleteRecord(key)) return;
      const before = this.pendingDeletes.has(key);
      if (this.pendingDeletes.has(key)) {
        this.pendingDeletes.delete(key);
      } else {
        this.pendingDeletes.add(key);
      }
      this.syncPermissionServicePendingDeletes();
      const after = this.pendingDeletes.has(key);
      if (before !== after) {
        this.pushHistory({
          type: 'toggle-delete',
          payload: {
            recordId: key,
            before,
            after
          }
        });
      }
      const editingRowId = this.editingCell ? this.getVisibleRowAt(this.editingCell.rowIndex)?.id : '';
      if (editingRowId && String(editingRowId) === key) {
        this.exitEditMode();
      }
      this.updateVirtualRows(true);
      this.updateDirtyBadge();
      this.updateStats();
    }

    handleOverlayKeydown(event) {
      if (!this.isOpen) return;
      const target = event.target;
      if (!this.root?.contains(target)) return;
      if (this.saving) {
        if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
        }
        return;
      }
      if (!this.isComposing && (event.ctrlKey || event.metaKey) && !event.altKey) {
        const key = String(event.key || '').toLowerCase();
        const wantsUndo = key === 'z' && !event.shiftKey;
        const wantsRedo = key === 'y' || (key === 'z' && event.shiftKey);
        if (wantsUndo || wantsRedo) {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          if (wantsUndo) {
            this.undo();
          } else {
            this.redo();
          }
          return;
        }
      }

      if (this.columnPanelLayer && this.columnPanelLayer.contains(target)) {
        if (event.key === 'Escape') {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          this.closeColumnManager();
        }
        return;
      }

      if (this.filterPanel?.panel && this.filterPanel.panel.contains(target)) {
        if (event.key === 'Escape') {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          this.closeFilterPanel();
        }
        return;
      }
      if (this.filterPanel && event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        this.closeFilterPanel();
        return;
      }

      if (this.subtableEditor?.panel && this.subtableEditor.panel.contains(target)) {
        if (event.key === 'Escape') {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          this.closeSubtableEditor(true);
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
      const recordId = targetInput?.dataset.recordId || '';
      const isChoiceField = this.isChoiceField(field);
      const isSubtableField = this.isSubtableField(field);
      const isEditableField = this.permissionService.canEditCell(recordId, fieldCode);
      const isSpaceKey = event.key === ' ' || event.key === 'Spacebar';
      const isArrowKey = event.key === 'ArrowUp'
        || event.key === 'ArrowDown'
        || event.key === 'ArrowLeft'
        || event.key === 'ArrowRight';

      if (this.handleOverlayArrowNavigation(event, targetInput instanceof HTMLInputElement ? targetInput : null)) {
        return;
      }

      if (targetInput && targetInput.readOnly && (event.key === 'Enter' || event.key === 'F2')) {
        if (event.key === 'Enter' && this.isComposing) return;
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        const rowIndex = Number(targetInput.dataset.rowIndex || '0');
        const colIndex = Number(targetInput.dataset.colIndex || '0');
        this.armedCell = { rowIndex, colIndex };
        this.enterEditMode(rowIndex, colIndex, targetInput, true);
        return;
      }

      if (targetInput && targetInput.readOnly && isSubtableField) {
        const wantsOpen = !event.ctrlKey && !event.metaKey && !event.altKey && (event.key === 'Enter' || isSpaceKey);
        if (wantsOpen) {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          const rowIndex = Number(targetInput.dataset.rowIndex || '0');
          const colIndex = Number(targetInput.dataset.colIndex || '0');
          this.openSubtableEditor(rowIndex, colIndex, targetInput);
          return;
        }
      }

      if (targetInput && targetInput.readOnly && isChoiceField && isEditableField) {
        const wantsDirectOpen = !event.ctrlKey && !event.metaKey && !event.altKey && (event.key === 'Enter' || isSpaceKey);
        if (wantsDirectOpen) {
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
        if (isChoiceField) {
          const wantsPicker = isSpaceKey && !event.ctrlKey && !event.metaKey && !event.altKey;
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
        if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'v') {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          this.pasteClipboard();
          return;
        }
        if (isArrowKey) {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          this.navigateSelectionByArrow(event.key, event.shiftKey, targetInput);
          return;
        }
        if (event.key === 'Escape') {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          this.exitEditMode();
          return;
        }
        if (event.key === 'Enter' || event.key === 'Tab') {
          if (event.key === 'Enter' && this.isComposing) return;
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          const isTabKey = event.key === 'Tab';
          const rowDelta = isTabKey ? 0 : (event.shiftKey ? -1 : 1);
          const colDelta = isTabKey ? (event.shiftKey ? -1 : 1) : 0;
          this.moveEditingFocus(rowDelta, colDelta);
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
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'v') {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        this.pasteClipboard();
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
        if ((event.key === 'Enter' || event.key === 'F2') && this.armedCell) {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          const { rowIndex, colIndex } = this.armedCell;
          const armedInput = this.getInput(rowIndex, colIndex);
          if (armedInput) {
            this.enterEditMode(rowIndex, colIndex, armedInput, true);
          }
          return;
        }
        if (event.key === 's' && (event.ctrlKey || event.metaKey)) {
          event.preventDefault();
          event.stopPropagation();
          event.stopImmediatePropagation();
          this.save();
        }
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
        if (coords) {
          this.setSelectionSingle(coords.row, coords.col);
        }
        return;
      }
      if (event.key === 'Tab') {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        const delta = event.shiftKey ? -1 : 1;
        this.exitEditMode();
        const coords = this.moveFocus(targetInput, 0, delta);
        if (coords) {
          this.setSelectionSingle(coords.row, coords.col);
        }
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

    isOverlayArrowKey(event) {
      const key = String(event?.key || '');
      return key === 'ArrowUp'
        || key === 'ArrowDown'
        || key === 'ArrowLeft'
        || key === 'ArrowRight';
    }

    isFocusInsideOverlay(target = null) {
      if (!this.root) return false;
      const active = document.activeElement;
      if (active && this.root.contains(active)) return true;
      if (target instanceof Node && this.root.contains(target)) return true;
      if (
        this.selection
        && (target === document.body || target === document.documentElement || target === document)
      ) {
        return true;
      }
      return false;
    }

    handleOverlayGlobalArrowKeydown(event) {
      if (!this.isOpen) return;
      if (event.defaultPrevented) return;
      const input = event.target instanceof HTMLInputElement ? event.target : null;
      this.handleOverlayArrowNavigation(event, input);
    }

    handleOverlayArrowNavigation(event, input = null) {
      if (!this.isOpen) return false;
      if (!this.isOverlayArrowKey(event)) return false;
      if (!this.isFocusInsideOverlay(event.target)) return false;

      const target = event.target;
      if (this.columnPanelLayer && this.columnPanelLayer.contains(target)) return false;
      if (this.filterPanel?.panel && this.filterPanel.panel.contains(target)) return false;
      if (this.subtableEditor?.panel && this.subtableEditor.panel.contains(target)) return false;
      if (this.radioPicker?.panel && this.radioPicker.panel.contains(target)) return false;
      if (this.editingCell) return false;

      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();

      const delta = this.getArrowDelta(event.key);
      if (!delta) return true;
      const origin = this.getNavigationOrigin(input);
      if (!origin) return true;
      const next = this.computeNextCoords(origin.row, origin.col, delta.rowDelta, delta.colDelta);
      if (!next) return true;
      if (event.shiftKey && !this.selection) {
        this.selection = {
          anchorRow: origin.row,
          anchorCol: origin.col,
          startRow: origin.row,
          endRow: origin.row,
          startCol: origin.col,
          endCol: origin.col
        };
      }
      this.setSelection(next.row, next.col, { extend: event.shiftKey, focus: true });
      return true;
    }

    computeNextCoords(rowIndex, colIndex, rowDelta, colDelta) {
      let nextRow = rowIndex + rowDelta;
      let nextCol = colIndex + colDelta;

      const rowCount = this.getVisibleRowCount();
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
      return { row: nextRow, col: nextCol };
    }

    moveEditingFocus(rowDelta, colDelta) {
      const active = document.activeElement;
      if (!(active instanceof HTMLInputElement)) return;
      this.onInputChanged(active);
      const prevRow = Number(active.dataset.rowIndex || '0');
      const prevCol = Number(active.dataset.colIndex || '0');
      const previousEditingCell = this.editingCell
        ? {
          ...this.editingCell,
          beforeValue: this.deepClone(this.editingCell.beforeValue)
        }
        : null;
      const next = this.computeNextCoords(prevRow, prevCol, rowDelta, colDelta);
      if (!next) return;
      const prevInput = this.getInput(prevRow, prevCol) || active;
      this.setInputEditingVisual(prevInput, false);
      this.commitCellEditHistory(previousEditingCell);
      this.editingCell = this.createEditingState(next.row, next.col, false);
      const nextInput = this.getInput(next.row, next.col);
      this.setInputEditingVisual(nextInput, true);
      this.setSelectionSingle(next.row, next.col);
      this.pendingEdit = { rowIndex: next.row, colIndex: next.col };
      this.armedCell = { rowIndex: next.row, colIndex: next.col };
      this.repaintSelection();
      this.navigating = true;
      try {
        this.focusCell(next.row, next.col);
      } finally {
        setTimeout(() => {
          this.navigating = false;
        }, 0);
      }
    }

    moveFocus(currentInput, rowDelta, colDelta) {
      const rowIndex = Number(currentInput.dataset.rowIndex || '0');
      const colIndex = Number(currentInput.dataset.colIndex || '0');
      const next = this.computeNextCoords(rowIndex, colIndex, rowDelta, colDelta);
      if (!next) return null;
      const { row: nextRow, col: nextCol } = next;
      this.focusCell(nextRow, nextCol);
      return { row: nextRow, col: nextCol };
    }

    getArrowDelta(key) {
      if (key === 'ArrowUp') return { rowDelta: -1, colDelta: 0 };
      if (key === 'ArrowDown') return { rowDelta: 1, colDelta: 0 };
      if (key === 'ArrowLeft') return { rowDelta: 0, colDelta: -1 };
      if (key === 'ArrowRight') return { rowDelta: 0, colDelta: 1 };
      return null;
    }

    getNavigationOrigin(input = null) {
      if (input instanceof HTMLInputElement) {
        const row = Number(input.dataset.rowIndex || '0');
        const col = Number(input.dataset.colIndex || '0');
        if (Number.isFinite(row) && Number.isFinite(col)) {
          return { row, col };
        }
      }
      if (this.editingCell) {
        return { row: this.editingCell.rowIndex, col: this.editingCell.colIndex };
      }
      if (this.armedCell) {
        return { row: this.armedCell.rowIndex, col: this.armedCell.colIndex };
      }
      if (this.selection) {
        return { row: this.selection.endRow, col: this.selection.endCol };
      }
      if (this.getVisibleRowCount() > 0 && this.fields.length > 0) {
        return { row: 0, col: 0 };
      }
      return null;
    }

    navigateSelectionByArrow(key, extend = false, input = null) {
      const delta = this.getArrowDelta(key);
      if (!delta) return null;
      const origin = this.getNavigationOrigin(input);
      if (!origin) return null;
      const next = this.computeNextCoords(origin.row, origin.col, delta.rowDelta, delta.colDelta);
      if (!next) return null;

      if (this.editingCell) {
        this.exitEditMode();
      }

      this.setSelection(next.row, next.col, { extend, focus: true });
      return next;
    }

    hasUnsavedChanges() {
      if (this.pendingDeletes.size > 0) return true;
      if (this.rows.some((row) => this.isNewRowId(row?.id))) return true;
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

    isNewRowId(recordId) {
      return String(recordId || '').startsWith('NEW:');
    }

    toRecordPayloadValue(field, value) {
      if (this.isSubtableField(field)) {
        const rows = Array.isArray(value) ? value : [];
        return rows.map((row) => ({
          id: row?.id ? String(row.id) : undefined,
          value: row && typeof row.value === 'object' ? JSON.parse(JSON.stringify(row.value)) : {}
        })).map((row) => {
          if (!row.id) delete row.id;
          return row;
        });
      }
      return value ?? '';
    }

    buildRecordPayload(recordId) {
      const fields = this.diff.get(recordId);
      if (!fields) return null;
      const record = {};
      Object.keys(fields).forEach((fieldCode) => {
        const field = this.fieldMap.get(fieldCode);
        if (!field) return;
        const item = fields[fieldCode];
        if (!item) return;
        if (item.type === 'SUBTABLE') {
          record[fieldCode] = {
            value: this.toRecordPayloadValue(field, item.value)
          };
          return;
        }
        if (!this.permissionService.canEditCell(recordId, fieldCode)) return;
        record[fieldCode] = { value: this.toRecordPayloadValue(field, item.value) };
      });
      return Object.keys(record).length ? record : null;
    }

    createPutBatches() {
      const entries = [];
      Array.from(this.diff.keys()).forEach((recordId) => {
        if (this.isNewRowId(recordId)) return;
        if (this.pendingDeletes.has(recordId)) return;
        const record = this.buildRecordPayload(recordId);
        if (!record) return;
        const revision = this.revisionMap.get(recordId);
        const item = { id: recordId, record };
        if (revision !== undefined) item.revision = revision;
        entries.push(item);
      });
      const batches = [];
      for (let i = 0; i < entries.length; i += 50) {
        batches.push(entries.slice(i, i + 50));
      }
      return batches;
    }

    createPostBatches() {
      const entries = [];
      Array.from(this.diff.keys()).forEach((recordId) => {
        if (!this.isNewRowId(recordId)) return;
        if (this.pendingDeletes.has(recordId)) return;
        const record = this.buildRecordPayload(recordId);
        if (!record) return;
        entries.push({ tempId: recordId, record });
      });
      const batches = [];
      for (let i = 0; i < entries.length; i += 50) {
        batches.push(entries.slice(i, i + 50));
      }
      return batches;
    }

    createDeleteBatches() {
      const ids = Array.from(this.pendingDeletes).filter((id) => !this.isNewRowId(id));
      const batches = [];
      for (let i = 0; i < ids.length; i += 50) {
        batches.push(ids.slice(i, i + 50));
      }
      return batches;
    }

    isRequiredCheckTargetField(field) {
      if (!field || !field.required) return false;
      if (this.isSubtableField(field)) return false;
      return this.permissionService.canEditCell('NEW:required-check', String(field.code || ''));
    }

    isMissingRequiredValue(field, value) {
      if (!this.isRequiredCheckTargetField(field)) return false;
      if (field.type === 'NUMBER') {
        return String(value ?? '').trim() === '';
      }
      return String(value ?? '').trim() === '';
    }

    hasRequiredMissingInInvalidSet() {
      for (const key of this.invalidCells) {
        const sep = key.lastIndexOf(':');
        if (sep <= 0) continue;
        const recordId = key.slice(0, sep);
        const fieldCode = key.slice(sep + 1);
        if (!this.isNewRowId(recordId)) continue;
        const field = this.fieldMap.get(fieldCode);
        const row = this.rowMap.get(recordId);
        if (!field || !row) continue;
        if (this.isMissingRequiredValue(field, row.values[fieldCode])) {
          return true;
        }
      }
      return false;
    }

    validateRequiredForNewRows(postBatches) {
      if (!Array.isArray(postBatches) || !postBatches.length) return true;
      const ids = Array.from(new Set(
        postBatches.flatMap((batch) => (Array.isArray(batch) ? batch.map((entry) => String(entry?.tempId || '')) : []))
      )).filter(Boolean);
      if (!ids.length) return true;

      let firstErrorCell = null;
      let hasMissing = false;
      ids.forEach((recordId) => {
        const row = this.rowMap.get(recordId);
        if (!row) return;
        this.fields.forEach((field) => {
          if (!this.isRequiredCheckTargetField(field)) return;
          const key = `${recordId}:${field.code}`;
          const missing = this.isMissingRequiredValue(field, row.values[field.code]);
          const input = this.findInput(recordId, field.code);
          if (missing) {
            hasMissing = true;
            this.invalidCells.add(key);
            this.invalidValueCache.delete(key);
            if (input?.parentElement) {
              input.parentElement.classList.add('pb-overlay__cell--error');
              input.parentElement.dataset.errorReason = 'required';
            }
            if (!firstErrorCell) {
              const rowIndex = this.getVisibleRowIndexById(recordId);
              const colIndex = this.fieldIndexMap.get(field.code);
              if (rowIndex >= 0 && colIndex !== undefined) {
                firstErrorCell = { rowIndex, colIndex };
              }
            }
          } else {
            this.invalidCells.delete(key);
            this.invalidValueCache.delete(key);
            if (input?.parentElement) {
              input.parentElement.classList.remove('pb-overlay__cell--error');
              if (input.parentElement.dataset.errorReason === 'required') {
                delete input.parentElement.dataset.errorReason;
              }
            }
          }
        });
      });

      if (!hasMissing) return true;
      this.updateVirtualRows(true);
      this.updateDirtyBadge();
      this.updateStats();
      this.notify(resolveText(this.language, 'toastRequiredMissing'));
      if (firstErrorCell) {
        this.setSelectionSingle(firstErrorCell.rowIndex, firstErrorCell.colIndex);
        this.focusCell(firstErrorCell.rowIndex, firstErrorCell.colIndex);
      }
      return false;
    }

    markBatchError(batch) {
      if (!Array.isArray(batch)) return;
      batch.forEach((entry) => {
        const rowId = String(entry?.id || entry?.tempId || '');
        if (!rowId) return;
        const row = this.rowMap.get(rowId);
        if (!row) return;
        this.fields.forEach((field) => {
          const input = this.findInput(rowId, field.code);
          if (!input || !input.parentElement) return;
          input.parentElement.classList.add('pb-overlay__cell--error');
        });
      });
    }

    setSaving(flag) {
      const saving = Boolean(flag);
      this.saving = saving;

      if (saving) {
        this.exitEditMode();
        this.closeRadioPicker();
      }

      if (this.saveButton) {
        this.saveButton.disabled = saving;
        this.saveButton.classList.toggle('is-saving', saving);
      }
      this.updatePermissionUiState();
      if (this.closeButton) this.closeButton.disabled = saving;
      if (this.columnButton) {
        this.columnButton.disabled = saving || !this.fields.length;
      }
      this.updateHistoryButtons();

      if (this.saveButtonLabel) {
        this.saveButtonLabel.textContent = resolveText(this.language, saving ? 'btnSaving' : 'btnSave');
      }
      if (this.saveButtonSpinner) {
        this.saveButtonSpinner.classList.toggle('pb-overlay__spinner--hidden', !saving);
      }

      this.inputsByRow.forEach((inputs) => {
        inputs.forEach((input) => {
          if (!input) return;
          input.disabled = saving;
          if (saving) input.readOnly = true;
        });
      });

      if (!saving) {
        this.updateVirtualRows(true);
        this.updateDirtyBadge();
      }
    }

    getRefetchFieldCodes() {
      const codes = new Set(
        this.fields
          .map((field) => String(field?.code || '').trim())
          .filter(Boolean)
      );
      this.fieldsMeta.forEach((field) => {
        if (!field || field.type !== 'SUBTABLE') return;
        const code = String(field.code || '').trim();
        if (code) codes.add(code);
      });
      return Array.from(codes);
    }

    async refetchRowsByIds(ids) {
      const uniqueIds = Array.from(new Set(
        (ids || [])
          .map((id) => String(id || '').trim())
          .filter((id) => id && !this.isNewRowId(id))
      ));
      if (!uniqueIds.length || !this.appId) return;

      const fields = this.getRefetchFieldCodes();
      const fetchedIds = new Set();
      const chunkSize = 100;

      for (let i = 0; i < uniqueIds.length; i += chunkSize) {
        const chunk = uniqueIds.slice(i, i + chunkSize);
        const response = await this.postFn('EXCEL_GET_RECORDS_BY_IDS', {
          appId: this.appId,
          ids: chunk,
          fields
        });
        if (!response?.ok) {
          console.warn('[kintone-excel-overlay] refetchRowsByIds failed', {
            ids: chunk,
            error: response?.error || 'unknown error'
          });
          continue;
        }
        const records = Array.isArray(response.records) ? response.records : [];
        records.forEach((record) => {
          const id = String(record?.$id?.value || '').trim();
          if (id) fetchedIds.add(id);
          this.refreshRowFromRemote(record);
        });
      }

      const missingIds = uniqueIds.filter((id) => !fetchedIds.has(id));
      if (missingIds.length) {
        console.warn('[kintone-excel-overlay] refetchRowsByIds returned partial records', { missingIds });
      }
    }

    async save() {
      if (!this.canMutateOverlay(true)) return;
      if (this.saving) return;
      if (this.invalidCells.size > 0) {
        const key = this.hasRequiredMissingInInvalidSet() ? 'toastRequiredMissing' : 'toastInvalidCells';
        this.notify(resolveText(this.language, key));
        return;
      }
      let putBatches = this.createPutBatches();
      let postBatches = this.createPostBatches();
      let deleteBatches = this.createDeleteBatches();
      if (!putBatches.length && !postBatches.length && !deleteBatches.length) {
        this.notify(resolveText(this.language, 'toastNoChanges'));
        return;
      }
      if (!this.validateRequiredForNewRows(postBatches)) {
        return;
      }
      this.setSaving(true);
      let anySaved = false;
      const savedIds = new Set();
      try {
        for (let index = 0; index < putBatches.length;) {
          let batch = putBatches[index];
          let attempt = 0;
          while (attempt < 2) {
            attempt += 1;
            const response = await this.postFn('EXCEL_PUT_RECORDS', {
              appId: this.appId,
              records: batch
            });
            if (response?.ok) {
              this.applySaveResult(batch, response.result);
              batch.forEach((entry) => {
                const id = String(entry?.id || '').trim();
                if (id) savedIds.add(id);
              });
              anySaved = true;
              break;
            }
            if (attempt === 1 && this.isConflictResponse(response)) {
              const handling = await this.handleConflict(batch);
              if (handling === 'retry') {
                this.notify(resolveText(this.language, 'conflictRetry'));
                putBatches = this.createPutBatches();
                postBatches = this.createPostBatches();
                deleteBatches = this.createDeleteBatches();
                if (!putBatches.length && !postBatches.length && !deleteBatches.length) {
                  this.updateDirtyBadge();
                  this.updateStats();
                  this.notify(resolveText(this.language, 'conflictNoChanges'));
                  return;
                }
                index = 0;
                batch = putBatches[index];
                attempt = 0;
                continue;
              }
              if (handling === 'cleared') {
                this.updateDirtyBadge();
                this.updateStats();
                this.notify(resolveText(this.language, 'conflictNoChanges'));
                return;
              }
            }
            const err = new Error(response?.error || 'save failed');
            this.markBatchError(batch);
            if (this.isConflictResponse(response)) {
              err.conflict = true;
            }
            throw err;
          }
          index += 1;
        }
        for (const batch of postBatches) {
          const response = await this.postFn('EXCEL_POST_RECORDS', {
            appId: this.appId,
            records: batch.map((entry) => ({ ...entry.record }))
          });
          if (!response?.ok) {
            this.markBatchError(batch);
            throw new Error(response?.error || 'create failed');
          }
          const createdIds = this.applyCreateResult(batch, response.result);
          createdIds.forEach((id) => {
            const normalized = String(id || '').trim();
            if (normalized) savedIds.add(normalized);
          });
          anySaved = true;
        }
        for (const batch of deleteBatches) {
          const response = await this.postFn('EXCEL_DELETE_RECORDS', {
            appId: this.appId,
            ids: batch
          });
          if (!response?.ok) {
            throw new Error(response?.error || 'delete failed');
          }
          this.applyDeleteResult(batch);
          anySaved = true;
        }
        if (anySaved) {
          if (savedIds.size) {
            await this.refetchRowsByIds(Array.from(savedIds));
          }
          this.notify(resolveText(this.language, 'toastSaveSuccess'));
          this.diff.clear();
          this.pendingDeletes.clear();
          this.syncPermissionServicePendingDeletes();
          this.clearHistory();
          this.updateDirtyBadge();
          this.updateStats();
          this.requireReload = false;
        }
      } catch (error) {
        const key = error?.conflict ? 'conflictFailed' : 'toastSaveFailed';
        this.notify(resolveText(this.language, key));
        console.error('[kintone-excel-overlay] save failed', error);
      } finally {
        this.setSaving(false);
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
          const field = this.fieldMap.get(fieldCode);
          const value = this.cloneFieldValue(field, entry.record[fieldCode].value);
          row.original[fieldCode] = this.cloneFieldValue(field, value);
          row.values[fieldCode] = this.cloneFieldValue(field, value);
          const input = this.findInput(entry.id, fieldCode);
          if (input) {
            input.dataset.originalValue = this.formatCellDisplayValue(field, row.original[fieldCode]);
            input.value = this.formatCellDisplayValue(field, row.values[fieldCode]);
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

    applyCreateResult(batch, result) {
      const ids = Array.isArray(result?.ids) ? result.ids : [];
      const revisions = Array.isArray(result?.revisions) ? result.revisions : [];
      const createdIds = [];
      batch.forEach((entry, index) => {
        const tempId = entry.tempId;
        const newId = ids[index] != null ? String(ids[index]) : '';
        if (!tempId || !newId) return;
        createdIds.push(newId);
        const row = this.rowMap.get(tempId);
        if (!row) return;
        const newRevision = revisions[index] != null ? String(revisions[index]) : undefined;
        Object.keys(entry.record || {}).forEach((fieldCode) => {
          const field = this.fieldMap.get(fieldCode);
          const value = this.cloneFieldValue(field, entry.record[fieldCode]?.value);
          row.original[fieldCode] = this.cloneFieldValue(field, value);
          row.values[fieldCode] = this.cloneFieldValue(field, value);
        });
        row.id = newId;
        row.isNew = false;
        if (newRevision !== undefined) row.revision = newRevision;

        const rowIndex = this.rowIndexMap.get(tempId);
        this.rowMap.delete(tempId);
        this.rowMap.set(newId, row);
        if (rowIndex !== undefined) {
          this.rowIndexMap.delete(tempId);
          this.rowIndexMap.set(newId, rowIndex);
        }
        this.revisionMap.delete(tempId);
        if (newRevision !== undefined) this.revisionMap.set(newId, newRevision);
        this.diff.delete(tempId);
        this.clearInvalidStateForRecord(tempId);
        this.permissionService.removeRecordAcl(tempId);
        this.permissionService.upsertRecordAcl({
          id: newId,
          record: { viewable: true, editable: true, deletable: true },
          fields: {}
        });

        this.fields.forEach((field) => {
          const input = this.findInput(newId, field.code);
          if (!input) return;
          input.dataset.recordId = newId;
          input.dataset.originalValue = this.formatCellDisplayValue(field, row.original[field.code]);
          this.applyCellVisualState(input, row, field.code);
        });
      });
      this.recomputeFilteredRowIds();
      this.updateVirtualRows(true);
      return createdIds;
    }

    applyDeleteResult(ids) {
      const deleteIds = new Set(
        (ids || [])
          .map((id) => String(id || '').trim())
          .filter((id) => id && !this.isNewRowId(id))
      );
      if (!deleteIds.size) return;

      if (this.editingCell) {
        const editingRowId = String(this.getVisibleRowAt(this.editingCell.rowIndex)?.id || '').trim();
        if (editingRowId && deleteIds.has(editingRowId)) {
          this.exitEditMode();
        }
      }

      this.rows = this.rows.filter((row) => !deleteIds.has(String(row?.id || '').trim()));
      deleteIds.forEach((id) => {
        this.pendingDeletes.delete(id);
        this.diff.delete(id);
        this.rowMap.delete(id);
        this.rowIndexMap.delete(id);
        this.revisionMap.delete(id);
        this.permissionService.removeRecordAcl(id);
        this.clearInvalidStateForRecord(id);
      });
      this.syncPermissionServicePendingDeletes();

      this.rebuildRowMaps();
      this.recomputeFilteredRowIds();
      const visibleCount = this.getVisibleRowCount();
      if (!visibleCount) {
        this.selection = null;
        this.armedCell = null;
      } else if (this.selection) {
        const nextRow = Math.max(0, Math.min(visibleCount - 1, this.selection.startRow || 0));
        const nextCol = Math.max(0, Math.min(this.fields.length - 1, this.selection.startCol || 0));
        this.setSelectionSingle(nextRow, nextCol);
      }
      this.updateVirtualRows(true);
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
      this.detachUiLanguageListener();
      this.closeColumnManager(true);
      this.closeFilterPanel();
      this.closeRadioPicker();
      this.closeSubtableEditor();
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
      this.permissionWarningEl = null;
      this.dirtyBadge = null;
      this.saveButton = null;
      this.saveButtonSpinner = null;
      this.saveButtonLabel = null;
      this.addRowButton = null;
      this.undoButton = null;
      this.redoButton = null;
      this.prevPageButton = null;
      this.nextPageButton = null;
      this.pageLabelElement = null;
      this.closeButton = null;
      this.titleElement = null;
      this.modeBadge = null;
      this.appName = '';
      this.baseQuery = '';
      this.query = '';
      this.pageOffset = 0;
      this.totalCount = 0;
      this.paging = false;
      this.statsElements = {
        sum: null,
        avg: null,
        count: null,
        selected: null,
        filtered: null,
        pendingDelete: null,
        newRows: null
      };
      this.fields = [];
      this.fieldsMeta = [];
      this.fieldMap.clear();
      this.fieldsMetaMap.clear();
      this.fieldIndexMap.clear();
      this.rows = [];
      this.rowMap.clear();
      this.rowIndexMap.clear();
      this.permissionService.clearRecordAcl();
      this.aclStatus = { loaded: false, failed: false };
      this.pendingDeletes.clear();
      this.syncPermissionServicePendingDeletes();
      this.filters.clear();
      this.filteredRowIds = null;
      this.filterPanel = null;
      this.syncPermissionServiceFields();
      this.permissionService.setAppPermission({
        editable: true,
        addable: true,
        deletable: true,
        source: 'local'
      });
      this.inputsByRow.clear();
      this.diff.clear();
      this.undoStack = [];
      this.redoStack = [];
      this.isReplayingHistory = false;
      this.invalidCells.clear();
      this.revisionMap.clear();
      this.invalidValueCache.clear();
      this.rowHeaderContainer = null;
      this.tableContainer = null;
      this.virtualStart = 0;
      this.virtualEnd = 0;
      this.pendingFocus = null;
      this.pendingEdit = null;
      this.pasteFlashTimer = null;
      this.needsFilterReapply = false;
      this.sortState = { fieldCode: '', direction: '' };
      this.editingCell = null;
      this.viewOnlyNoticeShown = false;
      this.selection = null;
      this.armedCell = null;
      this.navigating = false;
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
        input.dataset.originalValue = this.formatCellDisplayValue(field, existing.original[field.code]);
        if (!diffEntry[field.code]) {
          input.value = this.formatCellDisplayValue(field, existing.values[field.code]);
        }
        this.applyCellVisualState(input, existing, field.code);
      });
      if (this.hasActiveFilters()) {
        this.recomputeFilteredRowIds();
      }
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
        const field = this.fieldMap.get(fieldCode);
        if (this.isSubtableField(field)) {
          row.values[fieldCode] = this.cloneFieldValue(field, info.value);
          this.applyCellVisualState(input, row, fieldCode);
          return;
        }
        input.dataset.originalValue = this.formatCellDisplayValue(field, row.original[fieldCode]);
        input.value = this.formatCellDisplayValue(field, info.value ?? '');
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













