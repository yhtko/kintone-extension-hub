// page-bridge.js
// 繝壹・繧ｸ繧ｳ繝ｳ繝・く繧ｹ繝医〒蜍穂ｽ懶ｼ・ kintone.api 縺檎峩謗･菴ｿ縺医ｋ・・
(function () {
  if (window.__kfav_pageBridgeLoaded) return;
  window.__kfav_pageBridgeLoaded = true;

  const ORIGIN = location.origin;
const viewsCache = new Map(); // appId -> views
const fieldsCache = new Map(); // appId -> fields metadata cache

function normalizeChoiceList(prop) {
  if (!prop || typeof prop !== 'object' || !prop.options) return [];
  return Object.values(prop.options)
    .map((opt) => (opt && typeof opt === 'object' && opt.label ? String(opt.label) : ''))
    .filter((label) => label);
}

function normalizeLookupMeta(code, prop) {
  const src = prop && typeof prop === 'object' ? (prop.lookup || null) : null;
  if (!src || typeof src !== 'object') return undefined;
  const relatedAppRaw = src.relatedApp;
  const relatedApp = (typeof relatedAppRaw === 'string' || typeof relatedAppRaw === 'number')
    ? String(relatedAppRaw)
    : (relatedAppRaw && typeof relatedAppRaw === 'object'
      ? String(relatedAppRaw.app || relatedAppRaw.id || '')
      : '');
  const keyField = String(
    src.keyField
    || src.lookupField
    || src.relatedKeyField
    || src.fieldCode
    || ''
  ).trim();
  const out = {};
  if (keyField) out.keyField = keyField;
  if (relatedApp) out.relatedApp = relatedApp;
  return Object.keys(out).length ? out : undefined;
}

function normalizeSubtableMeta(prop) {
  const fieldsObj = prop && typeof prop === 'object' && prop.fields && typeof prop.fields === 'object'
    ? prop.fields
    : null;
  if (!fieldsObj) return undefined;
  const fields = Object.keys(fieldsObj).map((code) => {
    const child = fieldsObj[code] || {};
    return {
      code,
      label: child.label || '',
      type: child.type || '',
      required: Boolean(child.required),
      choices: normalizeChoiceList(child)
    };
  });
  return { fields };
}

function normalizeFieldMetaFromProperties(properties) {
  if (!properties || typeof properties !== 'object') return [];
  const metas = Object.keys(properties).map((code) => {
    const prop = properties[code] || {};
    const meta = {
      code,
      label: prop.label || '',
      type: prop.type || '',
      required: Boolean(prop.required),
      choices: normalizeChoiceList(prop)
    };
    const lookup = normalizeLookupMeta(code, prop);
    if (lookup) meta.lookup = lookup;
    if (meta.type === 'SUBTABLE') {
      const subtable = normalizeSubtableMeta(prop);
      if (subtable) meta.subtable = subtable;
    }
    return meta;
  });
  markLookupAutoFields(properties, metas);
  return metas;
}

function collectLookupRelatedCodes(value, knownCodes, out) {
  if (!value) return;
  if (typeof value === 'string') {
    if (knownCodes.has(value)) out.add(value);
    return;
  }
  if (Array.isArray(value)) {
    value.forEach((item) => collectLookupRelatedCodes(item, knownCodes, out));
    return;
  }
  if (typeof value === 'object') {
    Object.values(value).forEach((item) => collectLookupRelatedCodes(item, knownCodes, out));
  }
}

function markLookupAutoFields(properties, metas) {
  const knownCodes = new Set(Object.keys(properties || {}));
  if (!knownCodes.size) return;
  const metaMap = new Map((metas || []).map((meta) => [meta.code, meta]));
  Object.keys(properties || {}).forEach((sourceCode) => {
    const prop = properties[sourceCode] || {};
    const lookup = prop.lookup;
    if (!lookup || typeof lookup !== 'object') return;
    const keyField = String(
      lookup.keyField
      || lookup.lookupField
      || lookup.relatedKeyField
      || lookup.fieldCode
      || sourceCode
      || ''
    ).trim();
    const related = new Set();
    collectLookupRelatedCodes(lookup, knownCodes, related);
    related.delete(sourceCode);
    if (keyField) related.delete(keyField);
    related.forEach((targetCode) => {
      const target = metaMap.get(targetCode);
      if (!target) return;
      target.lookupAuto = true;
      if (!target.lookup || !target.lookup.keyField) {
        target.lookup = { ...(target.lookup || {}), keyField };
      }
    });
  });
}

  async function getAppName(appId) {
    if (!appId) return '';
    const sdk = window.kintone;
    if (!sdk?.api) return '';
    try {
      const resp = await sdk.api(sdk.api.url('/k/v1/app.json', true), 'GET', { id: String(appId) });
      if (resp?.app?.name) return resp.app.name;
      if (resp?.name) return resp.name;
    } catch (_e) { /* ignore */ }

    const selectors = [
      '.gaia-argoui-app-index-appname span',
      '.gaia-argoui-app-index-toolbar-title span',
      '.gaia-argoui-app-title span',
      '[data-test-id="app-title-text"]',
      '[data-test-id="app-title"]'
    ];
    for (const selector of selectors) {
      try {
        const el = document.querySelector(selector);
        if (el && el.textContent) {
          const text = el.textContent.trim();
          if (text) return text;
        }
      } catch (_e) { /* ignore */ }
    }

    const docTitle = document.title || '';
    if (docTitle) {
      const parts = docTitle.split(' - ');
      if (parts.length) return parts[0].trim();
      return docTitle.trim();
    }

    return '';
  }

  async function getViews(appId) {
    const resp = await window.kintone.api(window.kintone.api.url('/k/v1/app/views', true), 'GET', { app: String(appId) });
    return resp.views || {};
  }

  async function getViewsCached(appId) {
    const key = String(appId);
    if (!viewsCache.has(key)) {
      viewsCache.set(key, await getViews(appId));
    }
    return viewsCache.get(key) || {};
  }

  async function getFields(appId) {
    const key = String(appId || '');
    if (!key) return [];
    if (fieldsCache.has(key)) {
      return fieldsCache.get(key).map((field) => ({ ...field }));
    }

    const fields = await fetchFieldsWithFallback(key);
    fieldsCache.set(key, fields);
    return fields.map((field) => ({ ...field }));
  }

  async function getFieldsMeta(appId) {
    const key = String(appId || '');
    if (!key) return [];
    const sdk = window.kintone;
    if (!sdk?.api) return [];
    try {
      const resp = await sdk.api(sdk.api.url('/k/v1/app/form/fields', true), 'GET', { app: key });
      return normalizeFieldMetaFromProperties(resp?.properties || {});
    } catch (error) {
      if (typeof console !== 'undefined' && console.warn) {
        console.warn('[kintone-favorites] getFieldsMeta failed via form API, falling back to getFields', error);
      }
      const fallback = await getFields(key);
      return fallback.map((field) => ({
        code: field.code,
        label: field.label || '',
        type: field.type || '',
        required: false,
        choices: Array.isArray(field.choices) ? field.choices : []
      }));
    }
  }

  async function fetchFieldsWithFallback(appId) {
    const sdk = window.kintone;
    if (!sdk?.api) return [];
    try {
      // 繧｢繝励Μ縺ｮ繝輔ぅ繝ｼ繝ｫ繝牙ｮ夂ｾｩ繧貞叙蠕暦ｼ域ｨｩ髯舌′縺ｪ縺・ｴ蜷医・萓句､問・蜻ｼ縺ｳ蜃ｺ縺怜・縺ｧ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ・・
      const resp = await sdk.api(sdk.api.url('/k/v1/app/form/fields', true), 'GET', { app: String(appId) });
      // 霑泌唆: { code, label, type } 縺ｮ驟榊・
      const props = resp?.properties || {};
      return Object.keys(props).map((code) => ({
        code,
        label: props[code].label || '',
        type: props[code].type || '',
        choices: normalizeChoiceList(props[code])
      }));
    } catch (error) {
      if (typeof console !== 'undefined' && console.warn) {
        console.warn('[kintone-favorites] getFields failed via form API, falling back to records API', error);
      }
      const fallback = await fetchFieldsFromRecords(appId);
      if (fallback.length) return fallback;
      throw error;
    }
  }

  function collectFieldLabelsFromDom() {
    const labels = new Map();
    const selectors = [
      'th[data-field-code]',
      '[role="columnheader"][data-field-code]'
    ];
    for (const selector of selectors) {
      try {
        const nodes = document.querySelectorAll(selector);
        nodes.forEach((node) => {
          const code = node?.getAttribute?.('data-field-code');
          if (!code || labels.has(code)) return;
          const text = (node.textContent || '').trim();
          if (text) labels.set(code, text);
        });
      } catch (_e) {
        // ignore selector errors
      }
    }
    return labels;
  }

  async function fetchFieldsFromRecords(appId) {
    const sdk = window.kintone;
    if (!sdk?.api) return [];
    try {
      const params = {
        app: String(appId),
        query: 'limit 1',
        totalCount: false
      };
      const resp = await sdk.api(sdk.api.url('/k/v1/records', true), 'GET', params);
      const firstRecord = Array.isArray(resp?.records) && resp.records.length ? resp.records[0] : null;
      if (!firstRecord) return [];
      const labels = collectFieldLabelsFromDom();
      const fields = [];
      for (const [code, meta] of Object.entries(firstRecord)) {
        if (!code || code.startsWith('$')) continue;
        if (!meta || typeof meta !== 'object' || !meta.type) continue;
        fields.push({
          code,
          label: labels.get(code) || '',
          type: meta.type,
          choices: []
        });
      }
      return fields;
    } catch (fallbackError) {
      if (typeof console !== 'undefined' && console.warn) {
        console.warn('[kintone-favorites] getFields fallback via records API failed', fallbackError);
      }
      return [];
    }
  }

  const EDITABLE_TYPES = new Set(['SINGLE_LINE_TEXT', 'NUMBER', 'DATE', 'RADIO_BUTTON', 'DROP_DOWN']);

  function extractQueryFromViews(viewsObj, viewIdOrName) {
    const entries = Object.entries(viewsObj || {});
    const byId = entries.find(([, v]) => String(v.id) === String(viewIdOrName));
    if (byId) return (byId[1].filterCond || '') + (byId[1].sort ? ` order by ${byId[1].sort}` : '');
    const byName = entries.find(([name]) => name === viewIdOrName);
    if (byName) {
      const v = byName[1];
      return (v.filterCond || '') + (v.sort ? ` order by ${v.sort}` : '');
    }

    return '';
  }

  function filterCondFromViews(viewsObj, viewIdOrName) {
    const entries = Object.entries(viewsObj || {});
    const byId = entries.find(([, v]) => String(v.id) === String(viewIdOrName));
    if (byId) return byId[1].filterCond || '';
    const byName = entries.find(([name]) => name === viewIdOrName);
    if (byName) return byName[1].filterCond || '';
    return '';
  }

  function combineQuery(base, extra) {
    const A = (base || '').trim();
    const B = (extra || '').trim();
    if (A && B) return `${A} and (${B})`;
    return A || B || '';
  }

  function splitQueryParts(query) {
    const original = String(query || '');
    const lower = original.toLowerCase();
    const markers = [];
    const orderIndex = lower.indexOf(' order by ');
    if (orderIndex >= 0) markers.push(orderIndex);
    const limitIndex = lower.indexOf(' limit ');
    if (limitIndex >= 0) markers.push(limitIndex);
    const offsetIndex = lower.indexOf(' offset ');
    if (offsetIndex >= 0) markers.push(offsetIndex);
    if (!markers.length) {
      return { filter: original.trim(), trailing: '' };
    }
    const first = Math.min(...markers);
    const filter = original.slice(0, first).trim();
    const trailing = original.slice(first).trim();
    return { filter, trailing };
  }

  function hasOrderClause(query) {
    return /\border\s+by\b/i.test(query || '');
  }

  function extractColumnOrder(view) {
    const order = [];
    if (!view) return order;
    if (Array.isArray(view.columns)) {
      view.columns.forEach((col) => {
        if (!col) return;
        let code = null;
        if (typeof col === 'string') {
          code = col;
        } else if (col && typeof col === 'object') {
          if (col.field) code = col.field;
          else if (col.code) code = col.code;
          else if (col.id) code = col.id;
        }
        if (code && !order.includes(code)) order.push(code);
      });
    }
    if (!order.length && Array.isArray(view.fields)) {
      view.fields.forEach((col) => {
        if (!col) return;
        let code = null;
        if (typeof col === 'string') {
          code = col;
        } else if (col && typeof col === 'object') {
          if (col.field) code = col.field;
          else if (col.code) code = col.code;
          else if (col.id) code = col.id;
        }
        if (code && !order.includes(code)) order.push(code);
      });
    }
    if (!order.length && Array.isArray(view.layout)) {
      view.layout.forEach((row) => {
        if (!row || typeof row !== 'object') return;
        if (Array.isArray(row.fields)) {
          row.fields.forEach((field) => {
            if (!field) return;
            let code = null;
            if (typeof field === 'string') {
              code = field;
            } else if (field && typeof field === 'object') {
              if (field.field) code = field.field;
              else if (field.code) code = field.code;
              else if (field.id) code = field.id;
            }
            if (code && !order.includes(code)) order.push(code);
          });
        }
      });
    }
    return order;
  }

  async function countRecords(appId, query) {
    const params = { app: String(appId), totalCount: true, size: 1 };
    if (query && query.trim()) params.query = query;
    const resp = await window.kintone.api(window.kintone.api.url('/k/v1/records', true), 'GET', params);
    return parseInt(resp.totalCount || '0', 10);
  }

  function normalizePermissionState(value, fallback = 'unknown') {
    if (value === 'allow' || value === 'deny' || value === 'unknown') return value;
    if (typeof value === 'boolean') return value ? 'allow' : 'deny';
    if (typeof value === 'number') return value !== 0 ? 'allow' : 'deny';
    if (typeof value === 'string') {
      const normalized = value.trim().toLowerCase();
      if (!normalized) return fallback;
      if (['allow', 'allowed', 'true', '1', 'yes', 'read', 'write', 'grant', 'granted'].includes(normalized)) return 'allow';
      if (['deny', 'denied', 'false', '0', 'no', 'none', 'forbidden'].includes(normalized)) return 'deny';
      if (['unknown', 'unresolved'].includes(normalized)) return 'unknown';
      return fallback;
    }
    return fallback;
  }

  function normalizeRightSet(raw, fallback = { view: 'unknown', add: 'unknown', edit: 'unknown', delete: 'unknown', source: 'unknown' }) {
    const source = raw && typeof raw === 'object' ? raw : {};
    const view = normalizePermissionState(
      source.view ?? source.read ?? source.browse ?? source.canView ?? source.isViewable,
      fallback.view
    );
    const add = normalizePermissionState(
      source.add ?? source.create ?? source.write ?? source.canAdd ?? source.canCreate,
      fallback.add
    );
    const edit = normalizePermissionState(
      source.edit ?? source.update ?? source.write ?? source.canEdit ?? source.canUpdate,
      fallback.edit
    );
    const del = normalizePermissionState(
      source.delete ?? source.remove ?? source.canDelete ?? source.canRemove,
      fallback.delete
    );
    let sourceState = fallback.source || 'unknown';
    if (source.source === 'api' || source.source === 'fallback' || source.source === 'unknown') {
      sourceState = source.source;
    } else if (view !== 'unknown' || add !== 'unknown' || edit !== 'unknown' || del !== 'unknown') {
      sourceState = 'api';
    }
    return {
      view,
      add,
      edit,
      delete: del,
      source: sourceState
    };
  }

  function emptyRecordPermissionMap(ids) {
    const out = {};
    (ids || []).forEach((id) => {
      const key = String(id || '').trim();
      if (!key) return;
      out[key] = { view: 'unknown', edit: 'unknown', delete: 'unknown', source: 'unknown' };
    });
    return out;
  }

  function pickAppPermissionCandidate(resp) {
    if (!resp || typeof resp !== 'object') return null;
    const candidates = [
      resp.permissions,
      resp.permission,
      resp.rights,
      resp.acl,
      resp.appPermissions,
      resp.result
    ];
    for (const candidate of candidates) {
      if (candidate && typeof candidate === 'object' && !Array.isArray(candidate)) {
        return candidate;
      }
    }
    return null;
  }

  function pickRecordPermissionEntries(resp) {
    if (!resp || typeof resp !== 'object') return [];
    const candidates = [
      resp.records,
      resp.permissions,
      resp.results,
      resp.items
    ];
    for (const candidate of candidates) {
      if (Array.isArray(candidate)) return candidate;
      if (candidate && typeof candidate === 'object') {
        if (Array.isArray(candidate.records)) return candidate.records;
        if (Array.isArray(candidate.items)) return candidate.items;
        if (Array.isArray(candidate.results)) return candidate.results;
      }
    }
    return [];
  }

  async function fetchAppPermissions(appId) {
    const fallback = { view: 'unknown', add: 'unknown', edit: 'unknown', delete: 'unknown', source: 'unknown' };
    const sdk = window.kintone;
    if (!sdk?.api) return fallback;
    const app = String(appId || '').trim();
    if (!app) return fallback;

    const attempts = [
      { path: '/k/v1/app/acl/evaluate', method: 'POST', payload: { app } },
      { path: '/k/v1/app/permissions', method: 'GET', payload: { app } },
      { path: '/k/v1/app/acl', method: 'GET', payload: { app } }
    ];

    for (const attempt of attempts) {
      try {
        const resp = await sdk.api(sdk.api.url(attempt.path, true), attempt.method, attempt.payload);
        console.debug('[excel-permission-api]', {
          appId: app,
          type: 'app',
          path: attempt.path,
          result: resp
        });
        const candidate = pickAppPermissionCandidate(resp) || resp;
        return normalizeRightSet(candidate, { view: 'unknown', add: 'unknown', edit: 'unknown', delete: 'unknown', source: 'api' });
      } catch (_err) {
        // try next
      }
    }
    console.debug('[excel-permission-api]', {
      appId: app,
      type: 'app',
      result: 'fallback_unknown'
    });
    return fallback;
  }

  async function fetchRecordPermissions(appId, ids) {
    const idList = (ids || []).map((id) => String(id || '').trim()).filter(Boolean);
    const fallback = emptyRecordPermissionMap(idList);
    const sdk = window.kintone;
    if (!sdk?.api) return fallback;
    const app = String(appId || '').trim();
    if (!app || !idList.length) return fallback;

    const attempts = [
      { path: '/k/v1/records/acl/evaluate', method: 'POST', payload: { app, ids: idList } },
      { path: '/k/v1/records/permissions', method: 'GET', payload: { app, ids: idList } }
    ];

    for (const attempt of attempts) {
      try {
        const resp = await sdk.api(sdk.api.url(attempt.path, true), attempt.method, attempt.payload);
        console.debug('[excel-permission-api]', {
          appId: app,
          type: 'record',
          path: attempt.path,
          ids: idList,
          result: resp
        });
        const entries = pickRecordPermissionEntries(resp);
        if (!entries.length) continue;
        const out = { ...fallback };
        entries.forEach((entry) => {
          const id = String(entry?.id ?? entry?.recordId ?? entry?.record ?? '').trim();
          if (!id) return;
          const raw = entry?.permissions || entry?.permission || entry?.rights || entry;
          const normalized = normalizeRightSet(raw, { view: 'unknown', add: 'unknown', edit: 'unknown', delete: 'unknown', source: 'api' });
          out[id] = {
            view: normalized.view,
            edit: normalized.edit,
            delete: normalized.delete,
            source: normalized.source || 'api'
          };
        });
        return out;
      } catch (_err) {
        // try next
      }
    }
    console.debug('[excel-permission-api]', {
      appId: app,
      type: 'record',
      ids: idList,
      result: 'fallback_unknown'
    });
    return fallback;
  }

  async function fetchEvaluatedRecordAcl(appId, ids) {
    const idList = (ids || []).map((id) => String(id || '').trim()).filter(Boolean);
    if (!idList.length) return [];
    const sdk = window.kintone;
    if (!sdk?.api) return [];
    const app = String(appId || '').trim();
    if (!app) return [];
    const resp = await sdk.api(
      sdk.api.url('/k/v1/records/acl/evaluate.json', true),
      'GET',
      { app, ids: idList }
    );
    console.debug('[excel-permission-api]', {
      appId: app,
      type: 'record_acl_evaluate',
      ids: idList,
      result: resp
    });
    if (Array.isArray(resp?.rights)) return resp.rights;
    if (Array.isArray(resp?.records)) return resp.records;
    return [];
  }

  function tzOffsetStr(d) {
    // "+09:00" 縺ｮ繧医≧縺ｪ陦ｨ險倥ｒ霑斐☆
    const offMin = -d.getTimezoneOffset(); // JS縺ｯ縲袈TC縺九ｉ縺ｮ蟾ｮ縲阪ｒ騾・ｬｦ蜿ｷ縺ｧ霑斐☆
    const sign = offMin >= 0 ? '+' : '-';
    const a = Math.abs(offMin);
    const hh = String(Math.floor(a / 60)).padStart(2, '0');
    const mm = String(a % 60).padStart(2, '0');
    return `${sign}${hh}:${mm}`;
  }

  window.addEventListener('message', async (ev) => {
    if (ev.origin !== ORIGIN) return;
    const msg = ev.data || {};
    if (!msg || msg.__kfav__ !== true) return;

    const { id, type, payload } = msg;

    try {
      if (type === 'EXCEL_GET_APP_CONTEXT') {
        const appId = window.kintone?.app?.getId?.();
        const query = window.kintone?.app?.getQuery?.() || '';
        let appName = '';
        if (appId) {
          appName = await getAppName(appId);
        }
        window.postMessage({ __kfav__: true, replyTo: id, ok: true, appId, query, appName }, ORIGIN);
        return;
      }

      if (type === 'EXCEL_GET_LOGIN_USER') {
        const user = typeof window.kintone?.getLoginUser === 'function' ? window.kintone.getLoginUser() : null;
        if (!user) throw new Error('Could not retrieve login user');
        window.postMessage({
          __kfav__: true,
          replyTo: id,
          ok: true,
          user: { language: user.language || '' }
        }, ORIGIN);
        return;
      }

      if (type === 'CHECK_KINTONE_READY') {
        const ready = typeof window.kintone?.api === 'function';
        let appId = '';
        if (ready && typeof window.kintone?.app?.getId === 'function') {
          appId = window.kintone.app.getId() || '';
        }
        let language = '';
        if (ready && typeof window.kintone?.getLoginUser === 'function') {
          try {
            language = window.kintone.getLoginUser()?.language || '';
          } catch (_e) {}
        }
        window.postMessage({
          __kfav__: true,
          replyTo: id,
          ok: true,
          ready,
          appId,
          language
        }, ORIGIN);
        return;
      }

      if (type === 'EXCEL_GET_FIELDS') {
        const appId = payload?.appId;
        if (!appId) throw new Error('appId is required');
        const fields = await getFields(appId);
        const editable = fields.filter(f => EDITABLE_TYPES.has(f.type));
        window.postMessage({ __kfav__: true, replyTo: id, ok: true, fields: editable }, ORIGIN);
        return;
      }

      if (type === 'EXCEL_GET_FIELDS_META') {
        const appId = payload?.appId;
        if (!appId) throw new Error('appId is required');
        const fieldsMeta = await getFieldsMeta(appId);
        window.postMessage({ __kfav__: true, replyTo: id, ok: true, fieldsMeta }, ORIGIN);
        return;
      }

      if (type === 'EXCEL_GET_RECORDS') {
        const appId = payload?.appId;
        if (!appId) throw new Error('appId is required');
        const fields = Array.isArray(payload?.fields) ? payload.fields.filter(Boolean) : [];
        const params = {
          app: String(appId),
          fields: Array.from(new Set(['$id', '$revision'].concat(fields))),
          totalCount: true
        };
        const query = String(payload?.query || '');
        if (query.trim()) params.query = query.trim();
        const resp = await window.kintone.api(window.kintone.api.url('/k/v1/records', true), 'GET', params);
        window.postMessage({
          __kfav__: true,
          replyTo: id,
          ok: true,
          records: resp.records || [],
          totalCount: typeof resp.totalCount === 'number' ? resp.totalCount : Number(resp.totalCount || 0)
        }, ORIGIN);
        return;
      }

      if (type === 'EXCEL_GET_RECORDS_BY_IDS') {
        const appId = payload?.appId;
        const ids = Array.isArray(payload?.ids) ? payload.ids.filter(Boolean) : [];
        if (!appId) throw new Error('appId is required');
        if (!ids.length) {
          window.postMessage({ __kfav__: true, replyTo: id, ok: true, records: [] }, ORIGIN);
          return;
        }
        const fields = Array.isArray(payload?.fields) ? payload.fields.filter(Boolean) : [];
        const params = {
          app: String(appId),
          ids: ids.map(String),
          fields: Array.from(new Set(['$id', '$revision'].concat(fields)))
        };
        const resp = await window.kintone.api(window.kintone.api.url('/k/v1/records', true), 'GET', params);
        window.postMessage({ __kfav__: true, replyTo: id, ok: true, records: resp.records || [] }, ORIGIN);
        return;
      }

      if (type === 'EXCEL_GET_VIEW_INFO') {
        const appId = payload?.appId;
        if (!appId) throw new Error('appId is required');
        const currentQuery = String(payload?.currentQuery || '');
        const viewId = payload?.viewId ? String(payload.viewId) : '';
        const views = await getViewsCached(appId);
        const entries = Object.entries(views || {});
        let selectedView = null;
        let selectedViewName = '';
        if (viewId) {
          const found = entries.find(([name, v]) => String(v.id) === viewId);
          if (found) {
            selectedView = found[1];
            selectedViewName = found[0];
          }
        }
        if (!selectedView) {
          const currentName = typeof window.kintone?.app?.getViewName === 'function' ? window.kintone.app.getViewName() : '';
          if (currentName && views[currentName]) {
            selectedView = views[currentName];
            selectedViewName = currentName;
          }
        }
        if (!selectedView && entries.length) {
          selectedView = entries[0][1];
          selectedViewName = entries[0][0];
        }

        const baseFilter = selectedView?.filterCond || '';
        const { filter: currentFilter, trailing } = splitQueryParts(currentQuery);
        const mergedFilter = combineQuery(baseFilter, currentFilter);
        const parts = [];
        if (mergedFilter) parts.push(mergedFilter);
        if (trailing) {
          parts.push(trailing);
        } else if (selectedView?.sort && !hasOrderClause(currentQuery)) {
          parts.push(`order by ${selectedView.sort}`);
        }
        let finalQuery = parts.join(' ').trim();
        if (!finalQuery) {
          finalQuery = currentQuery.trim() || baseFilter.trim() || '';
        }

        const fieldOrder = extractColumnOrder(selectedView);

        window.postMessage({
          __kfav__: true,
          replyTo: id,
          ok: true,
          query: finalQuery,
          fieldOrder,
          viewId: selectedView?.id ? String(selectedView.id) : '',
          viewName: selectedViewName || ''
        }, ORIGIN);
        return;
      }

      if (type === 'EXCEL_PUT_RECORDS') {
        const appId = payload?.appId;
        const records = Array.isArray(payload?.records) ? payload.records : null;
        if (!appId || !records) throw new Error('appId and records are required');
        const req = { app: String(appId), records };
        const resp = await window.kintone.api(window.kintone.api.url('/k/v1/records', true), 'PUT', req);
        window.postMessage({ __kfav__: true, replyTo: id, ok: true, result: resp }, ORIGIN);
        return;
      }

      if (type === 'EXCEL_POST_RECORDS') {
        const appId = payload?.appId;
        const records = Array.isArray(payload?.records) ? payload.records : null;
        if (!appId || !records) throw new Error('appId and records are required');
        const req = { app: String(appId), records };
        const resp = await window.kintone.api(window.kintone.api.url('/k/v1/records', true), 'POST', req);
        window.postMessage({ __kfav__: true, replyTo: id, ok: true, result: resp }, ORIGIN);
        return;
      }

      if (type === 'EXCEL_DELETE_RECORDS') {
        const appId = payload?.appId;
        const ids = Array.isArray(payload?.ids) ? payload.ids.map((v) => String(v || '').trim()).filter(Boolean) : [];
        if (!appId) throw new Error('appId is required');
        if (!ids.length) {
          window.postMessage({ __kfav__: true, replyTo: id, ok: true, result: { ids: [] } }, ORIGIN);
          return;
        }
        const req = { app: String(appId), ids };
        const resp = await window.kintone.api(window.kintone.api.url('/k/v1/records', true), 'DELETE', req);
        window.postMessage({ __kfav__: true, replyTo: id, ok: true, result: resp || { ids } }, ORIGIN);
        return;
      }

      if (type === 'EXCEL_EVALUATE_RECORD_ACL') {
        const appId = payload?.appId;
        const ids = Array.isArray(payload?.ids) ? payload.ids : [];
        if (!appId) throw new Error('appId is required');
        const rights = await fetchEvaluatedRecordAcl(appId, ids);
        window.postMessage({ __kfav__: true, replyTo: id, ok: true, rights }, ORIGIN);
        return;
      }

      if (type === 'COUNT_VIEW') {
        const { appId, viewIdOrName } = payload;
        const views = await getViews(appId);
        const query = extractQueryFromViews(views, viewIdOrName);
        const count = await countRecords(appId, query);
        window.postMessage({ __kfav__: true, replyTo: id, ok: true, count }, ORIGIN);
        return;
      }

      if (type === 'COUNT_APP_QUERY') {
        const { appId, query } = payload;
        const count = await countRecords(appId, query || '');
        window.postMessage({ __kfav__: true, replyTo: id, ok: true, count }, ORIGIN);
        return;
      }

      if (type === 'GET_APP_NAME') {
        const appId = payload?.appId;
        if (!appId) throw new Error('appId is required');
        const resp = await window.kintone.api(
          window.kintone.api.url('/k/v1/app.json', true),
          'GET',
          { id: String(appId) }
        );
        window.postMessage({
          __kfav__: true,
          replyTo: id,
          ok: true,
          name: String(resp?.name || '')
        }, ORIGIN);
        return;
      }

      if (type === 'COUNT_BULK') {
        const items = Array.isArray(payload?.items) ? payload.items : [];
        const counts = {};
        const errors = {};
        for (const item of items) {
          const key = item?.id;
          if (!key) continue;
          try {
            if (!item.appId) {
              counts[key] = null;
              errors[key] = 'appId is not set';
              continue;
            }
            const appId = String(item.appId);
            let finalQuery = String(item.query || '').trim();
            if (item.viewIdOrName) {
              const views = await getViewsCached(appId);
              const base = filterCondFromViews(views, item.viewIdOrName);
              finalQuery = combineQuery(base, finalQuery);
            }
            const count = await countRecords(appId, finalQuery);
            counts[key] = count;
          } catch (e) {
            counts[key] = null;
            errors[key] = String(e?.message || e);
          }
        }
        const payloadOut = { __kfav__: true, replyTo: id, ok: true, counts };
        if (Object.keys(errors).length) payloadOut.errors = errors;
        window.postMessage(payloadOut, ORIGIN);
        return;
      }

      if (type === 'LIST_VIEWS') {
        const { appId } = payload;
        const views = await getViews(appId);
        window.postMessage({ __kfav__: true, replyTo: id, ok: true, views }, ORIGIN);
        return;
      }
      // 逶ｴ霑代せ繧ｱ繧ｸ繝･繝ｼ繝ｫ蜿門ｾ暦ｼ壻ｻｻ諢上い繝励Μ ﾃ・莉ｻ諢上・譌･莉倥ヵ繧｣繝ｼ繝ｫ繝・ﾃ・蜈・騾ｱ髢・
      if (type === 'LIST_SCHEDULE') {
        const { appId, dateField, endDate, extraQuery = '', titleField = '', limit = 20 } = payload;
        const today = new Date();
        const yyyy = today.getFullYear();
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        const dd = String(today.getDate()).padStart(2, '0');
        const todayDate = `${yyyy}-${mm}-${dd}`;
        const off = tzOffsetStr(today);
        const startDT = `${todayDate}T00:00:00${off}`;
        const endDT = `${endDate}T23:59:59${off}`;

        // 繝輔ぅ繝ｼ繝ｫ繝牙梛繧貞叙蠕励＠縺ｦ DATE / DATETIME 繧貞愛螳夲ｼ亥､ｱ謨玲凾縺ｯDATE縺ｨ縺励※蜃ｦ逅・ｼ・
        let fieldType = 'DATE';
        try {
          const fieldsMeta = await getFields(appId);
          const meta = fieldsMeta.find(f => f.code === dateField);
          if (meta && meta.type) fieldType = meta.type; // "DATE" | "DATETIME" 縺ｪ縺ｩ
        } catch (_e) { /* 讓ｩ髯舌′辟｡縺・ｭ峨・鮟吶▲縺ｦ繝輔か繝ｼ繝ｫ繝舌ャ繧ｯ */ }

        let q;
        if (fieldType === 'DATETIME') {
          q = `${dateField} >= "${startDT}" and ${dateField} <= "${endDT}"`;
        } else {
          q = `${dateField} >= "${todayDate}" and ${dateField} <= "${endDate}"`;
        }
        // 遨ｺ逋ｽ髯､螟厄ｼ亥ｿｵ縺ｮ縺溘ａ・・
        q += ` and ${dateField} != ""`;
        if (extraQuery && String(extraQuery).trim()) q += ` and (${extraQuery})`;
        q += ` order by ${dateField} asc limit ${Math.max(1, Math.min(100, limit))}`;

        const fields = Array.from(new Set(['$id', dateField].concat(titleField ? [titleField] : [])));
        const resp = await window.kintone.api(window.kintone.api.url('/k/v1/records', true), 'GET',
          { app: String(appId), query: q, fields });
        //縺昴・縺ｾ縺ｾ霑斐☆・域紛蠖｢縺ｯ諡｡蠑ｵ蛛ｴ縺ｧ・・
        window.postMessage({ __kfav__: true, replyTo: id, ok: true, records: resp.records || [] }, ORIGIN);
        return;
      }
      // 繝輔ぅ繝ｼ繝ｫ繝我ｸ隕ｧ・亥梛莉倥″・峨ｒ霑斐☆・啅I縺ｮ驕ｸ謚櫁｣懷勧逕ｨ
      if (type === 'LIST_FIELDS') {
        const { appId } = payload;
        const fields = await getFields(appId);
        window.postMessage({ __kfav__: true, replyTo: id, ok: true, fields }, ORIGIN);
        return;
      }

      if (type === 'PIN_FETCH') {
        const items = Array.isArray(payload?.items) ? payload.items : [];
        const results = {};
        for (const item of items) {
          const key = item?.id;
          if (!key) continue;
          try {
            if (!item.appId || !item.recordId) {
              results[key] = { ok: false, error: 'appId / recordId is required' };
              continue;
            }
            const params = {
              app: String(item.appId),
              id: String(item.recordId)
            };
            if (Array.isArray(item.fields) && item.fields.length) {
              params.fields = item.fields;
            }
            const resp = await window.kintone.api(window.kintone.api.url('/k/v1/record', true), 'GET', params);
            results[key] = { ok: true, record: resp.record || {} };
          } catch (error) {
            results[key] = { ok: false, error: String(error?.message || error) };
          }
        }
        window.postMessage({ __kfav__: true, replyTo: id, ok: true, results }, ORIGIN);
        return;
      }
    } catch (e) {
      const payloadOut = {
        __kfav__: true,
        replyTo: id,
        ok: false,
        error: String(e?.message || e)
      };
      if (e && typeof e === 'object') {
        if ('code' in e) payloadOut.errorCode = e.code;
        if ('status' in e) payloadOut.status = e.status;
        if ('errors' in e) payloadOut.errorDetails = e.errors;
      }
      window.postMessage(payloadOut, ORIGIN);
    }
  });
})();





