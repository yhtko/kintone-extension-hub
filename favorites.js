'use strict';

import {
  loadFavorites,
  saveFavorites,
  sortFavorites,
  getActiveTab,
  parseKintoneUrl,
  isKintoneUrl,
  hasHostPermission,
  sendCountBulk,
  createId
} from './core.js';

export function initFavorites(options = {}) {
  const doc = options.document || document;
  const state = {
    favorites: [],
    refreshInterval: options.refreshInterval ?? 30000,
    countTimer: null,
    storageListener: null,
    filterText: '',
    badgeStatus: new Map()
  };

  const favListEl = doc.getElementById('favList');
  if (!favListEl) {
    return {
      dispose() {},
      async refreshCounts() {}
    };
  }

  const openOptionsBtn = doc.getElementById('openOptions');
  const refreshCountsBtn = doc.getElementById('refreshCounts');
  const noticeEl = doc.getElementById('notice');
  const filterEl = doc.getElementById('filter');

  function setNotice(msg) {
    if (!noticeEl) return;
    if (msg) {
      noticeEl.textContent = msg;
      noticeEl.classList.remove('hidden');
    } else {
      noticeEl.textContent = '';
      noticeEl.classList.add('hidden');
    }
  }

  function renderEmptyMessage() {
    favListEl.innerHTML = '<li style="padding:10px 0;color:#888;font-size:12px;">ウォッチリストはまだありません。ウォッチしたいビューを右クリックするか、上部の「＋」ボタンから登録できます。</li>';
  }

  function renderFavorites(items) {
    favListEl.innerHTML = '';
    if (!items.length) {
      renderEmptyMessage();
      return;
    }

    for (const f of items) {
      const li = doc.createElement('li');
      li.className = 'fav';
      li.dataset.id = f.id;

      const left = doc.createElement('div');
      const title = doc.createElement('div');
      title.className = 'fav-title';
      title.textContent = f.label || '(no label)';
      left.appendChild(title);

      const right = doc.createElement('div');
      right.className = 'actions';
      const openLink = doc.createElement('a');
      openLink.className = 'open';
      openLink.href = f.url;
      openLink.target = '_blank';
      openLink.rel = 'noopener';
      openLink.textContent = '開く';
      const badge = doc.createElement('div');
      badge.className = 'badge';
      const badgeState = state.badgeStatus.get(f.id);
      if (badgeState) {
        badge.textContent = badgeState.text;
        badge.title = badgeState.title;
        if (badgeState.loading) badge.classList.add('loading');
      } else {
        const initialState = { text: '-', title: '件数', loading: false };
        state.badgeStatus.set(f.id, initialState);
        badge.textContent = initialState.text;
        badge.title = initialState.title;
      }

      right.appendChild(openLink);
      right.appendChild(badge);

      li.appendChild(left);
      li.appendChild(right);
      favListEl.appendChild(li);
    }
  }

  function filteredFavorites() {
    const keyword = state.filterText.trim().toLowerCase();
    if (!keyword) return state.favorites;
    return state.favorites.filter((item) => {
      const targets = [item.label, item.url, item.host, item.viewIdOrName, item.query];
      return targets.some((value) => typeof value === 'string' && value.toLowerCase().includes(keyword));
    });
  }

  function renderCurrentFavorites() {
    const filtered = filteredFavorites();
    if (!filtered.length) {
      favListEl.innerHTML = '';
      if (!state.favorites.length) {
        renderEmptyMessage();
      } else {
        favListEl.innerHTML = '<li class="muted" style="padding:10px 0;">検索条件に一致するウォッチ対象はありません。</li>';
      }
      return;
    }
    renderFavorites(filtered);
  }

  function badgeElementFor(id) {
    const escapeId = (value) => {
      if (value == null) return '';
      if (globalThis.CSS?.escape) return CSS.escape(String(value));
      return String(value).replace(/[^\w-]/g, (ch) => `\\${ch}`);
    };
    return favListEl.querySelector(`.fav[data-id="${escapeId(id)}"] .badge`);
  }

  function setBadgeLoading(id, on = true) {
    const prev = state.badgeStatus.get(id) || { text: '-', title: '件数', loading: false };
    const next = {
      text: on ? '...' : prev.text,
      title: on ? '取得中' : prev.title,
      loading: on
    };
    state.badgeStatus.set(id, next);

    const el = badgeElementFor(id);
    if (!el) return;
    el.classList.toggle('loading', on);
    el.textContent = next.text;
    el.title = next.title;
  }

  function setBadgeValue(id, val, title = '件数') {
    const textValue = val === null || val === undefined ? '-' : String(val);
    state.badgeStatus.set(id, { text: textValue, title, loading: false });

    const el = badgeElementFor(id);
    if (!el) return;
    el.classList.remove('loading');
    el.textContent = textValue;
    el.title = title;
  }

  function groupByHost(items) {
    const map = new Map();
    for (const it of items) {
      if (!it.host) continue;
      if (!map.has(it.host)) map.set(it.host, []);
      map.get(it.host).push(it);
    }
    return map;
  }

  async function refreshCounts() {
    if (!state.favorites.length) return;
    setNotice('');

    for (const it of state.favorites) setBadgeLoading(it.id, true);

    const grouped = groupByHost(state.favorites);
    const missingPerm = new Set();
    const failedHosts = new Set();
    for (const [host, items] of grouped.entries()) {
      try {
        const permitted = await hasHostPermission(host);
        if (!permitted) {
          for (const it of items) setBadgeValue(it.id, null, 'このホストの許可が必要です（設定）');
          missingPerm.add(host);
          continue;
        }
        try {
          await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host });
        } catch (_warmErr) {}
        const payloadItems = items.map((it) => ({
          id: it.id,
          appId: it.appId,
          viewIdOrName: it.viewIdOrName || '',
          query: it.query || ''
        }));
        const res = await sendCountBulk(host, payloadItems);
        if (res?.ok) {
          const counts = res.counts || {};
          const errors = res.errors || {};
          for (const it of items) {
            if (errors[it.id]) {
              setBadgeValue(it.id, null, errors[it.id]);
              failedHosts.add(host);
            } else if (Object.prototype.hasOwnProperty.call(counts, it.id)) {
              setBadgeValue(it.id, counts[it.id]);
            } else {
              setBadgeValue(it.id, null, '取得できませんでした');
              failedHosts.add(host);
            }
          }
        } else {
          const message = res?.error || '取得失敗';
          for (const it of items) setBadgeValue(it.id, null, message);
          failedHosts.add(host);
        }
      } catch (e) {
        const message = String(e?.message || e);
        for (const it of items) setBadgeValue(it.id, null, message);
        failedHosts.add(host);
      }
    }

    if (missingPerm.size) {
      setNotice(`⚠ 以下のドメインの許可が必要です：${Array.from(missingPerm).join(', ')}`);
    } else if (failedHosts.size) {
      setNotice('一部の件数取得に失敗しました。サインイン状態と権限を確認してください。');
    }
  }

  async function loadAndRenderFavorites() {
    const list = await loadFavorites();
    state.favorites = sortFavorites(list);
    const validIds = new Set(state.favorites.map((item) => item.id));
    for (const key of Array.from(state.badgeStatus.keys())) {
      if (!validIds.has(key)) state.badgeStatus.delete(key);
    }
    renderCurrentFavorites();
    if (!state.favorites.length) {
      setNotice('');
    }
    const hosts = new Set(state.favorites.map((item) => item.host).filter(Boolean));
    for (const host of hosts) {
      try {
        await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host });
      } catch (_e) {}
    }
  }

  function attachStorageListener() {
    if (!chrome?.storage?.onChanged) return;
    const listener = (changes, area) => {
      if (area !== 'sync') return;
      if (changes.kintoneFavorites || changes.kfavBadgeTargetId) {
        loadAndRenderFavorites().catch(() => {});
      }
    };
    chrome.storage.onChanged.addListener(listener);
    state.storageListener = listener;
  }

  async function addFavoriteFromActiveTab() {
    const tab = await getActiveTab();
    if (!tab || !isKintoneUrl(tab.url || '')) {
      setNotice('kintone のタブを前面にしてください。');
      return false;
    }
    const parsed = parseKintoneUrl(tab.url);
    if (!parsed.host) {
      setNotice('kintone のURL解析に失敗しました');
      return false;
    }
    const existing = await loadFavorites();
    if (existing.some((x) => x.url === parsed.url)) {
      setNotice('このビューは既に登録済みです');
      return false;
    }
    const labelFromTab = (tab.title || '').replace(/ - kintone.*/, '') || '(no label)';
    const item = {
      id: createId(),
      label: labelFromTab,
      url: parsed.url,
      host: parsed.host,
      appId: parsed.appId,
      viewIdOrName: parsed.viewIdOrName,
      query: '',
      order: existing.length,
      pinned: false
    };
    const next = [...existing, item];
    await saveFavorites(next);
    state.favorites = sortFavorites(next);
    state.badgeStatus.set(item.id, { text: '-', title: '件数', loading: false });
    renderCurrentFavorites();
    setNotice('ウォッチリストに追加しました！ ⟳ で件数を取得できます');
    try {
      await chrome.runtime.sendMessage({ type: 'WARMUP_CONNECTOR', host: parsed.host });
    } catch (_e) {}
    return true;
  }

  function wireSearch() {
    if (!filterEl) return;
    const apply = () => {
      state.filterText = filterEl.value || '';
      renderCurrentFavorites();
    };
    filterEl.addEventListener('input', apply);
    filterEl.addEventListener('change', apply);
    apply();
  }

  function wireEvents() {
    if (openOptionsBtn) {
      openOptionsBtn.addEventListener('click', () => chrome.runtime.openOptionsPage());
    }
    if (refreshCountsBtn) {
      refreshCountsBtn.addEventListener('click', () => refreshCounts());
    }
    wireSearch();
  }

  async function boot() {
    await loadAndRenderFavorites();
    await refreshCounts().catch(() => {});
    if (state.refreshInterval > 0) {
      state.countTimer = setInterval(() => {
        refreshCounts().catch(() => {});
      }, state.refreshInterval);
    }
  }

  wireEvents();
  attachStorageListener();
  boot().catch((e) => setNotice(String(e?.message || e)));

  return {
    refreshCounts,
    quickAdd: addFavoriteFromActiveTab,
    dispose() {
      if (state.countTimer) {
        clearInterval(state.countTimer);
        state.countTimer = null;
      }
      if (state.storageListener) {
        chrome.storage.onChanged.removeListener(state.storageListener);
        state.storageListener = null;
      }
    }
  };
}










