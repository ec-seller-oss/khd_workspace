/**
 * Land Scanner — 物件検索ページ (search.js)
 *
 * 機能:
 *  - 検索フォームで POST /api/search に送信
 *  - 結果をテーブルに動的描画
 *  - 利回り色分け (≥8%=緑, ≥7%=黄, <7%=赤)
 *  - カラムヘッダークリックでソート
 *  - ページネーション
 *  - ローディングスピナー
 *  - ボリュームチェックボタン → /volume-check?propertyId=xxx
 */

(function () {
  'use strict';

  // ---- 状態 -------------------------------------------------------------------
  let _allResults   = [];
  let _currentPage  = 1;
  let _sortCol      = 'scraped_at';
  let _sortDir      = 'desc';
  const PER_PAGE    = 50;

  // ---- DOMContentLoaded -------------------------------------------------------
  document.addEventListener('DOMContentLoaded', function () {
    const form = document.getElementById('search-form');
    if (form) {
      form.addEventListener('submit', function (e) {
        e.preventDefault();
        _currentPage = 1;
        _doSearch();
      });
    }
  });

  // ---- 検索実行 ---------------------------------------------------------------
  async function _doSearch() {
    const payload = _buildPayload();

    _showLoading(true);
    _hideAll();

    try {
      const resp = await fetch('/api/search', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });

      if (resp.status === 401) {
        location.href = '/login';
        return;
      }
      if (!resp.ok) {
        const err = await resp.text();
        throw new Error(err);
      }

      const data = await resp.json();
      _allResults = data.results || [];

      _updateResultsHeader(_allResults.length);
      _renderTable();

    } catch (e) {
      console.error('[search] error:', e);
      _showInitialMsg('検索エラー: ' + e.message);
    } finally {
      _showLoading(false);
    }
  }

  // ---- ペイロード構築 ---------------------------------------------------------
  function _buildPayload() {
    const lines = Array.from(
      document.querySelectorAll('.search-line:checked')
    ).map(cb => cb.value);

    const prefs = Array.from(
      document.querySelectorAll('.search-pref:checked')
    ).map(cb => cb.value);

    return {
      lines:                lines,
      prefectures:          prefs,
      min_yield:            _numOrNull('min-yield'),
      max_land_price:       _numOrNull('max-land'),
      max_total_investment: _numOrNull('max-total'),
      min_area:             _numOrNull('min-area'),
      max_area:             _numOrNull('max-area'),
      room_area:            _numOrDefault('calc-room-area', 20),
      build_cost_per_tsubo: _numOrDefault('calc-build-cost', 80),
      monthly_rent:         _numOrDefault('calc-rent', 5),
      demolition_unit:      _numOrDefault('calc-demolition', 3),
      living_ratio:         70,
      days:                 7,
      page:                 1,
      per_page:             1000,  // クライアント側でページネーション
    };
  }

  // ---- テーブル描画 -----------------------------------------------------------
  function _renderTable() {
    const sorted = _sortResults([..._allResults]);
    const totalPages = Math.max(1, Math.ceil(sorted.length / PER_PAGE));
    _currentPage = Math.min(_currentPage, totalPages);

    const page = sorted.slice((_currentPage - 1) * PER_PAGE, _currentPage * PER_PAGE);

    const tbody = document.getElementById('results-body');
    if (!tbody) return;

    if (sorted.length === 0) {
      _showNoResults();
      return;
    }

    document.getElementById('results-table-wrap')?.classList.remove('hidden');
    document.getElementById('no-results')?.classList.add('hidden');
    document.getElementById('initial-msg')?.classList.add('hidden');

    tbody.innerHTML = page.map(_buildRow).join('');

    // ページネーション更新
    document.getElementById('page-info').textContent = `${_currentPage} / ${totalPages} ページ`;
    document.getElementById('prev-page').disabled = _currentPage <= 1;
    document.getElementById('next-page').disabled = _currentPage >= totalPages;
  }

  // ---- 1行のHTML構築 ---------------------------------------------------------
  function _buildRow(item) {
    const yc = item.yield_pct >= 8 ? 'yield-high' : item.yield_pct >= 7 ? 'yield-mid' : 'yield-low';
    const sourceColors = { suumo: 'blue', athome: 'green', homes: 'orange', reins: 'purple' };
    const sc = sourceColors[item.source] || 'gray';

    const titleHtml = item.url
      ? `<a href="${_esc(item.url)}" target="_blank" rel="noopener" class="hover:text-blue-600 hover:underline">${_esc(item.title || '（タイトルなし）')}</a>`
      : _esc(item.title || '（タイトルなし）');

    const stationHtml = item.nearest_station
      ? `${_esc(item.nearest_station)}${item.walk_minutes ? `<span class="text-gray-400 ml-1">徒歩${item.walk_minutes}分</span>` : ''}`
      : '--';

    return `
    <tr class="hover:bg-blue-50 transition-colors">
      <td class="px-4 py-3 text-gray-500 text-xs whitespace-nowrap">${item.scraped_at || '--'}</td>
      <td class="px-4 py-3">
        <span class="inline-block bg-${sc}-100 text-${sc}-700 text-[10px] font-semibold px-2 py-0.5 rounded-full uppercase">${_esc(item.source)}</span>
      </td>
      <td class="px-4 py-3 max-w-xs">
        <div class="font-medium text-gray-800 text-xs truncate">${titleHtml}</div>
        <div class="text-gray-400 text-xs truncate">${_esc(item.address || '')}</div>
      </td>
      <td class="px-4 py-3 text-right font-semibold text-gray-800 whitespace-nowrap">
        ${item.price ? item.price.toLocaleString() + '万円' : '--'}
      </td>
      <td class="px-4 py-3 text-right text-gray-600 whitespace-nowrap">
        ${item.area ? item.area.toFixed(1) + '㎡' : '--'}
      </td>
      <td class="px-4 py-3 text-xs text-gray-600 whitespace-nowrap">${stationHtml}</td>
      <td class="px-4 py-3 text-right whitespace-nowrap">
        <span class="${yc}">${item.yield_pct != null ? item.yield_pct.toFixed(2) + '%' : '--'}</span>
      </td>
      <td class="px-4 py-3 text-center whitespace-nowrap">
        <a href="/volume-check?propertyId=${item.id}&source=${_esc(item.source)}"
          class="inline-block bg-blue-600 hover:bg-blue-700 text-white text-[11px] font-semibold px-2.5 py-1 rounded transition-colors">
          ボリュームチェック
        </a>
      </td>
    </tr>`;
  }

  // ---- ソート（グローバル: HTMLのonclick用）----------------------------------
  window.sortResults = function (col) {
    if (_sortCol === col) {
      _sortDir = _sortDir === 'asc' ? 'desc' : 'asc';
    } else {
      _sortCol = col;
      _sortDir = 'desc';
    }
    document.querySelectorAll('.sort-icon').forEach(el => {
      el.textContent = el.dataset.col === col ? (_sortDir === 'asc' ? '↑' : '↓') : '↕';
    });
    _renderTable();
  };

  function _sortResults(arr) {
    return arr.sort((a, b) => {
      let va = a[_sortCol] ?? '';
      let vb = b[_sortCol] ?? '';
      if (typeof va === 'string') va = va.toLowerCase();
      if (typeof vb === 'string') vb = vb.toLowerCase();
      if (va < vb) return _sortDir === 'asc' ? -1 : 1;
      if (va > vb) return _sortDir === 'asc' ? 1 : -1;
      return 0;
    });
  }

  // ---- ページ変更（グローバル）-----------------------------------------------
  window.changePage = function (delta) {
    const totalPages = Math.max(1, Math.ceil(_allResults.length / PER_PAGE));
    _currentPage = Math.max(1, Math.min(totalPages, _currentPage + delta));
    _renderTable();
  };

  // ---- UI ヘルパー ------------------------------------------------------------
  function _showLoading(show) {
    document.getElementById('loading')?.classList.toggle('hidden', !show);
  }

  function _hideAll() {
    document.getElementById('initial-msg')?.classList.add('hidden');
    document.getElementById('results-table-wrap')?.classList.add('hidden');
    document.getElementById('no-results')?.classList.add('hidden');
    document.getElementById('results-header')?.classList.add('hidden');
  }

  function _updateResultsHeader(count) {
    const header = document.getElementById('results-header');
    header?.classList.remove('hidden');
    const countEl = document.getElementById('results-count');
    if (countEl) countEl.textContent = count;
    const updEl = document.getElementById('results-updated');
    if (updEl) updEl.textContent = new Date().toLocaleTimeString('ja-JP');
  }

  function _showNoResults() {
    document.getElementById('results-table-wrap')?.classList.add('hidden');
    document.getElementById('initial-msg')?.classList.add('hidden');
    document.getElementById('no-results')?.classList.remove('hidden');
  }

  function _showInitialMsg(msg) {
    const el = document.getElementById('initial-msg');
    if (el) {
      el.querySelector('p')?.remove();
      const p = document.createElement('p');
      p.textContent = msg;
      el.appendChild(p);
      el.classList.remove('hidden');
    }
  }

  // ---- ユーティリティ ---------------------------------------------------------
  function _numOrNull(id) {
    const v = parseFloat(document.getElementById(id)?.value);
    return isNaN(v) ? null : v;
  }
  function _numOrDefault(id, def) {
    const v = parseFloat(document.getElementById(id)?.value);
    return isNaN(v) ? def : v;
  }
  function _esc(str) {
    if (str == null) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
})();
