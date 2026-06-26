/**
 * Land Scanner — ユーザー設定ページ (settings.js)
 *
 * 呼び出し元 (settings.html) から:
 *   initSettingsPage(patternMap, stationsData)
 *
 * patternMap  : { [patternId]: { id, name, enabled, stations, lines, ... } }
 * stationsData: stations.json 相当の配列
 */

(function () {
  'use strict';

  // ---- 定数 -------------------------------------------------------------------
  const DEBOUNCE_MS = 600;   // 自動保存のデバウンス
  const API_BASE    = '/api';

  // ---- 状態 -------------------------------------------------------------------
  let _patternMap    = {};
  let _currentId     = null;
  let _debounceTimer = null;
  let _selectedStations = []; // [{ id, name, line }, ...]

  // ---- エントリポイント -------------------------------------------------------
  window.initSettingsPage = function (patternMap, stationsData) {
    _patternMap = patternMap || {};

    // 最初のパターンを選択
    const ids = Object.keys(_patternMap);
    if (ids.length > 0) {
      _currentId = parseInt(ids[0], 10);
    }

    // 設定モードマップを初期化
    _initMap(stationsData);

    // 最初のパターンをフォームに展開
    if (_currentId) {
      _populateForm(_patternMap[_currentId]);
    }

    // フォームの変更を監視してデバウンス自動保存
    _attachFormListeners();
  };

  // ---- パターン選択（グローバル関数: HTMLのonclick用）-----------------------
  window.selectPattern = function (patternId) {
    _currentId = patternId;

    // タブのアクティブ切り替え
    document.querySelectorAll('.pattern-tab').forEach(btn => {
      const isActive = parseInt(btn.dataset.patternId, 10) === patternId;
      btn.classList.toggle('active', isActive);
      if (!isActive) {
        btn.classList.remove('border-blue-500', 'text-blue-600', 'bg-blue-50');
        btn.classList.add('border-gray-200', 'text-gray-600');
      }
    });

    // フォームを表示
    document.getElementById('pattern-form')?.classList.remove('hidden');

    // フォームにデータをセット
    const pattern = _patternMap[patternId];
    if (pattern) {
      document.getElementById('current-pattern-id').value = patternId;
      _populateForm(pattern);
    }
  };

  // ---- パターン作成（グローバル）--------------------------------------------
  window.createPattern = async function () {
    const count = Object.keys(_patternMap).length;
    if (count >= 3) {
      _showToast('最大3パターンまでです', 'error');
      return;
    }

    try {
      const resp = await fetch(`${API_BASE}/settings/pattern`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: `パターン${count + 1}` }),
      });
      if (!resp.ok) throw new Error(await resp.text());
      // ページリロードで新パターンを反映
      location.reload();
    } catch (e) {
      _showToast('作成に失敗しました: ' + e.message, 'error');
    }
  };

  // ---- パターン保存（グローバル）--------------------------------------------
  window.savePattern = async function () {
    if (!_currentId) return;
    await _saveCurrentPattern();
  };

  // ---- パターン削除（グローバル）--------------------------------------------
  window.deletePattern = async function () {
    if (!_currentId) return;
    const patternName = _patternMap[_currentId]?.name || 'このパターン';
    if (!confirm(`「${patternName}」を削除しますか？\nこの操作は元に戻せません。`)) return;

    try {
      const resp = await fetch(`${API_BASE}/settings/pattern/${_currentId}`, {
        method: 'DELETE',
      });
      if (!resp.ok) throw new Error(await resp.text());
      location.reload();
    } catch (e) {
      _showToast('削除に失敗しました: ' + e.message, 'error');
    }
  };

  // ---- 内部: フォームにデータをセット ----------------------------------------
  function _populateForm(pattern) {
    if (!pattern) return;

    _setVal('pattern-name',     pattern.name);
    _setChecked('pattern-enabled', pattern.enabled);
    _setVal('build-cost',       pattern.build_cost_per_tsubo);
    _setVal('room-area',        pattern.room_area);
    _setVal('monthly-rent',     pattern.monthly_rent);
    _setVal('living-ratio',     pattern.living_ratio);
    _setVal('demolition-unit',  pattern.demolition_unit);
    _setVal('yield-threshold',  pattern.yield_threshold);
    _setVal('max-land-price',   pattern.max_land_price ?? '');
    _setVal('max-total',        pattern.max_total_investment ?? '');
    _setVal('min-area',         pattern.min_area ?? '');
    _setVal('max-area',         pattern.max_area ?? '');
    _setChecked('notify-email',         pattern.notify_email);
    _setChecked('notify-announcement',  pattern.notify_announcement);
    _setVal('notify-timing',    pattern.notify_timing);

    // 沿線チェックボックス
    document.querySelectorAll('.line-checkbox').forEach(cb => {
      cb.checked = (pattern.lines || []).includes(cb.value);
    });

    // 選択済み駅
    _selectedStations = Array.isArray(pattern.stations) ? [...pattern.stations] : [];
    _renderSelectedStationTags();

    // マップに選択状態を反映
    if (window.LandScannerMap) {
      window.LandScannerMap.setSelectedStations(_selectedStations);
    }
  }

  // ---- 内部: フォームからデータを収集 ----------------------------------------
  function _collectFormData() {
    const lines = [];
    document.querySelectorAll('.line-checkbox:checked').forEach(cb => lines.push(cb.value));

    return {
      name:                 _getVal('pattern-name')    || 'パターン',
      enabled:              _getChecked('pattern-enabled'),
      stations:             _selectedStations,
      lines:                lines,
      build_cost_per_tsubo: _getNum('build-cost')       || 80,
      room_area:            _getNum('room-area')         || 20,
      monthly_rent:         _getNum('monthly-rent')      || 5,
      living_ratio:         _getNum('living-ratio')      || 70,
      demolition_unit:      _getNum('demolition-unit')   || 3,
      yield_threshold:      _getNum('yield-threshold')   || 7,
      max_land_price:       _getNumOrNull('max-land-price'),
      max_total_investment: _getNumOrNull('max-total'),
      min_area:             _getNumOrNull('min-area'),
      max_area:             _getNumOrNull('max-area'),
      notify_email:         _getChecked('notify-email'),
      notify_announcement:  _getChecked('notify-announcement'),
      notify_timing:        _getVal('notify-timing') || 'daily',
    };
  }

  // ---- 内部: 保存（PUT）-------------------------------------------------------
  async function _saveCurrentPattern() {
    if (!_currentId) return;
    const data = _collectFormData();

    try {
      const resp = await fetch(`${API_BASE}/settings/pattern/${_currentId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!resp.ok) throw new Error(await resp.text());

      // ローカルキャッシュも更新
      if (_patternMap[_currentId]) {
        Object.assign(_patternMap[_currentId], data);
      }
      _showToast('保存しました ✓', 'success');
    } catch (e) {
      _showToast('保存に失敗しました: ' + e.message, 'error');
    }
  }

  // ---- 内部: フォームリスナー -------------------------------------------------
  function _attachFormListeners() {
    const form = document.getElementById('settings-form');
    if (!form) return;

    form.addEventListener('input', _onFormChange);
    form.addEventListener('change', _onFormChange);
  }

  function _onFormChange() {
    clearTimeout(_debounceTimer);
    _debounceTimer = setTimeout(_saveCurrentPattern, DEBOUNCE_MS);
  }

  // ---- 内部: マップ初期化 -----------------------------------------------------
  function _initMap(stationsData) {
    const container = document.getElementById('settings-map');
    if (!container || !window.LandScannerMap) return;

    window.LandScannerMap.init('settings-map', stationsData, {
      mode: 'settings',
      selectedStations: _selectedStations,
    });

    // 駅トグルコールバック
    window.onStationToggle = function (station, isSelected) {
      if (isSelected) {
        if (!_selectedStations.find(s => s.id === station.id)) {
          _selectedStations.push({ id: station.id, name: station.name, line: station.line });
        }
      } else {
        _selectedStations = _selectedStations.filter(s => s.id !== station.id);
      }
      _renderSelectedStationTags();

      // デバウンス自動保存
      clearTimeout(_debounceTimer);
      _debounceTimer = setTimeout(_saveCurrentPattern, DEBOUNCE_MS);
    };
  }

  // ---- 内部: 選択済み駅タグ表示 ----------------------------------------------
  function _renderSelectedStationTags() {
    const container = document.getElementById('selected-stations');
    const noMsg     = document.getElementById('no-stations-msg');
    if (!container) return;

    // 既存タグを削除（noMsg は保持）
    Array.from(container.children).forEach(el => {
      if (el.id !== 'no-stations-msg') el.remove();
    });

    if (_selectedStations.length === 0) {
      noMsg?.classList.remove('hidden');
    } else {
      noMsg?.classList.add('hidden');
      _selectedStations.forEach(s => {
        const tag = document.createElement('span');
        tag.className = 'station-tag';
        tag.innerHTML = `${s.name}
          <button type="button" aria-label="${s.name}を解除"
            onclick="removeStation(${s.id})">×</button>`;
        container.appendChild(tag);
      });
    }
  }

  // ---- タグのXボタン（グローバル）--------------------------------------------
  window.removeStation = function (stationId) {
    _selectedStations = _selectedStations.filter(s => s.id !== stationId);
    _renderSelectedStationTags();
    if (window.LandScannerMap) {
      window.LandScannerMap.setSelectedStations(_selectedStations);
    }
    clearTimeout(_debounceTimer);
    _debounceTimer = setTimeout(_saveCurrentPattern, DEBOUNCE_MS);
  };

  // ---- トースト表示 -----------------------------------------------------------
  function _showToast(msg, type) {
    const toast = document.getElementById('toast');
    const span  = document.getElementById('toast-msg');
    if (!toast || !span) return;

    span.textContent = msg;
    toast.classList.remove('hidden', 'hide');

    if (type === 'error') {
      toast.classList.remove('bg-gray-900');
      toast.classList.add('bg-red-700');
    } else {
      toast.classList.remove('bg-red-700');
      toast.classList.add('bg-gray-900');
    }

    toast.classList.add('show');

    clearTimeout(toast._hideTimer);
    toast._hideTimer = setTimeout(() => {
      toast.classList.remove('show');
      toast.classList.add('hide');
      setTimeout(() => toast.classList.add('hidden'), 320);
    }, 2500);
  }

  // ---- ユーティリティ ---------------------------------------------------------
  function _setVal(id, v) {
    const el = document.getElementById(id);
    if (el) el.value = (v === null || v === undefined) ? '' : v;
  }
  function _setChecked(id, v) {
    const el = document.getElementById(id);
    if (el) el.checked = !!v;
  }
  function _getVal(id) {
    return document.getElementById(id)?.value || '';
  }
  function _getChecked(id) {
    return !!document.getElementById(id)?.checked;
  }
  function _getNum(id) {
    return parseFloat(document.getElementById(id)?.value) || 0;
  }
  function _getNumOrNull(id) {
    const v = parseFloat(document.getElementById(id)?.value);
    return isNaN(v) ? null : v;
  }
})();
