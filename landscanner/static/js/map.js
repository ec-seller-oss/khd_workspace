/**
 * Land Scanner — Leaflet マップモジュール
 *
 * 使い方:
 *   window.MAP_MODE = 'home'      ← 読み取り専用（ポップアップのみ）
 *   window.MAP_MODE = 'settings'  ← 駅クリックで選択/解除 + コールバック
 *   window.SELECTED_STATIONS = [{ id, name, line }, ...]  ← 初期選択
 *
 * settings モード用コールバック:
 *   window.onStationToggle = function(station, isSelected) { ... }
 */

(function () {
  'use strict';

  // ---- 定数 -------------------------------------------------------------------
  const DEFAULT_CENTER = [35.72, 139.75];
  const DEFAULT_ZOOM   = 9;

  // service_count に応じた色
  function serviceColor(count) {
    if (count >= 7) return '#2563eb'; // 青
    if (count >= 6) return '#0d9488'; // ティール
    if (count >= 5) return '#16a34a'; // 緑
    if (count >= 4) return '#65a30d'; // 黄緑
    if (count >= 3) return '#ca8a04'; // 黄
    if (count >= 2) return '#ea580c'; // オレンジ
    return '#dc2626';                  // 赤
  }

  // 選択済みマーカー色（強調）
  const SELECTED_COLOR   = '#7c3aed'; // 紫
  const UNSELECTED_ALPHA = 0.75;

  // ---- 状態 -------------------------------------------------------------------
  let _map        = null;
  let _markers    = {};  // { [station.id]: L.CircleMarker }
  let _stations   = [];
  let _selected   = {};  // { [station.id]: stationObj }
  let _mode       = 'home';

  // ---- 初期化 -----------------------------------------------------------------

  /**
   * マップを初期化する
   * @param {string} containerId  HTML要素のID
   * @param {Array}  stations     駅データ配列
   * @param {Object} options      { mode: 'home'|'settings', selectedStations: [...] }
   */
  function initMap(containerId, stations, options) {
    options = options || {};
    _mode     = options.mode || window.MAP_MODE || 'home';
    _stations = stations || [];

    // 初期選択駅
    const initSel = options.selectedStations || window.SELECTED_STATIONS || [];
    _selected = {};
    initSel.forEach(s => { _selected[s.id] = s; });

    // Leaflet マップ生成
    _map = L.map(containerId, { zoomControl: true }).setView(DEFAULT_CENTER, DEFAULT_ZOOM);

    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '© <a href="https://www.openstreetmap.org/">OpenStreetMap</a> contributors',
      maxZoom: 18,
    }).addTo(_map);

    // 駅マーカー追加
    _stations.forEach(station => _addMarker(station));
  }

  // ---- マーカー追加 -----------------------------------------------------------
  function _addMarker(station) {
    if (!station.lat || !station.lng) return;

    const isSelected = !!_selected[station.id];
    const marker = _createMarker(station, isSelected);
    marker.addTo(_map);
    _markers[station.id] = marker;

    // ポップアップ
    marker.bindPopup(_buildPopup(station, isSelected), { maxWidth: 220 });

    // settings モードのみクリックで選択
    if (_mode === 'settings') {
      marker.on('click', function (e) {
        L.DomEvent.stopPropagation(e);
        _toggleStation(station);
      });
    }
  }

  function _createMarker(station, isSelected) {
    const color  = isSelected ? SELECTED_COLOR : serviceColor(station.service_count);
    const radius = isSelected ? 10 : 8;
    return L.circleMarker([station.lat, station.lng], {
      radius:      radius,
      fillColor:   color,
      color:       '#fff',
      weight:      isSelected ? 3 : 2,
      opacity:     1,
      fillOpacity: isSelected ? 0.95 : UNSELECTED_ALPHA,
    });
  }

  function _buildPopup(station, isSelected) {
    const selLabel = _mode === 'settings'
      ? `<div style="margin-top:6px;font-size:11px;color:${isSelected ? '#7c3aed' : '#6b7280'}">
           ${isSelected ? '✓ 選択中（クリックで解除）' : 'クリックして選択'}
         </div>`
      : '';
    return `<div style="min-width:160px;font-family:sans-serif">
      <div style="font-weight:700;font-size:14px;margin-bottom:4px">${station.name}駅</div>
      <div style="font-size:12px;color:#555;margin-bottom:2px">${station.line}</div>
      <div style="font-size:12px;color:#555;margin-bottom:2px">${station.prefecture || ''}</div>
      <div style="font-size:12px">月額家賃相場: <strong>${station.monthly_rent}万円</strong></div>
      <div style="font-size:11px;color:#888">対応サービス: ${station.service_count}件</div>
      ${selLabel}
    </div>`;
  }

  // ---- 選択トグル（settings モード）------------------------------------------
  function _toggleStation(station) {
    const wasSelected = !!_selected[station.id];

    if (wasSelected) {
      delete _selected[station.id];
    } else {
      _selected[station.id] = station;
    }

    // マーカー更新
    const isNowSelected = !wasSelected;
    const oldMarker = _markers[station.id];
    if (oldMarker) {
      _map.removeLayer(oldMarker);
    }
    const newMarker = _createMarker(station, isNowSelected);
    newMarker.addTo(_map);
    _markers[station.id] = newMarker;

    // ポップアップを新しいマーカーに付け直す
    newMarker.bindPopup(_buildPopup(station, isNowSelected), { maxWidth: 220 });
    newMarker.on('click', function (e) {
      L.DomEvent.stopPropagation(e);
      _toggleStation(station);
    });

    // コールバック
    if (typeof window.onStationToggle === 'function') {
      window.onStationToggle(station, isNowSelected);
    }
  }

  // ---- 外部から選択状態を更新 -------------------------------------------------
  function setSelectedStations(stationList) {
    _selected = {};
    stationList.forEach(s => { _selected[s.id] = s; });

    // 全マーカーを再描画
    Object.keys(_markers).forEach(id => {
      _map.removeLayer(_markers[id]);
      delete _markers[id];
    });
    _stations.forEach(station => _addMarker(station));
  }

  // ---- 現在の選択リストを取得 -------------------------------------------------
  function getSelectedStations() {
    return Object.values(_selected);
  }

  // ---- 公開 API ---------------------------------------------------------------
  window.LandScannerMap = {
    init:                initMap,
    setSelectedStations: setSelectedStations,
    getSelectedStations: getSelectedStations,
    toggleStation:       _toggleStation,
  };
})();
