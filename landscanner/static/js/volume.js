/**
 * ボリュームチェック - リアルタイム計算（サーバー往復なし）
 */

const inputIds = [
  'land-price', 'land-area', 'far', 'bcr',
  'monthly-rent', 'room-area', 'rooms',
  'build-cost', 'living-ratio',
  'demolition-unit', 'has-old-house', 'misc-cost',
];

function calcRooms(landArea, far, livingRatio, roomArea) {
  const floorArea = landArea * far / 100;
  const rooms = Math.floor(floorArea * livingRatio / 100 / roomArea);
  return Math.max(rooms, 1);
}

function calculate() {
  const landPrice = parseFloat(document.getElementById('land-price').value);
  const landArea = parseFloat(document.getElementById('land-area').value);
  const far = parseFloat(document.getElementById('far').value) || 200;
  const bcr = parseFloat(document.getElementById('bcr').value) || 60;
  const monthlyRent = parseFloat(document.getElementById('monthly-rent').value) || 5;
  const roomArea = parseFloat(document.getElementById('room-area').value) || 20;
  const roomsInput = parseInt(document.getElementById('rooms').value);
  const buildCost = parseFloat(document.getElementById('build-cost').value) || 80;
  const livingRatio = parseFloat(document.getElementById('living-ratio').value) || 70;
  const demolitionUnit = parseFloat(document.getElementById('demolition-unit').value) || 3;
  const hasOldHouse = document.getElementById('has-old-house').checked;
  const miscInput = parseFloat(document.getElementById('misc-cost').value);

  if (!landPrice || !landArea) {
    resetResults();
    return;
  }

  // 延床面積
  const floorArea = landArea * far / 100;

  // 部屋数
  const rooms = roomsInput > 0 ? roomsInput : calcRooms(landArea, far, livingRatio, roomArea);

  // 年間賃料
  const annualRent = monthlyRent * rooms * 12;

  // 建設費用
  const construction = floorArea / 3.3 * buildCost;

  // 解体費
  let demolition = 0;
  if (hasOldHouse) {
    const estimatedOldArea = landArea * bcr / 100 * 2;
    demolition = estimatedOldArea * demolitionUnit;
  }

  // 諸費用
  const misc = (!isNaN(miscInput) && miscInput > 0)
    ? miscInput
    : landPrice * 0.07 + construction * 0.04;

  // 総投資額
  const total = landPrice + construction + demolition + misc;

  // 利回り
  const yieldPct = total > 0 ? (annualRent / total * 100) : 0;

  // 坪単価
  const tsuboPrice = landArea > 0 ? landPrice / (landArea / 3.3) : 0;
  const ichiTsubo = far > 0 ? tsuboPrice / (far / 100) : 0;

  // 表示更新
  updateResults({
    rooms, floorArea, annualRent, construction,
    demolition, misc, total, yieldPct, tsuboPrice, ichiTsubo
  });
}

function updateResults(r) {
  const fmt = (v, d = 1) => v.toFixed(d).replace(/\B(?=(\d{3})+(?!\d))/g, ',');

  const yieldEl = document.getElementById('result-yield');
  const yieldLabelEl = document.getElementById('result-yield-label');

  yieldEl.textContent = fmt(r.yieldPct, 2) + '%';

  if (r.yieldPct >= 8) {
    yieldEl.className = 'text-5xl font-extrabold text-green-600';
    yieldLabelEl.textContent = '高利回り — 優良候補';
    yieldLabelEl.className = 'text-sm mt-2 text-green-600 font-semibold';
  } else if (r.yieldPct >= 7) {
    yieldEl.className = 'text-5xl font-extrabold text-yellow-500';
    yieldLabelEl.textContent = '中利回り — 要精査';
    yieldLabelEl.className = 'text-sm mt-2 text-yellow-600 font-semibold';
  } else {
    yieldEl.className = 'text-5xl font-extrabold text-red-500';
    yieldLabelEl.textContent = '低利回り — 要再検討';
    yieldLabelEl.className = 'text-sm mt-2 text-red-500 font-semibold';
  }

  document.getElementById('result-rooms').textContent = r.rooms + '部屋';
  document.getElementById('result-floor-area').textContent = fmt(r.floorArea) + ' ㎡';
  document.getElementById('result-annual-rent').textContent = fmt(r.annualRent) + ' 万円';
  document.getElementById('result-construction').textContent = fmt(r.construction) + ' 万円';
  document.getElementById('result-demolition').textContent = fmt(r.demolition) + ' 万円';
  document.getElementById('result-misc').textContent = fmt(r.misc) + ' 万円';
  document.getElementById('result-total').textContent = fmt(r.total) + ' 万円';
  document.getElementById('result-tsubo').textContent = fmt(r.tsuboPrice) + ' 万円/坪';
  document.getElementById('result-ichi').textContent = fmt(r.ichiTsubo) + ' 万円/坪';

  // 古家なしは解体費行を薄く
  const demolRow = document.getElementById('demolition-row');
  if (demolRow) demolRow.style.opacity = document.getElementById('has-old-house').checked ? '1' : '0.4';
}

function resetResults() {
  document.getElementById('result-yield').textContent = '--.--%';
  document.getElementById('result-yield').className = 'text-5xl font-extrabold text-gray-300';
  document.getElementById('result-yield-label').textContent = '';
  ['result-rooms','result-floor-area','result-annual-rent','result-construction',
   'result-demolition','result-misc','result-total','result-tsubo','result-ichi'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = '--';
  });
}

// URLパラメータから物件情報を自動入力
function loadFromProperty(data) {
  if (!data) return;
  if (data.price) document.getElementById('land-price').value = data.price;
  if (data.area) document.getElementById('land-area').value = data.area;
  if (data.far) document.getElementById('far').value = data.far;
  if (data.bcr) document.getElementById('bcr').value = data.bcr;
  if (data.has_old_house) document.getElementById('has-old-house').checked = true;

  // バナー表示
  const banner = document.getElementById('property-banner');
  if (banner) {
    banner.classList.remove('hidden');
    document.getElementById('property-banner-title').textContent = data.title || '物件データを読み込みました';
    document.getElementById('property-banner-sub').textContent =
      [data.nearest_station, data.address].filter(Boolean).join(' / ');
  }

  calculate();
}

// イベントリスナー登録
document.addEventListener('DOMContentLoaded', () => {
  inputIds.forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', calculate);
      el.addEventListener('change', calculate);
    }
  });

  // URLパラメータ経由の物件データ
  if (typeof PROPERTY_DATA !== 'undefined' && PROPERTY_DATA) {
    loadFromProperty(PROPERTY_DATA);
  } else {
    calculate();
  }
});
