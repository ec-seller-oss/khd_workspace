// ── KHD 運用専用 Web App（操縦席への書き込み口・建設GASと分離）──
// 新規Apps Scriptプロジェクトに丸ごと貼り→デプロイ→/exec URLをClaudeへ。
const COCKPIT_ID = "1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc";

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  var action = body.action, result;
  try {
    if      (action === "append_db02")     result = appendDb02(body);
    else if (action === "read_db02")       result = readDb02(body);
    else if (action === "update_db02")     result = updateDb02(body);
    else if (action === "delete_db02_rows") result = deleteDb02Rows(body);
    else if (action === "append_customer") result = appendCustomer(body);
    else if (action === "fix_qr_formulas") result = fixQrFormulas(body);
    else if (action === "ping")            result = { pong: true };
    else                                   result = { error: "unknown action: " + action };
    return ContentService.createTextOutput(JSON.stringify({ ok: true, result: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 所要4列＝常に数式で自動。日付=日付値(MM/DD)・時刻=時刻値・件数=数値書式を仕組みで保証。
var _TIME_COLS = ["予定開始", "予定終了", "実開始", "実終了"];
var _COUNT_COLS = ["提案数", "GIVE数", "相談数"];
var _FORMULA_COLS = ["予定所要(分)", "実所要(分)", "予実差分(分)", "達成率%"];

// ── appendDb02: 1行追記（日付/時刻/所要/件数を自動整形）────────────────────
function appendDb02(body) {
  const ss = SpreadsheetApp.openById((body && body.hostId) || COCKPIT_ID);
  var sh = _findDb02Sheet(ss);
  if (!sh) return { error: "02_作業DBが見つからない" };
  var info = _getHeaders(sh);
  if (!info) return { error: "ヘッダ(日付)が見つからない", sheet: sh.getName() };
  if (body.headers_only) return { ok: true, sheet: sh.getName(), header_row: info.hRow, headers: info.headers };
  var H = info.headers;

  var iDate = _findColExact(H, "日付");
  var iKS = _findColExact(H, "予定開始"), iKE = _findColExact(H, "予定終了");
  var iMS = _findColExact(H, "実開始"),   iME = _findColExact(H, "実終了");
  var iPlan = _findColExact(H, "予定所要(分)"), iAct = _findColExact(H, "実所要(分)");
  var iDiff = _findColExact(H, "予実差分(分)"), iRate = _findColExact(H, "達成率%");

  var row = body.row || {}, out = [], unmatched = [];
  for (var c = 0; c < H.length; c++) out.push("");
  for (var key in row) {
    var idx = _findCol(H, key);
    if (idx < 0) { unmatched.push(key); continue; }
    if (_FORMULA_COLS.indexOf(H[idx]) >= 0) continue; // 所要4列は数式で埋めるので静的値は無視
    var val = row[key];
    if (idx === iDate) { val = _toDate(val); }
    else if (_TIME_COLS.indexOf(H[idx]) >= 0) { var t = _toTimeFrac(val); if (t !== null) val = t; }
    out[idx] = val;
  }

  // 先頭挿入（ヘッダー直下=行2。最新が上）
  var newRow = info.hRow + 1;
  sh.insertRowBefore(newRow);
  sh.getRange(newRow, 1, 1, H.length).setValues([out]);

  // 書式：日付=MM/DD、時刻=HH:mm、件数=0（崩れ防止を仕組みで保証）
  if (iDate >= 0) sh.getRange(newRow, iDate + 1).setNumberFormat("MM/DD");
  _TIME_COLS.forEach(function (n) { var ix = _findColExact(H, n); if (ix >= 0) sh.getRange(newRow, ix + 1).setNumberFormat("HH:mm"); });
  _COUNT_COLS.forEach(function (n) { var ix = _findColExact(H, n); if (ix >= 0) sh.getRange(newRow, ix + 1).setNumberFormat("0"); });

  // 所要4列を自動数式で挿入（列は名前解決＝列移動に強い）
  _writeDurationFormulas(sh, newRow, iKS, iKE, iMS, iME, iPlan, iAct, iDiff, iRate);

  return { ok: true, appended_row: newRow, sheet: sh.getName(), headers: H, unmatched: unmatched };
}

// 予定所要/実所要/予実差分/達成率% を1行に数式で書き込む
function _writeDurationFormulas(sh, r, iKS, iKE, iMS, iME, iPlan, iAct, iDiff, iRate) {
  function L(ix) { return _colLetter(ix + 1); }
  if (iPlan >= 0 && iKS >= 0 && iKE >= 0)
    sh.getRange(r, iPlan + 1).setFormula('=IF(' + L(iKS) + r + '="","",ROUND((' + L(iKE) + r + '-' + L(iKS) + r + ')*1440,0))').setNumberFormat("0");
  if (iAct >= 0 && iMS >= 0 && iME >= 0)
    sh.getRange(r, iAct + 1).setFormula('=IF(' + L(iMS) + r + '="","",ROUND((' + L(iME) + r + '-' + L(iMS) + r + ')*1440,0))').setNumberFormat("0");
  if (iDiff >= 0 && iPlan >= 0 && iAct >= 0)
    sh.getRange(r, iDiff + 1).setFormula('=IF(OR(' + L(iPlan) + r + '="",' + L(iAct) + r + '=""),"",' + L(iPlan) + r + '-' + L(iAct) + r + ')').setNumberFormat("0");
  if (iRate >= 0 && iPlan >= 0 && iAct >= 0)
    sh.getRange(r, iRate + 1).setFormula('=IF(OR(' + L(iAct) + r + '="",' + L(iAct) + r + '=0,' + L(iPlan) + r + '="",' + L(iPlan) + r + '=0),"",ROUND(' + L(iAct) + r + '/' + L(iPlan) + r + '*100,1))').setNumberFormat("0");
}

// ── readDb02: 指定日付の全行を返す ────────────────────────────────────────
// body: { date: "2026-06-07" or "2026/06/07", limit: 50 }
function readDb02(body) {
  const ss = SpreadsheetApp.openById((body && body.hostId) || COCKPIT_ID);
  var sh = _findDb02Sheet(ss);
  if (!sh) return { error: "02_作業DBが見つからない" };
  var info = _getHeaders(sh);
  if (!info) return { error: "ヘッダ(日付)が見つからない" };

  var targetDate = (body.date || "").toString().replace(/-/g, "/");
  var shortDate = targetDate.slice(5); // "06/07"

  var lastRow = sh.getLastRow();
  var limit = body.limit || 200;
  var startRow = Math.max(info.hRow + 1, lastRow - limit + 1);
  if (startRow > lastRow) return { ok: true, rows: [], headers: info.headers };

  var data = sh.getRange(startRow, 1, lastRow - startRow + 1, info.headers.length).getValues();
  var results = [];
  var dateColIdx = _findCol(info.headers, "日付");

  for (var i = 0; i < data.length; i++) {
    var cellVal = data[i][dateColIdx];
    var cellStr = "";
    if (cellVal instanceof Date) {
      cellStr = Utilities.formatDate(cellVal, "Asia/Tokyo", "yyyy/MM/dd");
    } else {
      cellStr = String(cellVal).trim().replace(/-/g, "/");
    }
    if (!targetDate || cellStr === targetDate || cellStr.slice(-5) === shortDate) {
      var rowObj = {};
      for (var c = 0; c < info.headers.length; c++) {
        var v = data[i][c];
        rowObj[info.headers[c]] = (v instanceof Date) ? Utilities.formatDate(v, "Asia/Tokyo", "HH:mm") : String(v == null ? "" : v);
      }
      rowObj._row = startRow + i;
      results.push(rowObj);
    }
  }
  return { ok: true, rows: results, headers: info.headers, scanned: data.length };
}

// ── updateDb02: キーワード/行番号で行を特定して列を上書き（日付/時刻も整形）──
// body: { date, keyword, row_number, updates: {列名:値} }
function updateDb02(body) {
  const ss = SpreadsheetApp.openById((body && body.hostId) || COCKPIT_ID);
  var sh = _findDb02Sheet(ss);
  if (!sh) return { error: "02_作業DBが見つからない" };
  var info = _getHeaders(sh);
  if (!info) return { error: "ヘッダ(日付)が見つからない" };
  var H = info.headers;

  var targetRowNum = body.row_number || null;
  var targetDate = (body.date || "").toString().replace(/-/g, "/");
  var shortDate = targetDate.slice(-5);
  var keyword = (body.keyword || "").toString();
  var updates = body.updates || {};
  var dateColIdx = _findCol(H, "日付");
  var contentColIdx = _findCol(H, "内容");

  var matchedRow = -1;
  if (targetRowNum) {
    matchedRow = targetRowNum;
  } else {
    var lastRow = sh.getLastRow();
    var startRow = info.hRow + 1;
    var chunkSize = 500;
    outer: for (var s = startRow; s <= lastRow; s += chunkSize) {
      var end = Math.min(s + chunkSize - 1, lastRow);
      var data = sh.getRange(s, 1, end - s + 1, H.length).getValues();
      for (var i = 0; i < data.length; i++) {
        var cellVal = data[i][dateColIdx];
        var cellStr = "";
        if (cellVal instanceof Date) {
          cellStr = Utilities.formatDate(cellVal, "Asia/Tokyo", "yyyy/MM/dd");
        } else {
          cellStr = String(cellVal).trim().replace(/-/g, "/");
        }
        var dateMatch = !targetDate || cellStr === targetDate || cellStr.slice(-5) === shortDate;
        var content = String(data[i][contentColIdx] == null ? "" : data[i][contentColIdx]);
        var kwMatch = !keyword || content.indexOf(keyword) >= 0;
        if (dateMatch && kwMatch) { matchedRow = s + i; break outer; }
      }
    }
  }

  if (matchedRow < 0) return { ok: true, matched: false, note: "該当行なし", date: targetDate, keyword: keyword };

  var rowData = sh.getRange(matchedRow, 1, 1, H.length).getValues()[0];
  var applied = [];
  for (var key in updates) {
    var idx = _findCol(H, key);
    if (idx < 0) continue;
    if (_FORMULA_COLS.indexOf(H[idx]) >= 0) continue; // 所要4列は数式で持つので静的上書きしない
    var val = updates[key];
    if (H[idx] === "日付") { val = _toDate(val); }
    else if (_TIME_COLS.indexOf(H[idx]) >= 0) { var t = _toTimeFrac(val); if (t !== null) val = t; }
    rowData[idx] = val; applied.push(key);
  }
  sh.getRange(matchedRow, 1, 1, H.length).setValues([rowData]);

  // 書式の再保証
  var di = _findColExact(H, "日付"); if (di >= 0) sh.getRange(matchedRow, di + 1).setNumberFormat("MM/DD");
  _TIME_COLS.forEach(function (n) { var ix = _findColExact(H, n); if (ix >= 0) sh.getRange(matchedRow, ix + 1).setNumberFormat("HH:mm"); });
  _COUNT_COLS.forEach(function (n) { var ix = _findColExact(H, n); if (ix >= 0) sh.getRange(matchedRow, ix + 1).setNumberFormat("0"); });

  // 行に所要数式が無ければ補完（古い行・手追加行の救済）
  var iPlan = _findColExact(H, "予定所要(分)"), iAct = _findColExact(H, "実所要(分)");
  var needFormula = false;
  if (iPlan >= 0 && !String(sh.getRange(matchedRow, iPlan + 1).getFormula())) needFormula = true;
  if (iAct >= 0 && !String(sh.getRange(matchedRow, iAct + 1).getFormula())) needFormula = true;
  if (needFormula) {
    _writeDurationFormulas(sh, matchedRow,
      _findColExact(H, "予定開始"), _findColExact(H, "予定終了"),
      _findColExact(H, "実開始"), _findColExact(H, "実終了"),
      iPlan, iAct, _findColExact(H, "予実差分(分)"), _findColExact(H, "達成率%"));
  }

  return { ok: true, matched: true, updated_row: matchedRow, applied: applied };
}

// ── deleteDb02Rows: 指定行番号リストを削除（降順で処理） ─────────────────
function deleteDb02Rows(body) {
  const ss = SpreadsheetApp.openById((body && body.hostId) || COCKPIT_ID);
  var sh = _findDb02Sheet(ss);
  if (!sh) return { error: "02_作業DBが見つからない" };
  var rows = (body.row_numbers || []).slice().sort(function(a,b){ return b-a; });
  var deleted = [];
  for (var i = 0; i < rows.length; i++) {
    var rn = parseInt(rows[i]);
    if (rn > 1) { sh.deleteRow(rn); deleted.push(rn); }
  }
  return { ok: true, deleted: deleted };
}

// ── fixQrFormulas: 全データ行の所要4列を数式で上書き（既存行の一括修復用）──
// body: { from_row(省略時=ヘッダ+1) }  ※A列(日付)もMM/DDに揃える
function fixQrFormulas(body) {
  const ss = SpreadsheetApp.openById((body && body.hostId) || COCKPIT_ID);
  var sh = _findDb02Sheet(ss);
  if (!sh) return { error: "02_作業DBが見つからない" };
  var info = _getHeaders(sh);
  if (!info) return { error: "ヘッダ(日付)が見つからない" };
  var H = info.headers;

  var iKS = _findColExact(H, "予定開始"), iKE = _findColExact(H, "予定終了");
  var iMS = _findColExact(H, "実開始"),   iME = _findColExact(H, "実終了");
  var iPlan = _findColExact(H, "予定所要(分)"), iAct = _findColExact(H, "実所要(分)");
  var iDiff = _findColExact(H, "予実差分(分)"), iRate = _findColExact(H, "達成率%");
  if (iPlan < 0 || iAct < 0) return { error: "予定所要(分)/実所要(分)列が見つからない" };

  var startRow = (body && body.from_row) ? parseInt(body.from_row) : info.hRow + 1;
  var lastRow  = sh.getLastRow();
  if (startRow > lastRow) return { ok: true, fixed: 0, note: "データ行なし" };

  var count = 0;
  for (var r = startRow; r <= lastRow; r++) {
    _writeDurationFormulas(sh, r, iKS, iKE, iMS, iME, iPlan, iAct, iDiff, iRate);
    count++;
  }
  // A列(日付)：テキスト日付→日付値に変換してから MM/DD に統一（既存の崩れも修復）
  var iDate = _findColExact(H, "日付");
  if (iDate >= 0) {
    var dr = sh.getRange(startRow, iDate + 1, lastRow - startRow + 1, 1);
    var dv = dr.getValues();
    for (var k = 0; k < dv.length; k++) {
      var s = dv[k][0];
      if (typeof s === "string") { var m = String(s).match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/); if (m) dv[k][0] = new Date(+m[1], +m[2] - 1, +m[3]); }
    }
    dr.setValues(dv);
    dr.setNumberFormat("MM/DD");
  }
  // 時刻列：テキスト時刻→時刻値に変換（#VALUE!エラーの根治）してから HH:mm
  _TIME_COLS.forEach(function (n) {
    var ix = _findColExact(H, n); if (ix < 0) return;
    var tr = sh.getRange(startRow, ix + 1, lastRow - startRow + 1, 1), tv = tr.getValues();
    for (var k = 0; k < tv.length; k++) {
      var s = tv[k][0];
      if (typeof s === "string") { var m = String(s).match(/^(\d{1,2}):(\d{2})$/); if (m) tv[k][0] = (parseInt(m[1], 10) * 60 + parseInt(m[2], 10)) / 1440; }
    }
    tr.setValues(tv); tr.setNumberFormat("HH:mm");
  });
  // 件数列の書式も全行で再保証
  _COUNT_COLS.forEach(function (n) { var ix = _findColExact(H, n); if (ix >= 0) sh.getRange(startRow, ix + 1, lastRow - startRow + 1, 1).setNumberFormat("0"); });
  return { ok: true, fixed: count, from_row: startRow, to_row: lastRow };
}

// ── appendCustomer: 顧客マスターに1行追加 ─────────────────────────────────
function appendCustomer(body) {
  const ss = SpreadsheetApp.openById(COCKPIT_ID);
  const sh = ss.getSheets().find(function(s) { return s.getName().indexOf("顧客マスター") >= 0; });
  if (!sh) return { error: "顧客マスターが見つからない" };
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const lastRow = sh.getLastRow();
  var nextNum = 9;
  var colNo = -1;
  for (var c = 0; c < headers.length; c++) { if (headers[c] === "顧客番号") { colNo = c; break; } }
  if (colNo >= 0 && lastRow > 1) {
    var nos = sh.getRange(2, colNo + 1, lastRow - 1, 1).getValues().map(function(r){ return r[0]; });
    nos.forEach(function(v) { var m = String(v).match(/H(\d+)/); if (m) nextNum = Math.max(nextNum, parseInt(m[1]) + 1); });
  }
  var row = body.row || {};
  if (!row["顧客番号"]) row["顧客番号"] = "H" + String(nextNum).padStart(3, "0");
  var out = headers.map(function(h) { return row[h] !== undefined ? row[h] : ""; });
  sh.getRange(lastRow + 1, 1, 1, headers.length).setValues([out]);
  return { ok: true, appended_row: lastRow + 1, hNum: row["顧客番号"], headers: headers };
}

// ── ユーティリティ ────────────────────────────────────────────────────────
function _findDb02Sheet(ss) {
  var a = ss.getSheets();
  for (var i = 0; i < a.length; i++) { if (a[i].getName().indexOf("作業") >= 0) return a[i]; }
  return null;
}

function _getHeaders(sh) {
  var lc = sh.getLastColumn();
  var scan = sh.getRange(1, 1, Math.min(8, Math.max(1, sh.getLastRow())), lc).getValues();
  for (var r = 0; r < scan.length; r++) {
    var rv = [];
    for (var c = 0; c < scan[r].length; c++) rv.push((scan[r][c] == null ? "" : scan[r][c]).toString().trim());
    if (rv.indexOf("日付") >= 0) return { hRow: r + 1, headers: rv };
  }
  return null;
}

function _findCol(headers, key) {
  for (var c = 0; c < headers.length; c++) { if (headers[c] === key) return c; }
  for (var c = 0; c < headers.length; c++) {
    if (headers[c] && (headers[c].indexOf(key) >= 0 || key.indexOf(headers[c]) >= 0)) return c;
  }
  return -1;
}

function _findColExact(headers, key) {
  for (var c = 0; c < headers.length; c++) { if (headers[c] === key) return c; }
  return -1;
}

function _colLetter(n) {
  var s = '';
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// "2026-06-27"/"2026/6/27" → Dateオブジェクト（パース不可ならそのまま）
function _toDate(v) {
  if (v instanceof Date) return v;
  var s = String(v).trim();
  var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return v;
}

// "13:15"/"9:00" → 時刻シリアル(0〜1の小数)。時刻でなければnull
function _toTimeFrac(v) {
  if (typeof v === "number") return v;
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) { return (parseInt(m[1], 10) * 60 + parseInt(m[2], 10)) / 1440; }
  return null;
}
