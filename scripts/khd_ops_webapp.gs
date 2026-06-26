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

// ── appendDb02: 1行追記 ────────────────────────────────────────────────────
function appendDb02(body) {
  const ss = SpreadsheetApp.openById((body && body.hostId) || COCKPIT_ID);
  var sh = _findDb02Sheet(ss);
  if (!sh) return { error: "02_作業DBが見つからない" };
  var info = _getHeaders(sh);
  if (!info) return { error: "ヘッダ(日付)が見つからない", sheet: sh.getName() };
  if (body.headers_only) return { ok: true, sheet: sh.getName(), header_row: info.hRow, headers: info.headers };
  var row = body.row || {}, out = [], unmatched = [];
  for (var c = 0; c < info.headers.length; c++) out.push("");

  // Q(予実差分(分))・R(達成率%)は数式で埋めるため、Pythonからの静的値を無視
  // ※完全一致で検索（部分一致だと近似列に誤マッチする）
  var qIdx = _findColExact(info.headers, "予実差分(分)");
  var rIdx = _findColExact(info.headers, "達成率%");

  for (var key in row) {
    var idx = _findCol(info.headers, key);
    if (idx < 0) { unmatched.push(key); continue; }
    if (idx === qIdx || idx === rIdx) continue; // 数式列はスキップ
    out[idx] = row[key];
  }

  // 先頭挿入（ヘッダー直下の行2に挿入して最新が上に来るようにする）
  var newRow = info.hRow + 1;
  sh.insertRowBefore(newRow);
  // 日付列はテキスト書式を強制（Sheetsが "2026-06-08" を時刻に自動変換するのを防ぐ）
  var dateColIdx = _findColExact(info.headers, "日付");
  if (dateColIdx >= 0) {
    sh.getRange(newRow, dateColIdx + 1).setNumberFormat('@');
  }
  sh.getRange(newRow, 1, 1, info.headers.length).setValues([out]);

  // O列(予定所要(分))・P列(実所要(分))を使って Q・R に数式を挿入
  // ※完全一致キーを使う（"実所要"だけだと"実所要分"(M列)に誤マッチするため）
  var oIdx = _findColExact(info.headers, "予定所要(分)");
  var pIdx = _findColExact(info.headers, "実所要(分)");
  if (oIdx >= 0 && pIdx >= 0) {
    var oCol = _colLetter(oIdx + 1);
    var pCol = _colLetter(pIdx + 1);
    // Q: 予実差分 = 予定 - 実（どちらかが空なら空白）
    if (qIdx >= 0) {
      sh.getRange(newRow, qIdx + 1).setFormula(
        '=IF(OR(' + oCol + newRow + '="",' + pCol + newRow + '=""),"",'+oCol+newRow+'-'+pCol+newRow+')'
      );
    }
    // R: 達成率% = 実/予定×100（実が空または0なら空白）
    if (rIdx >= 0) {
      sh.getRange(newRow, rIdx + 1).setFormula(
        '=IF(OR(' + pCol + newRow + '="",' + pCol + newRow + '=0,' + oCol + newRow + '=""),"",ROUND('+pCol+newRow+'/'+oCol+newRow+'*100,1))'
      );
    }
  }

  return { ok: true, appended_row: newRow, sheet: sh.getName(), headers: info.headers, unmatched: unmatched };
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
  // "2026/06/07" → 短縮形 "06/07" も許容
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
    // フル一致 or 短縮一致
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

// ── updateDb02: キーワードで行を特定して列を上書き ────────────────────────
// body: { date, keyword, row_number (行番号直指定), updates: {列名:値} }
function updateDb02(body) {
  const ss = SpreadsheetApp.openById((body && body.hostId) || COCKPIT_ID);
  var sh = _findDb02Sheet(ss);
  if (!sh) return { error: "02_作業DBが見つからない" };
  var info = _getHeaders(sh);
  if (!info) return { error: "ヘッダ(日付)が見つからない" };

  var targetRowNum = body.row_number || null;
  var targetDate = (body.date || "").toString().replace(/-/g, "/");
  var shortDate = targetDate.slice(-5);
  var keyword = (body.keyword || "").toString();
  var updates = body.updates || {};
  var dateColIdx = _findCol(info.headers, "日付");
  var contentColIdx = _findCol(info.headers, "内容");

  var matchedRow = -1;
  if (targetRowNum) {
    matchedRow = targetRowNum;
  } else {
    // スキャンして一致行を探す
    var lastRow = sh.getLastRow();
    var startRow = info.hRow + 1;
    var chunkSize = 500;
    outer: for (var s = startRow; s <= lastRow; s += chunkSize) {
      var end = Math.min(s + chunkSize - 1, lastRow);
      var data = sh.getRange(s, 1, end - s + 1, info.headers.length).getValues();
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

  // 対象行を更新
  var rowData = sh.getRange(matchedRow, 1, 1, info.headers.length).getValues()[0];
  var applied = [];
  for (var key in updates) {
    var idx = _findCol(info.headers, key);
    if (idx >= 0) { rowData[idx] = updates[key]; applied.push(key); }
  }
  sh.getRange(matchedRow, 1, 1, info.headers.length).setValues([rowData]);
  return { ok: true, matched: true, updated_row: matchedRow, applied: applied };
}

// ── deleteDb02Rows: 指定行番号リストを削除（降順で処理） ─────────────────
// body: { row_numbers: [1024, 1025, ...] }
function deleteDb02Rows(body) {
  const ss = SpreadsheetApp.openById((body && body.hostId) || COCKPIT_ID);
  var sh = _findDb02Sheet(ss);
  if (!sh) return { error: "02_作業DBが見つからない" };
  var rows = (body.row_numbers || []).slice().sort(function(a,b){ return b-a; }); // 降順
  var deleted = [];
  for (var i = 0; i < rows.length; i++) {
    var rn = parseInt(rows[i]);
    if (rn > 1) { sh.deleteRow(rn); deleted.push(rn); }
  }
  return { ok: true, deleted: deleted };
}

// ── fixQrFormulas: 全データ行のQ(予実差分)・R(達成率%)を数式で上書き ────────
// body: { host_id, from_row(省略時=ヘッダ+1) }
function fixQrFormulas(body) {
  const ss = SpreadsheetApp.openById((body && body.hostId) || COCKPIT_ID);
  var sh = _findDb02Sheet(ss);
  if (!sh) return { error: "02_作業DBが見つからない" };
  var info = _getHeaders(sh);
  if (!info) return { error: "ヘッダ(日付)が見つからない" };

  // 完全一致で列を特定（部分一致だと誤マッチするため）
  var oIdx = _findColExact(info.headers, "予定所要(分)");
  var pIdx = _findColExact(info.headers, "実所要(分)");
  var qIdx = _findColExact(info.headers, "予実差分(分)");
  var rIdx = _findColExact(info.headers, "達成率%");
  if (oIdx < 0 || pIdx < 0) return { error: "予定所要(分)/実所要(分)列が見つからない" };
  if (qIdx < 0 && rIdx < 0) return { error: "予実差分(分)/達成率%列が見つからない" };

  var oCol = _colLetter(oIdx + 1);
  var pCol = _colLetter(pIdx + 1);
  var startRow = (body && body.from_row) ? parseInt(body.from_row) : info.hRow + 1;
  var lastRow  = sh.getLastRow();
  if (startRow > lastRow) return { ok: true, fixed: 0, note: "データ行なし" };

  var count = 0;
  for (var r = startRow; r <= lastRow; r++) {
    if (qIdx >= 0) {
      sh.getRange(r, qIdx + 1).setFormula(
        '=IF(OR(' + oCol + r + '="",' + pCol + r + '=""),"",'+oCol+r+'-'+pCol+r+')'
      );
    }
    if (rIdx >= 0) {
      sh.getRange(r, rIdx + 1).setFormula(
        '=IF(OR(' + pCol + r + '="",' + pCol + r + '=0,' + oCol + r + '=""),"",ROUND('+pCol+r+'/'+oCol+r+'*100,1))'
      );
    }
    count++;
  }
  return { ok: true, fixed: count, from_row: startRow, to_row: lastRow };
}

// ── appendCustomer: 顧客マスターに1行追加 ─────────────────────────────────
// body: { row: {列名:値} }  ※顧客番号は省略時に自動採番(H001〜)
function appendCustomer(body) {
  const ss = SpreadsheetApp.openById(COCKPIT_ID);
  const sh = ss.getSheets().find(function(s) { return s.getName().indexOf("顧客マスター") >= 0; });
  if (!sh) return { error: "顧客マスターが見つからない" };
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const lastRow = sh.getLastRow();
  // 自動採番
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
  // 部分一致フォールバック
  for (var c = 0; c < headers.length; c++) {
    if (headers[c] && (headers[c].indexOf(key) >= 0 || key.indexOf(headers[c]) >= 0)) return c;
  }
  return -1;
}

// 完全一致でヘッダ列を検索（_findColの部分一致誤マッチ防止用）
function _findColExact(headers, key) {
  for (var c = 0; c < headers.length; c++) { if (headers[c] === key) return c; }
  return -1;
}

// 列番号(1始まり) → 列文字(A, B, ..., Z, AA, ...)
function _colLetter(n) {
  var s = '';
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
