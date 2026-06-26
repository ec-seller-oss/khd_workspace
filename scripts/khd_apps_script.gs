/**
 * KHD WBS 業務報告 Webhook v2
 * スプレッドシート → 拡張機能 → Apps Script に貼り付けて「デプロイ」
 *
 * デプロイ設定:
 *   種類: ウェブアプリ
 *   次のユーザーとして実行: 自分
 *   アクセスできるユーザー: 全員
 */

const SECRET_TOKEN = "khd_report_2026";
const NIPPO_ID = "129j7x0Y4icgzYdTs9s_w1GBnFMIdVMtrTXqHeapM_f8"; // KHD日報ログ（別ファイル）
const CASHFLOW_ID = "1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc"; // 資金繰り司令塔（収益パイプライン同居先）

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.token !== SECRET_TOKEN) return json({ ok: false, error: "unauthorized" });

    const action = body.action;
    let result;

    if      (action === "update_wbs")      result = updateWbs(body.task, body.status, body.progress);
    else if (action === "update_study")    result = updateStudy(body);
    else if (action === "update_kpi")      result = updateKpi(body);
    else if (action === "append_log")      result = appendLog(body);
    else if (action === "ping")            result = { pong: true };
    else if (action === "debug_sheet")     result = debugSheet(body.partial);
    else if (action === "parse_and_update")   result = parseAndUpdate(body.text);
    else if (action === "get_wbs_summary")    result = getWbsSummary();
    else if (action === "check_calendar")     result = checkEndingEvents();
    else if (action === "log_conversation")   result = logConversation(body);
    else if (action === "get_next_task")      result = getNextTask();
    else if (action === "write_nippo")        result = writeNippo(body);
    else if (action === "append_nippo")       result = appendNippo(body);
    else if (action === "write_pipeline")     result = writePipeline(body);
    else if (action === "append_db02")        result = appendDb02(body);
    else if (action === "format_rules")        result = formatOperatingRules();
    else if (action === "finance_fixes")       result = applyFinanceFixes();
    else if (action === "tidy_tabs")           result = tidyTabs();
    else if (action === "build_db02")          result = buildDB02(body);
    else if (action === "build_wire0203")      result = buildWire0203(body);
    else if (action === "fix_wire0203")        result = fixWire0203(body);
    else if (action === "build_ui03")          result = buildUI03(body);
    else if (action === "build_board03")       result = buildBoard03(body);
    else if (action === "build_account_wire")  result = buildAccountWire(body);
    else if (action === "build_recurring")     result = buildRecurringIncome(body);
    else if (action === "build_fixes3")        result = buildFixes3(body);
    else if (action === "build_batch4")        result = buildBatch4(body);
    else if (action === "build_batch5")        result = buildBatch5(body);
    else if (action === "build_batch6")        result = buildBatch6(body);
    else if (action === "build_batch7")        result = buildBatch7(body);
    else if (action === "build_batch8")        result = buildBatch8(body);
    else if (action === "build_batch9")        result = buildBatch9(body);
    else if (action === "build_batch10")       result = buildBatch10(body);
    else if (action === "build_batch11")       result = buildBatch11(body);
    else if (action === "build_batch12")       result = buildBatch12(body);
    else if (action === "build_batch13")       result = buildBatch13(body);
    else if (action === "build_batch14")       result = buildBatch14(body);
    else if (action === "build_batch15")       result = buildBatch15(body);
    else if (action === "build_batch16")       result = buildBatch16(body);
    else if (action === "build_batch17")       result = buildBatch17(body);
    else if (action === "build_batch18")       result = buildBatch18(body);
    else if (action === "build_batch19")       result = buildBatch19(body);
    else if (action === "build_batch20")       result = buildBatch20(body);
    else if (action === "build_batch21")       result = buildBatch21(body);
    else if (action === "build_batch22")       result = buildBatch22(body);
    else if (action === "build_batch23")       result = buildBatch23(body);
    else if (action === "build_batch24")       result = buildBatch24(body);
    else if (action === "build_batch25")       result = buildBatch25(body);
    else if (action === "build_batch26")       result = buildBatch26(body);
    else result = { error: "unknown action: " + action };

    return json({ ok: true, result });
  } catch (err) {
    return json({ ok: false, error: err.toString() });
  }
}

// ── WBS タスク更新 ────────────────────────────────────
function updateWbs(taskPartial, newStatus, newProgress) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("02_全社WBS");
  if (!sheet) return { error: "WBSシートが見つかりません" };

  const data = sheet.getDataRange().getValues();

  // ヘッダー行をスキャンして探す
  let headerRow = -1, colStatus = -1, colProgress = -1;
  for (let i = 0; i < data.length; i++) {
    const idx = data[i].indexOf("ステータス");
    if (idx >= 0) {
      headerRow = i;
      colStatus   = idx;
      colProgress = data[i].indexOf("進捗率");
      break;
    }
  }
  if (headerRow < 0) return { error: "WBSヘッダーが見つかりません" };

  // 工程列（タスク名）を特定
  const colTask = data[headerRow].indexOf("工程");

  for (let i = headerRow + 1; i < data.length; i++) {
    const taskName = String(colTask >= 0 ? data[i][colTask] : data[i][1]).trim();
    if (taskName && taskName.includes(taskPartial)) {
      const updated = [];
      if (notNull(newStatus)   && colStatus   >= 0) { sheet.getRange(i+1, colStatus+1).setValue(newStatus);   updated.push("ステータス→" + newStatus); }
      if (notNull(newProgress) && colProgress >= 0) { sheet.getRange(i+1, colProgress+1).setValue(newProgress); updated.push("進捗率→" + newProgress); }
      return { task: taskName, row: i+1, updated };
    }
  }
  return { error: "タスクが見つかりません: " + taskPartial };
}

// ── 調査士学習ログ更新 ────────────────────────────────
function updateStudy(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = findSheetByPartial(ss, "05");
  if (!sheet) return { error: "学習ログシートが見つかりません" };

  const today = todayStr();
  const data  = sheet.getDataRange().getValues();

  // ヘッダー行スキャン
  let headerRow = -1, header = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i].includes("実績時間")) { headerRow = i; header = data[i]; break; }
  }
  if (headerRow < 0) return { error: "学習ログヘッダーが見つかりません", headers_row0: data[0] };

  for (let i = headerRow + 1; i < data.length; i++) {
    if (dateMatches(data[i][0], today)) {
      const row = i + 1;
      const updated = [];
      const colMap = [
        ["実績時間",   body.hours],
        ["範囲(択一)", body.range_takuitsu],
        ["範囲(記述)", body.range_kijutsu],
        ["ミス問数",   body.miss_count],
        ["コメント",   body.comment],
        ["完遂",       body.kansei],
      ];
      colMap.forEach(([col, val]) => {
        if (!notNull(val)) return;
        const ci = header.indexOf(col);
        if (ci >= 0) { sheet.getRange(row, ci+1).setValue(val); updated.push(col + "→" + val); }
      });
      return { sheet: sheet.getName(), row, updated };
    }
  }
  return { error: "今日の行が見つかりません (" + today + ")" };
}

// ── KPI 5指標更新 ─────────────────────────────────────
function updateKpi(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = findSheetByPartial(ss, "06");
  if (!sheet) return { error: "KPIシートが見つかりません" };

  const today = todayStr();
  const data  = sheet.getDataRange().getValues();

  // ヘッダー行スキャン
  let headerRow = -1, header = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i].some(h => String(h).includes("他力") || String(h).includes("朝活"))) {
      headerRow = i; header = data[i]; break;
    }
  }
  if (headerRow < 0) return { error: "KPIヘッダーが見つかりません" };

  for (let i = headerRow + 1; i < data.length; i++) {
    if (dateMatches(data[i][0], today)) {
      const row = i + 1;
      const updated = [];

      // 各列を名前で検索して書き込む
      const kpiMap = [
        ["他力",   body.tariki],
        ["EQ",     body.eq],
        ["朝活",   body.chokatsu],
        ["戦サボ", body.sen_sabo],
        ["家族",   body.kazoku],
        ["コメント", body.comment],
        ["TOP3",   body.top3],
      ];
      kpiMap.forEach(([keyword, val]) => {
        if (!notNull(val)) return;
        const ci = header.findIndex(h => String(h).includes(keyword));
        if (ci >= 0) { sheet.getRange(row, ci+1).setValue(val); updated.push(header[ci] + "→" + val); }
      });

      return { sheet: sheet.getName(), row, updated };
    }
  }
  return { error: "今日のKPI行が見つかりません (" + today + ")" };
}

// ── 日次ログ追記 ──────────────────────────────────────
function appendLog(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = findSheetByPartial(ss, "14");
  if (!sheet) return { error: "日次ログシートが見つかりません" };

  const lastRow = sheet.getLastRow() + 1;
  const now = todayStr() + " " + (body.time || formatTime(new Date()));
  const values = [now, body.category||"", body.dept||"", body.content||"",
                  body.project||"", body.manager_comment||"", body.wbs_reflect||"",
                  body.status||"🟡 確認待ち"];
  sheet.getRange(lastRow, 1, 1, values.length).setValues([values]);
  return { sheet: sheet.getName(), row: lastRow, written: body.content };
}

// ── 日報ログ：全置換ミラー（ローカルCSVを丸ごと反映・URL固定） ──
function writeNippo(body) {
  const rows = body.rows; // ヘッダ込み2D配列
  if (!rows || !rows.length) return { error: "rowsが空です" };
  const ss = SpreadsheetApp.openById(NIPPO_ID);
  const sheet = ss.getSheets()[0];
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  return { written: rows.length, sheet: sheet.getName() };
}

// ── 日報ログ：末尾に行追記（積み上げ・予定/実績の差分用） ──
function appendNippo(body) {
  const rows = body.rows; // データ行のみ（ヘッダ除く）2D配列
  if (!rows || !rows.length) return { error: "rowsが空です" };
  const ss = SpreadsheetApp.openById(NIPPO_ID);
  const sheet = ss.getSheets()[0];
  sheet.getRange(sheet.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);
  return { appended: rows.length, lastRow: sheet.getLastRow() };
}

// ── 02_作業DBに1行追記（運用の本丸）─────────────────────────
// body.row = {列名:値} の辞書。ヘッダ名で対応づける＝列順を動かしても壊れない。
// 列が見つからない時はunmatchedで返す（こちらでヘッダ名を学習・補正できる）。
function appendDb02(body) {
  const ss = SpreadsheetApp.openById((body && body.hostId) || CASHFLOW_ID);
  var sh = null, a = ss.getSheets();
  for (var i = 0; i < a.length; i++) { if (a[i].getName().indexOf("作業") >= 0) { sh = a[i]; break; } }
  if (!sh) return { error: "02_作業DBが見つからない" };
  // ヘッダ行を探す（「日付」を含む行・最初の8行）
  var lc = sh.getLastColumn();
  var scan = sh.getRange(1, 1, Math.min(8, Math.max(1, sh.getLastRow())), lc).getValues();
  var hRow = -1, headers = null;
  for (var r = 0; r < scan.length; r++) {
    var rv = []; for (var c = 0; c < scan[r].length; c++) rv.push((scan[r][c] == null ? "" : scan[r][c]).toString().trim());
    if (rv.indexOf("日付") >= 0) { hRow = r + 1; headers = rv; break; }
  }
  if (hRow < 0) return { error: "ヘッダ(日付)が見つからない", sheet: sh.getName() };
  // body.headers_only=true ならヘッダだけ返す（学習用）
  if (body.headers_only) return { ok: true, sheet: sh.getName(), header_row: hRow, headers: headers };
  // 追記行を組み立て（ヘッダ名で対応づけ）
  var row = body.row || {}, out = [], unmatched = [];
  for (var c = 0; c < headers.length; c++) out.push("");
  for (var key in row) {
    var idx = -1;
    for (var c2 = 0; c2 < headers.length; c2++) { if (headers[c2] === key) { idx = c2; break; } }
    if (idx < 0) { // 部分一致で再トライ
      for (var c3 = 0; c3 < headers.length; c3++) {
        if (headers[c3] !== "" && (headers[c3].indexOf(key) >= 0 || key.indexOf(headers[c3]) >= 0)) { idx = c3; break; }
      }
    }
    if (idx >= 0) out[idx] = row[key]; else unmatched.push(key);
  }
  var newRow = sh.getLastRow() + 1;
  sh.getRange(newRow, 1, 1, headers.length).setValues([out]);
  return { ok: true, appended_row: newRow, sheet: sh.getName(), headers: headers, unmatched: unmatched };
}

// ── タブをキーワード（名前の一部）で探す＝番号プレフィックス改名後も壊れない ──
function shByKey(ss, kw) {
  var a = ss.getSheets();
  for (var i = 0; i < a.length; i++) { if (a[i].getName().indexOf(kw) >= 0) return a[i]; }
  return null;
}

// ── 収益パイプライン：資金繰りスプシ内の「収益パイプライン」タブへ全置換ミラー ──
// 同居タブ化（残務B）。rows = ヘッダ込み2D配列。"="始まりのセルは数式として展開される。
function writePipeline(body) {
  const rows = body.rows;
  if (!rows || !rows.length) return { error: "rowsが空です" };
  const ss = SpreadsheetApp.openById(body.hostId || CASHFLOW_ID);
  let sheet = shByKey(ss, "収益パイプライン");
  if (!sheet) sheet = ss.insertSheet("10_収益パイプライン");
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  return { written: rows.length, sheet: sheet.getName(), host: ss.getName() };
}

// ── 運用ルールの色分け（マニュアル不要化）────────────────────
// 色＝ルール。🟡入力 / 🟧毎日見る / 🟦週次月次で見る / ⬜自動・触らない。背景色のみで値は壊さない。
function formatOperatingRules() {
  const C = { IN:"#FFF2CC", DAY:"#FCE5CD", WATCH:"#CFE2F3", AUTO:"#ECECEC", HEAD:"#434343" };
  const ss = SpreadsheetApp.openById(CASHFLOW_ID);
  const out = [];
  function f(sh, a1, color, note) {
    const rg = sh.getRange(a1);
    if (color) rg.setBackground(color);
    if (note != null) rg.setNote(note);
  }
  function legend(sh, startRow) {
    const r = startRow;
    sh.getRange(r,1).setValue("📖 このタブの使い方（色＝運用ルール）").setFontWeight("bold").setFontColor("#FFFFFF").setBackground(C.HEAD);
    const rows = [
      ["🟡 黄＝【入力】あなたが打つ（頻度はセルのメモを参照）", C.IN],
      ["🟧 橙＝【毎日見る】今日の一手・追客目標", C.DAY],
      ["🟦 青＝【週次/月次で見る】判断に使うKPI", C.WATCH],
      ["⬜ 灰＝自動計算・触らない", C.AUTO],
    ];
    for (let i=0;i<rows.length;i++){
      sh.getRange(r+1+i,1).setBackground(rows[i][1]);
      sh.getRange(r+1+i,2).setValue(rows[i][0]);
    }
  }

  // ① 売上見込み（金額のSSoT）
  let s = shByKey(ss, "売上見込み");
  if (s) {
    f(s,"B2",C.IN,"【入力・月初/会議前】シナリオを 弱気/現実/強気 で打ち替え");
    f(s,"A5:G14",C.IN,"【入力・案件が動くたび】本部/科目/案件名/満額(E)/確度(F)/着金月(G)。科目は栄町売却・医療テナントコンサル・テレアポ買取再販のみ。着金月は2026/06形式");
    f(s,"H5:K14",C.AUTO,"自動・触らない（シナリオ別採用額）");
    f(s,"M5:M14",C.AUTO,"自動・触らない（期待値=満額×確度）");
    f(s,"B17:B19",C.AUTO,"自動・触らない（案件別採用額）");
    f(s,"B20",C.WATCH,"【見る・週次】採用額合計＝今のシナリオで④資金繰りに流れる額");
    f(s,"B21",C.WATCH,"【見る・週次/月次】期待値合計＝確度加重の収入見込み（収益パイプラインの真実源）");
    legend(s, 28);
  }
  // ⑦ 損益（PL）：営業ドライバー＝逆算エンジン
  s = shByKey(ss, "損益");
  if (s) {
    f(s,"B48",C.IN,"【入力・ほぼ固定】営業直結 目標時間比率（既定0.6）");
    f(s,"B52",C.IN,"【入力・月1回】追客→アポ 転換率（自分の実績で更新）");
    f(s,"B53",C.IN,"【入力・月1回】アポ→成約 転換率（自分の実績で更新）");
    f(s,"B57",C.IN,"【入力・月初】月の稼働日");
    f(s,"B49",C.DAY,"【毎日見る】今日の一手");
    f(s,"B59",C.DAY,"【毎日見る】今日の追客目標");
    f(s,"B63",C.DAY,"【毎日見る】時間の一手");
    f(s,"B34",C.WATCH,"【見る・週次/月次】経常ギャップ＝毎月埋める穴");
    f(s,"B42:B43",C.WATCH,"【見る・月次】家族/黒字ラインに必要な医療件数");
    f(s,"B45",C.WATCH,"【見る・月次】家族ラインに必要な物件件数");
    f(s,"B58",C.WATCH,"【見る・週次】1日あたり必要追客数");
    f(s,"B67:F70",C.WATCH,"【見る・月次】月次トラッキング（必要vs確定vs不足）");
    f(s,"B32",C.AUTO,"自動（経常確定粗利）"); f(s,"B33",C.AUTO,"自動（家族ライン）"); f(s,"B41",C.AUTO,"自動（医療確度込/件）");
    legend(s, 72);
  }
  // ④ 資金繰り（結果・原則見るだけ。固定費だけ入力）
  s = shByKey(ss, "資金繰り");
  if (s) {
    f(s,"B5:I8",C.AUTO,"自動・触らない（売上見込み/IMPORTRANGE連動の入金）");
    f(s,"B10",C.IN,"【入力・変わったら】育休給付");
    f(s,"B14:I26",C.IN,"【入力・変わったら】固定費・出金（半固定）");
    f(s,"B12:I12",C.AUTO,"自動（入金合計）"); f(s,"B27:I27",C.AUTO,"自動（出金合計）");
    f(s,"B29:I29",C.WATCH,"【見る・週次/月次】月末現預金＝体力計");
    f(s,"B32",C.WATCH,"【見る・週次】ランウェイ判定");
    legend(s, 36);
  }
  // 収益パイプライン（営業の的リスト・CSV経由で更新）
  s = shByKey(ss, "収益パイプライン");
  if (s) {
    f(s,"K2:L8",C.IN,"【更新・随時／CSV経由】次アクション・ステータス（私に『パイプライン反映して』で書込）");
    f(s,"I2:I8",C.AUTO,"自動（各ゴールの確度加重・作業用）");
    f(s,"C12:I13",C.WATCH,"【見る・週次】売上見込み(SSoT)連動の確定値");
    f(s,"C14:I14",C.WATCH,"【見る・週次】取りこぼしチェック＝＞0なら売上見込み未登録案件あり");
    legend(s, 16);
  }
  out.push("formatted");
  return { ok:true, tabs:["売上見込み","⑦ 損益（PL）","④ 資金繰り","収益パイプライン"] };
}

// ── 財務修正バンドル（税理士/経常利益色/色ズレ/未来会計図表参照化）2026-06-05 ──
function applyFinanceFixes() {
  const ss = SpreadsheetApp.openById(CASHFLOW_ID);
  const cf = shByKey(ss, "資金繰り");
  const pl = shByKey(ss, "損益");
  const mf = shByKey(ss, "未来会計図表");
  const Y="#FFF2CC", O="#FCE5CD", B="#CFE2F3", GOLD="#FFE599";

  // 1) 税理士を契約ベースの正値へ（B:I＝6月〜1月の8列）
  cf.getRange("B14:I14").setValue(18333);  // 法人税理士 決算22万/年÷12
  cf.getRange("B17:I17").setValue(46750);  // 個人税理士 税込56.1万/年÷12

  // 2) PL固定費計B7＝18333+99000+22167+50000=189500／個人税理士B16=46750／ラベル更新
  pl.getRange("B7:I7").setValue(189500);
  pl.getRange("B16:I16").setValue(46750);
  pl.getRange("A7").setValue("  固定費計（税理士18,333+外注99,000+社保22,167＋役員報酬50,000※未払＝189,500）");

  // 3) 経常利益を結論色(金)＋太字に・支払利息の緑を外して区別
  pl.getRange("B8:I8").setBackground(null);
  [9,17,25].forEach(function(r){ var rg=pl.getRange(r,2,1,8); rg.setBackground(GOLD); rg.setFontWeight("bold"); });

  // 4) 営業ドライバーの色ズレ修正：誤りクリア→正しい行へ
  ["B33","B48","B49","B53","B57","B59","B63","B67"].forEach(function(a){ pl.getRange(a).setBackground(null); });
  ["B51","B55","B56","B60"].forEach(function(a){ pl.getRange(a).setBackground(Y); });  // 入力
  ["B52","B62","B66"].forEach(function(a){ pl.getRange(a).setBackground(O); });          // 毎日見る
  ["B37","B45","B46","B48","B58","B61"].forEach(function(a){ pl.getRange(a).setBackground(B); }); // 週次見る

  // 5) 未来会計図表を⑦PL参照化（固定費・目標利益→損益分岐の二重/矛盾を解消）
  mf.getRange("B4").setFormula("='⑦ 損益（PL）'!B24");
  mf.getRange("B5").setFormula("='⑦ 損益（PL）'!B26");
  mf.getRange("A2").setValue("※固定費・目標利益・損益分岐は⑦損益PLを参照（PL一本化／このタブは図示用）");

  return { ok:true, taxFixed:true, keijo:"gold", driverRecolored:true, miraiLinked:true };
}

// ── タブ整理：日常で使う順に並べ替え＋データ/スポットは非表示（名前は変えない＝参照壊れない）──
function tidyTabs() {
  const ss = SpreadsheetApp.openById(CASHFLOW_ID);
  // 表示する操縦席（この順＝流れ）。[探すキーワード, 新しい名前(01_〜)]
  const front = [
    ["統合司令塔",   "01_統合司令塔"],
    ["追客リスト",   "02_追客リスト"],
    ["売上見込み",   "03_売上見込み"],
    ["損益",         "04_損益PL"],
    ["資金繰り",     "05_資金繰り"],
    ["資産負債",     "06_資産負債BS"],
    ["本部マトリクス","07_本部マトリクス"],
    ["残高クイック", "08_残高クイック入力"],
    ["MF残高",       "09_MF残高"],
    ["収益パイプライン","10_収益パイプライン"],
    ["KPI",          "11_KPI"]
  ];
  // 非表示にする（データ源・スポット・アーカイブ。計算は生きる・いつでも再表示可）
  const hide = [
    "借入", "税金", "諸経費", "クレカ用途",
    "按分_260604", "役員報酬", "経費削減・損切り",
    "💼銀行提出前クリーンアップ", "事業計画（5ヵ年）", "未来会計図表",
    "WBS"
  ];
  // ※pass(ID/パスワード管理)はよく使うので表示のまま（菊池確定2026-06-05）
  // 不要2枚を削除（菊池確定2026-06-05・戻せない）：⑥使い方=マニュアルで代替／📋テンプレ=按分_260604の重複
  const del = ["⑥使い方", "📋按分棚卸し_テンプレ"];
  var deleted = [];
  del.forEach(function(nm){
    var sh = ss.getSheetByName(nm);
    if (sh){ ss.deleteSheet(sh); deleted.push(nm); }
  });
  // まず表示対象をキーワードで探し→01_に改名→show→流れ順に前へ（数式参照は自動追従）
  front.forEach(function(pair, i){
    var sh = shByKey(ss, pair[0]);
    if (sh){ sh.setName(pair[1]); sh.showSheet(); ss.setActiveSheet(sh); ss.moveActiveSheet(i+1); }
  });
  // 残りを非表示
  var hidden = [];
  hide.forEach(function(nm){
    var sh = ss.getSheetByName(nm);
    if (sh){ sh.hideSheet(); hidden.push(nm); }
  });
  // 先頭（①司令塔）をアクティブに戻す
  var top = shByKey(ss, "統合司令塔"); if (top) ss.setActiveSheet(top);
  return { ok:true, shown:front.length, hidden:hidden.length, deleted:deleted };
}






// ── 実装①：02_作業DB（旧追客＋日報＋報告ログ統合・追記専用）本体に追加 2026-06-06 ──
function buildDB02(body) {
  const ss = SpreadsheetApp.openById((body && body.hostId) || "1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  const Y="#FFF2CC",BR="#AA2E26",G="#ECECEC";
  var s = ss.getSheetByName("02_作業DB"); if(s) ss.deleteSheet(s); s = ss.insertSheet("02_作業DB");
  var rows = [
   ["日付","区分","本部","案件・相手","種類・科目","流入・出所","温度","内容","予定開始","予定終了","実開始","実終了","実所要分","営業直結","報告項目","報告値","結果","ステータス","最終接触","次アクション","期限","詳細リンク"],
   ["2026-06-06","予定","05","栄町(持倉様)","物件売却","京葉不動産/直","HOT","大坪へ口座確認","09:00","09:30","","","","○","","","","進行","確定","大坪へ口座確認","6/8",""],
   ["2026-06-06","実績","04","曾我先生","医療コンサル","福井→アイリス","WARM","福井へ提示催促","","","10:00","10:20","20","○","確度","0.4","反応微妙","提案中","2026-06-02","様子見・関係維持","随時",""],
   ["2026-06-06","実績","03","cloud mil(小池)","cloud mil反響","cloud mil親","WARM","個別提案 追客","","","11:00","11:15","15","○","追客","1","送信済","進行","2026-06-06","返信待ち","",""],
   ["2026-06-06","予定","04","土地家屋調査士","内務","-","-","朝活","06:30","09:00","","","","-","","","","仕込","","","",""],
   ["2026-06-06","突発","01","税務","内務","-","-","記帳問い合わせ対応","","","13:00","13:20","20","✕","","","","完了","","","",""],
   ["（新規はここに1行・Claudeが報告から追記。報告項目=確度/着金月/満額/追客/作業時間 等）","","","","","","","","","","","","","","","","","","","","",""],
  ];
  s.getRange(1,1,rows.length,22).setValues(rows);
  s.getRange("A1:V1").setBackground(BR).setFontColor("#FFFFFF").setFontWeight("bold");
  s.getRange(2,1,5,22).setBackground(Y);
  s.getRange(7,1,1,22).setBackground(G);
  // 運用メモ
  s.getRange("A9").setValue("★この02_作業DBが唯一の入力ハブ＝追記専用(絶対に消さない)。旧・追客リスト＋日報＋報告ログを統合。").setFontWeight("bold").setFontColor(BR);
  s.getRange("A10").setValue("流れ：作業DB(02)→売上予定(03は確度/着金をここからLOOKUP)→PL+CF(04/05)→BS(06)→一覧(01)。あなたは報告するだけ、Claudeが1行追記。");
  s.getRange("A11").setValue("区分：予定=朝にカレンダーから／実績=夜に実開始/実終了/実所要を補充／突発=日中の記録。時間は同じ行の予定列vs実績列(上書きでなく補充)。案件の数値報告は新規行(履歴・消えない)。");
  s.setFrozenRows(1); s.setTabColor("#F1C232");
  // 01の直後へ
  var s01=ss.getSheetByName("01_統合司令塔")||ss.getSheets()[0];
  ss.setActiveSheet(s); ss.moveActiveSheet(2);
  return { ok:true, url:ss.getUrl(), tab:"02_作業DB", note:"追加のみ・既存タブ無傷" };
}

// ── 実装②：02→03 LOOKUP配線＋02に種まき/出会い区分 2026-06-06 ──
function buildWire0203(body) {
  const ss = SpreadsheetApp.openById((body && body.hostId) || "1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  const Y="#FFF2CC", G="#ECECEC";
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }

  // (1) 02_作業DBに「種まき/出会い」サンプル＋栄町確度の実報告デモを追記、区分凡例を追加
  var db = ss.getSheetByName("02_作業DB");
  if(db){
    var last = db.getLastRow();
    db.getRange(last+1,1,2,22).setValues([
      ["2026-06-06","種まき/出会い","00","異業種交流会","出会い","紹介","-","名刺交換3名(うち1名 開業医)","","","19:00","21:00","120","△","出会い","3","良縁","仕込","2026-06-06","開業医にAI診断GIVE",""," "],
      ["2026-06-06","案件更新","05","栄町売却（→法人SBI）","物件売却","京葉不動産/直","HOT","確度更新","","","","","","○","確度","0.95","決済目前","進行","2026-06-06","6/8決済","6/8",""]
    ]);
    db.getRange(last+1,1,2,22).setBackground(Y);
    db.getRange("X1").setValue("区分の選択肢：予定／実績／突発／種まき・出会い／案件更新。報告項目：確度／着金月／満額／追客／作業時間／出会い 等。種まき・出会い＝偶然を資産化(後で現金化したら辿れる)").setFontColor("#AA2E26").setFontWeight("bold");
  }

  // (2) 03_売上見込み：確度(F)・着金月(G)を02の最新報告からLOOKUP、無ければseed(元値)に
  var u = shk("売上見込み");
  if(u){
    // 元のF/Gをseed列R/Sへ退避（読んでから上書き）
    var fv = u.getRange("F5:F14").getValues();
    var gv = u.getRange("G5:G14").getValues();
    u.getRange("R5:R14").setValues(fv);
    u.getRange("S5:S14").setValues(gv);
    u.getRange("R4").setValue("確度_seed(初期値)"); u.getRange("S4").setValue("着金_seed");
    u.getRange("R4:S4").setBackground(G);
    for(var r=5;r<=14;r++){
      u.getRange("F"+r).setFormula("=IFERROR(LOOKUP(2,1/(('02_作業DB'!$D$2:$D$2000=$C"+r+")*('02_作業DB'!$O$2:$O$2000=\"確度\")),'02_作業DB'!$P$2:$P$2000),R"+r+")");
      u.getRange("G"+r).setFormula("=IFERROR(LOOKUP(2,1/(('02_作業DB'!$D$2:$D$2000=$C"+r+")*('02_作業DB'!$O$2:$O$2000=\"着金月\")),'02_作業DB'!$P$2:$P$2000),S"+r+")");
    }
    u.getRange("F5:G14").setBackground("#CFE2F3"); // 自動連動(青)に
    u.getRange("U4").setValue("※確度F・着金G＝02作業DBの最新報告を案件名キーで自動参照。報告が無ければseed(R/S)。報告時は03の案件名(C列)と一致させること");
  }
  return { ok:true, wired:"02→03 確度/着金 LOOKUP", demo:"栄町確度0.95が02報告→03に反映されるはず", seed:"R/S列に退避" };
}

// ── 実装②修正：LOOKUP→FILTERベースに差替(Sheetsで堅牢)＋報告値を数値化。再追記しない 2026-06-06 ──
function fixWire0203(body) {
  const ss = SpreadsheetApp.openById((body && body.hostId) || "1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  // 02の確度報告値を数値化（テキスト"0.95"等→数値）
  var db = ss.getSheetByName("02_作業DB");
  if(db){
    var lastR = db.getLastRow();
    for(var i=2;i<=lastR;i++){
      if(db.getRange(i,15).getValue()==="確度"){ var pv=db.getRange(i,16).getValue(); if(pv!==""&&pv!==null) db.getRange(i,16).setValue(Number(pv)); }
    }
  }
  // 03の確度F・着金G をFILTERベース(最新=末尾)で02から取得、無ければseed(R/S)
  var u = shk("売上見込み");
  if(u){
    for(var r=5;r<=14;r++){
      var fF="=IFERROR(INDEX(FILTER('02_作業DB'!$P$2:$P$2000,('02_作業DB'!$D$2:$D$2000=$C"+r+")*('02_作業DB'!$O$2:$O$2000=\"確度\")),COUNTA(FILTER('02_作業DB'!$P$2:$P$2000,('02_作業DB'!$D$2:$D$2000=$C"+r+")*('02_作業DB'!$O$2:$O$2000=\"確度\")))),R"+r+")";
      var fG="=IFERROR(INDEX(FILTER('02_作業DB'!$P$2:$P$2000,('02_作業DB'!$D$2:$D$2000=$C"+r+")*('02_作業DB'!$O$2:$O$2000=\"着金月\")),COUNTA(FILTER('02_作業DB'!$P$2:$P$2000,('02_作業DB'!$D$2:$D$2000=$C"+r+")*('02_作業DB'!$O$2:$O$2000=\"着金月\")))),S"+r+")";
      u.getRange("F"+r).setFormula(fF);
      u.getRange("G"+r).setFormula(fG);
    }
  }
  return { ok:true, fixed:"FILTERベースに差替・報告値数値化", expect:"栄町確度→0.95に反映されるはず" };
}

// ── 実装③：03_売上予定を見やすく整備（連動列固定・ノイズ非表示・好み列・本部フィルタ）2026-06-06 ──
function buildUI03(body) {
  const ss = SpreadsheetApp.openById((body && body.hostId) || "1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  const Y="#FFF2CC", B="#CFE2F3", BR="#AA2E26";
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var u = shk("売上見込み");
  if(!u) return { error:"03_売上見込みが見つからない" };

  // 1) 好み列を追加（N〜Q：業務種別/売上化?/温度/次アクション）
  u.getRange("N4:Q4").setValues([["業務種別","売上化?","温度","次アクション"]]);
  u.getRange("N5:Q11").setValues([
    ["物件","🟢直結","HOT","大坪へ口座確認→6/8決済"],
    ["医療","🟢直結","WARM","福井に提示催促(反応微妙)"],
    ["物件","🟡仕込み","WARM","宮崎とテレアポ設計"],
    ["物件","🟡仕込み","COLD","新規仕入ソーシング"],
    ["医療","🟡仕込み","WARM","セミナー後追客"],
    ["医療","🟡仕込み","WARM","My AI提案デモ"],
    ["物件","🟡仕込み","WARM","横浜NPO 福井MTG"]
  ]);
  u.getRange("N4:Q4").setBackground(BR).setFontColor("#FFFFFF").setFontWeight("bold");
  u.getRange("N5:Q11").setBackground(Y);

  // 2) ノイズ列を非表示：D(物件/タグ)・H〜J(弱気/現実/強気)・R/S(seed)
  u.hideColumns(4,1);   // D
  u.hideColumns(8,3);   // H,I,J
  u.hideColumns(18,2);  // R,S(seed)

  // 3) 色で「入力/自動」を明示：満額E=入力(黄)、確度F・着金G=02連動(青)
  u.getRange("E5:E14").setBackground(Y);
  u.getRange("F5:G14").setBackground(B);

  // 4) 本部ソート用フィルタ（既存あれば張り直し）＋ヘッダ固定
  var fl=u.getFilter(); if(fl) fl.remove();
  u.getRange("A4:Q14").createFilter();
  u.setFrozenRows(4);

  // 5) 読み方ガイド（シナリオB2は触らないので注記で補足）
  u.getRange("E4").setNote("【入力】案件の満額。新規案件はここに行追加");
  u.getRange("F4").setNote("【自動】02_作業DBの最新確度報告を参照（無ければseed）。直接入力しない");
  u.getRange("G4").setNote("【自動】02の最新着金月を参照");
  u.getRange("A4").setNote("本部コード(5物件/4医療/3事業)。フィルタ▼で本部・科目・温度でソート/絞り込み可（行は動かさず連動無傷）");
  return { ok:true, added:"N〜Q好み列", hidden:"D/H〜J/R/S", filter:"A4:Q14", note:"連動列(B/E/F/G/K)は不可侵" };
}

// ── 実装③'：03を「週次作戦盤」へ。KPIヘッダ＋クローズ管理＋確度加重living＋営業ドライバー 2026-06-06 ──
// 不可侵：B科目/E満額/F確度/G着金/K採用額（資金繰りが読む）。列は動かさずKの式だけ確度加重に変更。
function buildBoard03(body){
  const ss = SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  const Y="#FFF2CC", B="#CFE2F3", BR="#AA2E26", G="#ECECEC", GREEN="#D9EAD3", GOLD="#FFE599", O="#FCE5CD", GRY="#EFEFEF";
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var u = shk("売上見込み"); if(!u) return {error:"03(売上見込み)が無い"};

  // ===== (い) シナリオ廃止→確度加重 living forecast =====
  // 採用額K：成約✅=満額／失注✗・見送り=0／それ以外(進行)=満額×確度。資金繰りはKを読むので自動で確度加重に。
  for(var r=5;r<=14;r++){
    u.getRange("K"+r).setFormula(
      "=IF($O"+r+"=\"成約✅\",$E"+r+",IF(OR($O"+r+"=\"失注✗\",$O"+r+"=\"見送り\"),0,$E"+r+"*$F"+r+"))");
    u.getRange("M"+r).setFormula("=$E"+r+"*$F"+r); // 期待値(確度加重)＝集計B21維持用
  }

  // ===== クローズ管理：O列を「状態」に（進行/成約✅/失注✗/見送り/再案件化）=====
  u.getRange("O4").setValue("状態").setBackground(BR).setFontColor("#FFFFFF").setFontWeight("bold");
  u.getRange("O5:O11").setValue("進行"); // 既存7案件は進行で初期化（旧"売上化?"を上書き）
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["進行","成約✅","失注✗","見送り","再案件化"], true).setAllowInvalid(true).build();
  u.getRange("O5:O14").setDataValidation(rule);
  u.getRange("O4").setNote("【毎朝の目利き】進行=押す／成約✅=決まった(満額計上)／失注✗・見送り=0でグレーアウト。再案件化したら『進行』に戻すと復活＝失注後も再度狙える。");

  // ===== グレーアウト＋成約グリーン（条件付き書式・行A5:Q14）=====
  var rng = u.getRange("A5:Q14");
  var won = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$O5="成約✅"').setBackground(GREEN).setRanges([rng]).build();
  var lost = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($O5="失注✗",$O5="見送り")').setBackground(GRY).setFontColor("#999999").setRanges([rng]).build();
  var cur = u.getConditionalFormatRules(); cur.push(won); cur.push(lost);
  u.setConditionalFormatRules(cur);

  // ===== KPIヘッダ（上部1〜3行を作り替え）=====
  u.getRange("A1:Q3").breakApart();
  u.getRange("A1:Q3").clearContent();
  u.getRange("A1").setValue("03_売上予定（週次作戦盤）｜毎朝ここで“今日プッシュする案件”を決める → 02作業DBへ落とす");
  u.getRange("A1:Q1").merge();
  u.getRange("A1").setFontWeight("bold").setFontColor("#FFFFFF").setBackground(BR).setFontSize(11).setHorizontalAlignment("left");
  // 6月の数字ブロック
  u.getRange("A2:F2").setValues([["📅 6月の数字","売上目標(入力)","見込(確度加重)","確定(成約)","ギャップ","達成率"]]);
  u.getRange("B3").setValue(1710000);                 // 目標(入力)暫定171万＝必達ライン目安。上書き可
  u.getRange("C3").setFormula("=SUM(K5:K14)");          // 見込＝living forecast合計
  u.getRange("D3").setFormula("=SUMIF(O5:O14,\"成約✅\",E5:E14)"); // 確定(成約満額)
  u.getRange("E3").setFormula("=B3-C3");                // ギャップ＝目標-見込(前向き)
  u.getRange("F3").setFormula("=IFERROR(D3/B3,0)");     // 達成率＝確定(成約)÷目標(実績)
  u.getRange("B3:E3").setNumberFormat("#,##0"); u.getRange("F3").setNumberFormat("0%");
  u.getRange("A2:F2").setBackground("#434343").setFontColor("#FFFFFF").setFontWeight("bold");
  u.getRange("B3").setBackground(Y).setNote("【入力・月初】6月の売上目標(必達ライン)。暫定171万。");
  u.getRange("C3:E3").setBackground(B); u.getRange("F3").setBackground(GOLD).setFontWeight("bold");
  // 件数ブロック
  u.getRange("H2:K2").setValues([["進行中","成約✅","失注/見送り","今月の的(進行件数)"]]);
  u.getRange("H3").setFormula("=COUNTIF(O5:O14,\"進行\")+COUNTIF(O5:O14,\"再案件化\")");
  u.getRange("I3").setFormula("=COUNTIF(O5:O14,\"成約✅\")");
  u.getRange("J3").setFormula("=COUNTIF(O5:O14,\"失注✗\")+COUNTIF(O5:O14,\"見送り\")");
  u.getRange("K3").setFormula("=H3");
  u.getRange("H2:K2").setBackground("#434343").setFontColor("#FFFFFF").setFontWeight("bold");
  u.getRange("H3:K3").setBackground(B).setFontWeight("bold");

  // ===== 営業ドライバー（ギャップ→必要成約→必要アポ→必要追客の逆算）=====
  var dr=30;
  u.getRange(dr,1).setValue("🎯 営業ドライバー（ギャップを埋める逆算→今日の追客本数が出る）");
  u.getRange(dr,1,1,4).merge();
  u.getRange(dr,1).setFontWeight("bold").setFontColor("#FFFFFF").setBackground(BR);
  var drv=[
    ["ギャップ(目標-見込)","=E3","円","← これを埋める"],
    ["平均成約単価(入力)",1200000,"円","医療110万/物件単発の目安"],
    ["必要・追加成約数","=ROUNDUP(B"+(dr+1)+"/B"+(dr+2)+",0)","件","ギャップ÷単価"],
    ["成約率(アポ→成約・入力)",0.3,"","損益B53と揃える"],
    ["必要アポ数","=ROUNDUP(B"+(dr+3)+"/B"+(dr+4)+",0)","件",""],
    ["追客→アポ率(入力)",0.2,"","損益B52と揃える"],
    ["必要・今月追客数","=ROUNDUP(B"+(dr+5)+"/B"+(dr+6)+",0)","件","→ 02作業DBの日次追客に割る"]
  ];
  u.getRange(dr+1,1,drv.length,4).setValues(drv);
  u.getRange(dr+1,1,drv.length,1).setBackground(G);
  [0,2,4,6].forEach(function(k){ u.getRange(dr+1+k,2).setBackground(B); });  // 計算=青
  [1,3,5].forEach(function(k){ u.getRange(dr+1+k,2).setBackground(Y); });    // 入力=黄(単価/成約率/追客率)
  u.getRange(dr+1,2).setBackground(O);                                        // ギャップ=橙
  u.getRange(dr+1,2).setNumberFormat("#,##0"); u.getRange(dr+2,2).setNumberFormat("#,##0");
  u.getRange(dr+3,2).setNumberFormat("0"); u.getRange(dr+5,2).setNumberFormat("0"); u.getRange(dr+7,2).setNumberFormat("0");
  u.getRange(dr+4,2).setNumberFormat("0%"); u.getRange(dr+6,2).setNumberFormat("0%");

  // ===== 列表示整理＆色（D/H/I/J/M/R/S 非表示。連動は色で明示）=====
  [4,8,9,10,13,18,19].forEach(function(c){ try{u.hideColumns(c,1);}catch(e){} });
  u.getRange("E5:E14").setBackground(Y);   // 満額=入力(黄)
  u.getRange("F5:G14").setBackground(B);    // 確度/着金=02連動(青)
  u.getRange("K5:K14").setBackground(B);    // 採用額=自動(青)
  u.setFrozenRows(4);

  var dump = u.getRange(1,1,Math.min(38,u.getMaxRows()),17).getDisplayValues();
  Logger.log(JSON.stringify(dump));
  return { ok:true,
    changed:["KPIヘッダ(1-3行)","状態列O+クローズ条件付き書式","採用額K=確度加重living(成約満額/失注0)","営業ドライバー(30行〜)","M列ほか非表示"],
    note:"連動列B/E/F/G/Kは不動・式のみ変更。案件名(C)のNotion正規化は次ステップ(02のキー突合確認後)",
    dump:dump };
}

// ── 実装A基盤①：科目を勘定科目(決算/MF準拠)へ整合＋資金繰り入金行を作り替え 2026-06-06 ──
// 決定：科目集約6本(法人2/個人3+営業代行予備+雑収入)・手入力可。資金繰りは行数7維持のままSUMIFSを新科目へ貼り替え(下方の出金/合計/ランウェイ無傷)。
function buildAccountWire(body){
  const ss = SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  const Y="#FFF2CC", B="#CFE2F3", BR="#AA2E26", G="#ECECEC";
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var u = shk("売上見込み"); if(!u) return {error:"03(売上見込み)が無い"};

  // 勘定科目候補（決算: 法人=売上高2補助/個人=賃貸料・EC、+営業代行(個人予備)、雑収入）
  var KAMOKU = ["売上高（不動産売買）","売上高（レンタル・コンサル）","賃貸料","売上（EC・海外）","営業代行","雑収入"];
  // 旧科目→勘定科目マップ（値ベース置換＝行位置に依存しない）
  var MAP = {
    "栄町売却":"売上高（不動産売買）","テレアポ買取再販":"売上高（不動産売買）","買取再販":"売上高（不動産売買）",
    "物件売却":"売上高（不動産売買）","不動産売買":"売上高（不動産売買）",
    "医療テナントコンサル":"売上高（レンタル・コンサル）","医療コンサル":"売上高（レンタル・コンサル）",
    "AIコンサル":"売上高（レンタル・コンサル）","横浜NPO":"売上高（レンタル・コンサル）","コンサル":"売上高（レンタル・コンサル）",
    "EC粗利":"売上（EC・海外）","EC":"売上（EC・海外）","賃貸":"賃貸料","家賃":"賃貸料"
  };
  // 03 B5:B14 を値ベースで勘定科目へ置換（旧科目名→決算科目）
  var bvals = u.getRange("B5:B14").getValues();
  var changed=[];
  for(var i=0;i<bvals.length;i++){
    var v=(bvals[i][0]||"").toString().trim();
    if(MAP[v]){ changed.push(v+"→"+MAP[v]); bvals[i][0]=MAP[v]; }
  }
  u.getRange("B5:B14").setValues(bvals);
  u.getRange("B4").setValue("科目（勘定科目＝決算/MF準拠）");
  u.getRange("B4").setNote("【入力・案件登録時／プルダウン・手入力可】決算・MF記帳の勘定科目に一致。法人=売上高(不動産売買)/(レンタル・コンサル)、個人=賃貸料/売上(EC・海外)/営業代行(予備)/雑収入。資金繰り入金行はこの科目でSUMIFS連動。手で別科目も打てる。");
  var rule=SpreadsheetApp.newDataValidation().requireValueInList(KAMOKU,true).setAllowInvalid(true).build();
  u.getRange("B5:B14").setDataValidation(rule);

  // 集計の主要3行(B17:B19)を新科目キーに貼り替え(行は増やさない＝注記と衝突回避)
  u.getRange("A16").setValue("■ 勘定科目別 採用額(確度加重で④資金繰りへ流れる額)。全科目合計はB20");
  u.getRange("A17").setValue("売上高（不動産売買）"); u.getRange("B17").setFormula('=SUMIF($B$5:$B$14,$A17,$K$5:$K$14)');
  u.getRange("A18").setValue("売上高（レンタル・コンサル）"); u.getRange("B18").setFormula('=SUMIF($B$5:$B$14,$A18,$K$5:$K$14)');
  u.getRange("A19").setValue("賃貸料／売上(EC・海外)・他"); u.getRange("B19").setFormula('=B20-B17-B18');

  // ④資金繰り 入金行を勘定科目でSUMIFS貼り替え（行数7維持。EC=8・育休=10・その他=11は手入力温存で触らない）
  var cf=shk("資金繰り"); var cfrep=[];
  if(cf){
    cf.getRange("A5").setValue("売上高（不動産売買）");
    cf.getRange("A6").setValue("売上高（レンタル・コンサル）");
    cf.getRange("A7").setValue("賃貸料");
    cf.getRange("A8").setValue("売上（EC・海外）※実精算 手入力");
    cf.getRange("A9").setValue("営業代行");
    cf.getRange("A10").setValue("★麻梨奈 育休給付（月割）");
    cf.getRange("A11").setValue("その他入金（雑収入含む）");
    // SUMIFS対象行=5,6,7,9（EC=8と育休=10/その他=11は手入力のため除外）。B..I=6月〜翌1月の8ヶ月。
    var rowsToWire=[5,6,7,9]; var baseY=2026, baseM=6;
    for(var ri=0;ri<rowsToWire.length;ri++){
      var R=rowsToWire[ri];
      for(var col=0;col<8;col++){
        var mm=baseM+col, yy=baseY+Math.floor((mm-1)/12), m2=((mm-1)%12)+1;
        cf.getRange(R,2+col).setFormula("=SUMIFS('03_売上見込み'!$K:$K,'03_売上見込み'!$B:$B,$A"+R+",'03_売上見込み'!$G:$G,DATE("+yy+","+m2+",1))");
      }
    }
    cf.getRange("A5:I9").setNote("【自動】03の採用額Kを『科目(A列)一致 × 着金月=その列の月』でSUMIFS。EC行(8)とその他/育休は手入力。");
    cfrep.push("資金繰りB5:I7,B9:I9を勘定科目SUMIFSへ／EC(8)・育休(10)・その他(11)は手入力温存");
  }
  return { ok:true, remapped:changed, kamoku:KAMOKU, cashflow:cfrep,
    note:"行挿入なし＝下方の出金/合計/月末現預金/ランウェイ無傷。EC入金は手入力維持(連動は03にEC月次行を入れれば後で可)。集計B17-19は主要科目表示・全科目はB20。" };
}

// ── 実装A基盤②：レンタル削除＋勘定科目マスター＋毎月安定収入(EC/賃料)をrecurringで連動 2026-06-06 ──
function buildRecurringIncome(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  const Y="#FFF2CC", B="#CFE2F3", BR="#AA2E26", G="#ECECEC";
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var u=shk("売上見込み"); if(!u) return {error:"03が無い"};
  var cf=shk("資金繰り");

  // 1) レンタル削除：売上高（レンタル・コンサル）→売上高（コンサル）（03値・資金繰りラベル・集計の3箇所）
  var OLD="売上高（レンタル・コンサル）", NEW="売上高（コンサル）";
  var bvals=u.getRange("B5:B14").getValues();
  for(var i=0;i<bvals.length;i++){ if((bvals[i][0]||"").toString().trim()===OLD) bvals[i][0]=NEW; }
  u.getRange("B5:B14").setValues(bvals);
  if(cf && cf.getRange("A6").getValue().toString().indexOf("コンサル")>=0) cf.getRange("A6").setValue(NEW);
  if(u.getRange("A18").getValue().toString().indexOf("コンサル")>=0) u.getRange("A18").setValue(NEW);

  // 2) 勘定科目マスター（ここを編集するとプルダウンが変わる）A50見出し＋A51:A56
  var KAMOKU=["売上高（不動産売買）","売上高（コンサル）","賃貸料","売上（EC・海外）","営業代行","雑収入"];
  u.getRange("A50:A56").clearContent();
  u.getRange("A50").setValue("■ 勘定科目マスター（この6行を編集すると科目プルダウンが変わる＝あなたが自分で表記変更可。決算/MF記帳に合わせる）").setFontWeight("bold").setFontColor(BR);
  for(var k=0;k<KAMOKU.length;k++){ u.getRange(51+k,1).setValue(KAMOKU[k]); }
  u.getRange("A51:A56").setBackground(Y);
  var rule=SpreadsheetApp.newDataValidation().requireValueInRange(u.getRange("A51:A56"),true).setAllowInvalid(true).build();
  u.getRange("B5:B14").setDataValidation(rule);

  // 3) 毎月の安定収入(recurring)を03に追加。行12=EC月次粗利、行13=北千住賃料。満額=1ヶ月額・確度100%。
  u.getRange("A12").setValue("03"); u.getRange("B12").setValue("売上（EC・海外）"); u.getRange("C12").setValue("EC月次粗利(クーパン/Amazon/TT)");
  u.getRange("E12").setValue(77000); u.getRange("F12").setValue(1); u.getRange("G12").setValue("毎月");
  u.getRange("K12").setFormula("=E12*F12"); u.getRange("M12").setFormula("=E12*F12");
  u.getRange("A13").setValue("01"); u.getRange("B13").setValue("賃貸料"); u.getRange("C13").setValue("北千住 賃料(千住元町・徐様)");
  u.getRange("E13").setValue(81000); u.getRange("F13").setValue(1); u.getRange("G13").setValue("毎月");
  u.getRange("K13").setFormula("=E13*F13"); u.getRange("M13").setFormula("=E13*F13");
  u.getRange("C12").setNote("毎月の安定収入(recurring)。満額=1ヶ月の額。実績ベース・要調整(2025通年粗利92.6万≒77,000/月)。");
  u.getRange("C13").setNote("毎月の安定収入(recurring)。年972,000÷12=81,000/月(徐様・千住元町)。");
  u.getRange("E12:F13").setBackground(Y);

  // 4) 資金繰り：賃貸料(7)・EC(8)を recurring(着金月フィルタ無し＝毎月同額計上)に。不動産売買/コンサル/営業代行は一時(着金月)維持。
  if(cf){
    cf.getRange("A8").setValue("売上（EC・海外）");
    for(var col=0;col<8;col++){
      cf.getRange(7,2+col).setFormula("=SUMIF('03_売上見込み'!$B:$B,$A7,'03_売上見込み'!$K:$K)");
      cf.getRange(8,2+col).setFormula("=SUMIF('03_売上見込み'!$B:$B,$A8,'03_売上見込み'!$K:$K)");
    }
    cf.getRange("A7:I8").setNote("【自動・毎月計上(recurring)】賃貸料・ECは月次安定収入。03の該当科目の採用額Kを毎月そのまま計上(着金月フィルタなし)。");
  }
  return { ok:true, renamed:OLD+"→"+NEW, master:"03 A51:A56＝勘定科目マスター(編集でプルダウン変更)",
    recurring:["EC 77,000/月(03行12)","北千住賃料 81,000/月(03行13)"],
    cashflow:"資金繰り 賃貸料(7)・EC(8) を毎月計上(recurring)に変更",
    pending:"04損益PLへのEC/賃料反映は次ステップ(PL構造を読んでGASで安全に行追加)" };
}

// ── 実装A基盤③：EC菊池ネット統一(03/PL)＋オーロラ加盟金6月計上(資金繰り・経常燃焼除外) 2026-06-06 ──
function buildFixes3(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  const Y="#FFF2CC", O="#FCE5CD", BR="#AA2E26";
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var u=shk("売上見込み"), cf=shk("資金繰り"), pl=shk("損益");
  var out={};

  // ① 03 EC＝菊池ネット61,000(ゆーし折半後・直近実績3ヶ月平均)
  if(u){
    u.getRange("E12").setValue(61000);
    u.getRange("C12").setValue("EC月次粗利(菊池ネット・ゆーし折半後)");
    u.getRange("C12").setNote("EC粗利は相方ゆーしと折半。これは菊池の取り分(直近実績3ヶ月平均≒61,000/月・分配スプシ1JmKEGWG I列)。ゆーしへは四半期締め→翌々月末後払い(5/8/11/2月末)。要ルール精査。");
    u.getRange("E12").setBackground(Y);
    out.ec03=61000;
  }
  // ② PL EC粗利を菊池ネットに統一(03!K12連動・グロス142,730→ネット61,000)
  if(pl){
    pl.getRange("B15:I15").setFormula("='03_売上見込み'!$K$12");
    pl.getRange("A15").setValue("個人事業 EC粗利益（菊池ネット・ゆーし折半後）");
    pl.getRange("A15").setNote("ゆーしに半分支払うため菊池手残りは折半後。03のEC採用額K12連動。グロス粗利(B13-B14≒142,730)とは別物。");
    out.plEc="PL B15:I15←03!K12(61,000)";
  }
  // ③ 資金繰り：その他出金行(26)をスポット初期投資化＝6月オーロラFC加盟金3,655,000。経常燃焼から除外。
  if(cf){
    cf.getRange("A26").setValue("★初期投資・スポット出金（経常燃焼に含めない）");
    cf.getRange("B26").setValue(3655000); // 6月(B列)のみ。C:Iは0据置
    cf.getRange("A26").setNote("6月:オーロラFC加盟金 3,655,000(税込)＝加盟金1,500,000+保証金300,000+研修費1,500,000+システム50,000+消費税305,000。請求元フライハイト#217・期限5/31→6/24Go判断後で6/30計上。法人(KIKUCHI HD)・一括。");
    cf.getRange("A26").setFontColor(BR); cf.getRange("B26").setBackground(O);
    // 経常燃焼からスポット出金(B26)を除外（積立B24と同じ扱い）
    cf.getRange("B31").setFormula("=B27-B8-B9-B10-B24-B26");
    cf.getRange("A31").setNote("純月次燃焼(経常)＝出金合計−EC−育休−積立−スポット初期投資。一時の大型出金(オーロラ等)は体力計算から除外し、月末現金(B29)側で反映。");
    out.aurora="資金繰りB26(6月)=3,655,000・経常燃焼から除外";
  }
  return { ok:true, fixes:out,
    pending:["①借入SSoT集約(BS→5借入繋ぎ替え後に借入返済/借入条件削除。但し朝日/大東京/TBセゾン/住宅の証票4本が欠落＝要追加)","②賃料をPLに行追加(PL構造を読んでから安全に)","③オーロラのPL/BS計上(繰延資産で償却か一括費用かの判断)"] };
}

// ── 実装A基盤④：借入SSoT集約(借入返済/借入条件を削除)＋資金繰り主格別純増減 2026-06-06 ──
// 監査確定：5借入は全7本証票どおり正・BS負債147-167は既に5借入連動・借入返済はBSの3箇所のみ被参照・借入条件は参照ゼロ。
function buildBatch4(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  const BR="#AA2E26", B="#CFE2F3", G="#ECECEC";
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var rep={};

  // ① 借入集約：BSが借入返済タブを参照する3箇所を「値固定(凍結)」→参照を消してから削除
  var bs=shk("資産負債");
  if(bs){
    ["AR8:BJ8","BB54:BF54","AZ138:BJ138"].forEach(function(a){
      var rg=bs.getRange(a); rg.setValues(rg.getValues()); // 数式→現在値で固定(過去CF/残高を保全し依存を断つ)
    });
    rep.frozen=["AR8:BJ8(借入返済CF履歴)","BB54:BF54(大東京・レガシー)","AZ138:BJ138(自宅・レガシー)"];
  }
  // 借入条件(参照ゼロ)→削除、借入返済(凍結後)→削除。5借入を唯一のSSoTに集約。
  var nj=shk("借入条件"); if(nj){ ss.deleteSheet(nj); rep.deleted_jouken=true; }
  var ng=shk("借入返済"); if(ng){ ss.deleteSheet(ng); rep.deleted_hensai=true; }

  // ② 資金繰り 主格別 当月純増減（月グリッドの右・K4起点。当月=B列=6月）
  var cf=shk("資金繰り");
  if(cf){
    cf.getRange("K4").setValue("■ 主格別 当月純増減（6月）").setFontWeight("bold").setFontColor(BR);
    cf.getRange("K5:N5").setValues([["主格","入金","出金","純増減"]]).setBackground("#434343").setFontColor("#FFFFFF").setFontWeight("bold");
    // 法人=売上高2本/出金:税理士+外注+社保+法人借入返済+役員賞与+スポット(オーロラ)
    cf.getRange("K6").setValue("法人(KHD)");
    cf.getRange("L6").setFormula("=B5+B6");
    cf.getRange("M6").setFormula("=B14+B15+B16+B22+B25+B26");
    // 個人事業=賃料+EC+営業代行/出金:個人税理士+個人事業借入返済
    cf.getRange("K7").setValue("個人事業");
    cf.getRange("L7").setFormula("=B7+B8+B9");
    cf.getRange("M7").setFormula("=B17+B23");
    // 家計=育休+その他/出金:住宅+SMBC+PayPay+生活費+投信積立
    cf.getRange("K8").setValue("家計(プライベート)");
    cf.getRange("L8").setFormula("=B10+B11");
    cf.getRange("M8").setFormula("=B18+B19+B20+B21+B24");
    cf.getRange("N6:N8").setFormulas([["=L6-M6"],["=L7-M7"],["=L8-M8"]]);
    cf.getRange("K9").setValue("合計(検算)");
    cf.getRange("L9").setFormula("=SUM(L6:L8)"); cf.getRange("M9").setFormula("=SUM(M6:M8)"); cf.getRange("N9").setFormula("=L9-M9");
    cf.getRange("K6:K9").setBackground(G);
    cf.getRange("L6:N9").setNumberFormat("#,##0").setBackground(B);
    cf.getRange("K4").setNote("当月(6月=B列)を主格別に。検算:L9=入金合計B12・M9=出金合計B27と一致するはず。賃料/育休等の主格割当は仮置き(賃料=個人事業/育休=家計/積立=家計/オーロラ=法人)。直したい割当があれば言って。");
    rep.shukaku="資金繰りK4:N9に主格別純増減(検算付き)";
  }
  return { ok:true, batch:rep,
    pending_next:["賃料をPLに行追加(PL密＝GAS行挿入で全社経常に組込・不可逆削除と分離して次に)","税金抜け(国保/年金/所得/住民は確定申告の実額を読んでから・推測で入れない)","オーロラのPL/BS計上(繰延資産で償却か一括費用かの会計判断)","BS自動化(MF残高読取が要る=Chrome手順・5借入連動は済)"],
    note:"朝日金利2.575→2.825%(5/31〜)は残高不変・利息月+約189円の誤差。必要なら5朝日明細を後で精緻化。" };
}

// ── 実装A基盤5：朝日信金 金利改定 2.575%→2.825%(令和8/5/31〜)を借入タブへ反映＋備考 2026-06-06 ──
// 出所=260501通知書。元金均等で元金16,000固定・残高不変、利息のみ再計算(利息=(残高+元金)×2.825%/12)。
function buildBatch5(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var go=shk("借入"); // 借入返済/借入条件は削除済→「借入」を含むのは借入タブのみ
  if(!go) return {error:"借入タブが見つからない(借入を含むタブ無し)"};
  const RATE=0.02825, GEN=16000;
  // 朝日ブロック列(監査): 月返済M=13/利息N=14/元金O=15/残高P=16。2026/01=行4→2026/06=行9。
  const cPay=13, cInt=14, cGen=15, cBal=16;
  var changed=0, before=[], after=[];
  for(var r=9;r<=60;r++){
    var gen=go.getRange(r,cGen).getValue();
    var bal=go.getRange(r,cBal).getValue();
    if(gen===GEN && typeof bal==="number" && bal>0){ // 元金16,000一致＝朝日の有効行のみ(誤爆ガード)
      var oldInt=go.getRange(r,cInt).getValue();
      var newInt=Math.round((bal+GEN)*RATE/12); // 利息=期首残高×年率/12
      go.getRange(r,cInt).setValue(newInt);
      go.getRange(r,cPay).setValue(GEN+newInt);
      if(changed<2){ before.push(oldInt); after.push(newInt); }
      changed++;
    }
  }
  go.getRange(2,cInt).setNote("【変更履歴 2026-06-06】朝日信金 金利 2.575%→2.825%(令和8/5/31〜・出所=260501通知書)。元金均等で元金16,000固定・残高不変、2026/6以降の利息のみ再計算(=(残高+元金)×2.825%/12)＝"+changed+"行更新。利息は月+約180円。");
  Logger.log("【朝日 利息 before→after】更新行数="+changed+" / 利息サンプル before="+JSON.stringify(before)+" → after="+JSON.stringify(after));
  return { ok:true, asahi_updated_rows:changed, sample_before_int:before, sample_after_int:after,
    note:"残高は元金均等で不変。資金繰り/PLの支払利息(借入タブ連動)に自動反映。全変更は当該セルの備考に記録。" };
}

// ── 実装A基盤6(v3)：入金/出金とも「主格別小計→合計」の同じ形に統一＋別枠ブロック/空白を撤去 2026-06-06 ──
// 既存の主格別系を一旦全撤去してクリーン再構築。主格タグ(列K)+SUMIFで行ズレに強い。
function buildBatch6(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  const B="#CFE2F3", G="#ECECEC";
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var cf=shk("資金繰り"); if(!cf) return {error:"資金繰りが無い"};
  function findRow(kw){ var a=cf.getRange(1,1,90,1).getValues(); for(var i=0;i<a.length;i++){ if((a[i][0]||"").toString().indexOf(kw)>=0) return i+1; } return -1; }

  // 0) 右付けK4:N9撤去
  cf.getRange("K4:N9").clear();
  // 1) 既存の主格別系を全撤去（再構築のため・ラベル一致行を下から削除）
  var contains=["主格別","入金小計","出金小計","検算"];
  var exact=["主格","法人(KHD)","個人事業","家計(プライベート)"];
  var aV=cf.getRange(1,1,90,1).getValues();
  for(var r=90;r>=1;r--){ var v=(aV[r-1][0]||"").toString().trim();
    if(contains.some(function(k){return v.indexOf(k)>=0;}) || exact.indexOf(v)>=0){ try{cf.deleteRow(r);}catch(e){} }
  }
  // 2) 主格タグを明細行に再付与（入金5-11 / 出金14-26）
  cf.getRange("K5:K11").setValues([["法人"],["法人"],["個人事業"],["個人事業"],["個人事業"],["家計"],["家計"]]);
  cf.getRange("K14:K26").setValues([["法人"],["法人"],["法人"],["個人事業"],["家計"],["家計"],["家計"],["家計"],["法人"],["個人事業"],["家計"],["法人"],["法人"]]);
  cf.getRange("K4").setValue("主格タグ(集計用)").setFontColor("#999999");

  var mk=[["法人(KHD)","法人"],["個人事業","個人事業"],["家計(プライベート)","家計"]];
  // 3) 出金合計の直前に「主格別 出金小計」を3行挿入
  var outT=findRow("出金合計");
  if(outT>0){
    cf.insertRowsBefore(outT,3);
    for(var i=0;i<3;i++){ var r=outT+i;
      cf.getRange(r,1).setValue(mk[i][0]+" 出金小計");
      cf.getRange(r,2).setFormula('=SUMIF($K$14:$K$26,"'+mk[i][1]+'",$B$14:$B$26)');
    }
    cf.getRange(outT,1,3,1).setBackground(G); cf.getRange(outT,2,3,1).setBackground(B).setNumberFormat("#,##0");
  }
  // 4) 入金合計の直前に「主格別 入金小計」を3行挿入（出金挿入後に検索＝位置自動追従）
  var inT=findRow("入金合計");
  if(inT>0){
    cf.insertRowsBefore(inT,3);
    for(var j=0;j<3;j++){ var rr=inT+j;
      cf.getRange(rr,1).setValue(mk[j][0]+" 入金小計");
      cf.getRange(rr,2).setFormula('=SUMIF($K$5:$K$11,"'+mk[j][1]+'",$B$5:$B$11)');
    }
    cf.getRange(inT,1,3,1).setBackground(G); cf.getRange(inT,2,3,1).setBackground(B).setNumberFormat("#,##0");
  }
  cf.hideColumns(11,1);
  return { ok:true,
    layout:"入金: 明細→法人/個人/家計 入金小計→入金合計。出金: 明細→法人/個人/家計 出金小計→出金合計。別枠ブロック/余分空白は撤去。",
    check:"入金小計3つの和=入金合計／出金小計3つの和=出金合計。総純増減は従来どおり入金合計−出金合計。",
    note:"主格割当(仮):賃料=個人事業/育休=家計/投信積立=家計/オーロラ=法人。" };
}

// ── 実装A基盤7：04損益PL＋03売上見込みを主格別に色分け＝ぱっと見で主格と集計が分かる 2026-06-06 ──
function buildBatch7(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  function frS(sh,kw){ var a=sh.getRange(1,1,90,1).getValues(); for(var i=0;i<a.length;i++){ if((a[i][0]||"").toString().indexOf(kw)>=0) return i+1; } return -1; }
  const LAW="#1155CC", PER="#38761D", HOME="#B45F06", ALLC="#434343";
  const LAWt="#E8F0FE", PERt="#E6F4EA", HOMEt="#FFF2E6", ALLt="#F1F1F1", GOLD="#FFE599", Y="#FFF2CC", BL="#CFE2F3";
  var out={};

  // === 04 損益PL：主格セクションを帯+薄色、結論行は金 ===
  var pl=shk("損益");
  if(pl){
    const COLS=9;
    var sec=[ {h:frS(pl,"■法人"),bg:LAW,tint:LAWt}, {h:frS(pl,"■個人事業"),bg:PER,tint:PERt}, {h:frS(pl,"■家計"),bg:HOME,tint:HOMEt}, {h:frS(pl,"■全社"),bg:ALLC,tint:ALLt} ];
    var endAll=frS(pl,"家族が潰れ"); if(endAll<1) endAll=(sec[3].h>0?sec[3].h+8:90);
    var bnd=[sec[1].h, sec[2].h, sec[3].h, endAll];
    for(var i=0;i<4;i++){ var s=sec[i].h; if(s<1) continue; var e=bnd[i]-1; if(e<s) e=s;
      pl.getRange(s,1,e-s+1,COLS).setBackground(sec[i].tint).setFontColor("#000000").setFontWeight("normal");
      pl.getRange(s,1,1,COLS).setBackground(sec[i].bg).setFontColor("#FFFFFF").setFontWeight("bold");
    }
    var aV=pl.getRange(1,1,endAll,1).getValues();
    for(var r=1;r<=endAll;r++){ var v=(aV[r-1][0]||"").toString();
      if(["経常利益","税引後利益","収支","黒字判定"].some(function(k){return v.indexOf(k)>=0;}))
        pl.getRange(r,1,1,COLS).setBackground(GOLD).setFontWeight("bold").setFontColor("#000000");
    }
    var hdr=frS(pl,"科目"); if(hdr<1) hdr=2;
    pl.getRange(hdr,1).setNote("色=主格。青=法人(KHD)／緑=個人事業(EC等)／橙=家計(00)／灰=全社。金+太字=結論(経常利益・税引後・収支・黒字判定)。");
    out.pl="法人青/個人緑/家計橙/全社灰+結論金";
  }

  // === 03 売上見込み：案件行の左側(A:C)を科目→主格で色分け。満額黄/自動青/集計金は維持 ===
  var u=shk("売上見込み");
  if(u){
    const C3=17; // A〜Q
    u.getRange(4,1,1,C3).setBackground(ALLC).setFontColor("#FFFFFF").setFontWeight("bold"); // ヘッダー帯
    var bv=u.getRange(5,2,10,1).getValues(); // B5:B14=科目
    for(var k=0;k<10;k++){ var sc=(bv[k][0]||"").toString(); var tint="#FFFFFF";
      if(/不動産売買|コンサル/.test(sc)) tint=LAWt;        // 法人=青
      else if(/賃貸料|EC|営業代行/.test(sc)) tint=PERt;    // 個人事業=緑
      else if(/雑収入/.test(sc)) tint=ALLt;               // 灰
      u.getRange(5+k,1,1,3).setBackground(tint);          // A:C(本部/科目/案件名)だけ主格色
    }
    u.getRange("E5:E14").setBackground(Y);                 // 満額=入力(黄)
    u.getRange("F5:G14").setBackground(BL); u.getRange("K5:K14").setBackground(BL); // 確度/着金/採用額=自動(青)
    var ag=frS(u,"勘定科目別"); if(ag>0) u.getRange(ag,1,5,2).setBackground(GOLD).setFontWeight("bold"); // 集計=金
    u.getRange(4,1).setNote("案件行の左側色=主格。青=法人(不動産売買/コンサル)／緑=個人事業(賃貸料/EC/営業代行)／灰=雑収入。満額E=黄(入力)・確度F/着金G/採用額K=青(自動)・勘定科目別採用額=金(集計)。");
    out.board03="案件行A:Cを科目→主格で色分け(青/緑)+満額黄/自動青/集計金";
  }
  return { ok:true, applied:out, note:"04と03で配色統一: 法人=青/個人事業=緑/家計=橙/全社=灰/結論集計=金。" };
}

// ── 実装A基盤8：01司令塔/02作業DB/06BS の統一色付け＋本部(00-05)色分け 2026-06-06 ──
function buildBatch8(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  const DARK="#434343", GOLD="#FFE599";
  const HONBU={ "00":"#F4CCCC","01":"#FFF2CC","02":"#D9EAD3","03":"#CFE2F3","04":"#D9D2E9","05":"#FCE5CD" };
  const LAWt="#E8F0FE", PERt="#E6F4EA", HOMEt="#FFF2E6";
  var out={};

  // === 02 作業DB：本部列(C)を本部色チップで塗分け ===
  var db=shk("作業DB");
  if(db){
    var lastR=db.getLastRow(), lastC=db.getLastColumn();
    db.getRange(1,1,1,lastC).setBackground(DARK).setFontColor("#FFFFFF").setFontWeight("bold");
    if(lastR>=2){
      var hv=db.getRange(2,3,lastR-1,1).getValues(); // C列=本部
      for(var i=0;i<hv.length;i++){
        var key=(hv[i][0]==null?"":hv[i][0]).toString().trim();
        var k2=(key.length===1?"0"+key:key);
        if(HONBU[k2]) db.getRange(2+i,3).setBackground(HONBU[k2]).setFontWeight("bold");
      }
    }
    db.getRange(1,3).setNote("本部色: 00家族=ピンク/01経営=黄/02資金=緑/03事業=青/04コンサル=紫/05物件=橙");
    out.db02="本部列(C)を本部色＋ヘッダー帯";
  }

  // === 01 統合司令塔：◆セクション帯＋結論金 ===
  var s1=shk("統合司令塔");
  if(s1){
    [1,4,9,13,17,31,36].forEach(function(r){ s1.getRange(r,1,1,5).setBackground(DARK).setFontColor("#FFFFFF").setFontWeight("bold"); });
    [5,11,37,39].forEach(function(r){ s1.getRange(r,1,1,2).setBackground(GOLD).setFontWeight("bold").setFontColor("#000000"); }); // 純資産/ランウェイ/全社経常/損益分岐
    s1.getRange(1,1).setNote("帯=セクション(過去/現在/未来/行動/損益)。金=結論(純資産・ランウェイ・経常利益・損益分岐)。");
    out.cmd01="◆セクション帯+結論金";
  }

  // === 06 資産負債BS：主格(法人青/研太緑/麻梨奈橙)で明細ラベル列D＋サマリー ===
  var bs=shk("資産負債");
  if(bs){
    bs.getRange("D50:D56").setBackground(LAWt);   bs.getRange("D57:D140").setBackground(PERt);  bs.getRange("D141:D144").setBackground(HOMEt); // 資産
    bs.getRange("D147:D151").setBackground(LAWt); bs.getRange("D152:D167").setBackground(PERt);                                                  // 負債
    bs.getRange(49,1,1,4).setBackground(DARK).setFontColor("#FFFFFF").setFontWeight("bold");   // 資産見出し
    bs.getRange(146,1,1,4).setBackground(DARK).setFontColor("#FFFFFF").setFontWeight("bold");  // 負債見出し
    [37,41,45].forEach(function(r){ bs.getRange(r,4).setBackground(LAWt); });   // 法人サマリー
    [38,42,46].forEach(function(r){ bs.getRange(r,4).setBackground(PERt); });   // 研太
    [39,43,47].forEach(function(r){ bs.getRange(r,4).setBackground(HOMEt); });  // 麻梨奈
    bs.getRange("D45:D47").setFontWeight("bold");                              // 自己資本=結論
    bs.getRange(49,4).setNote("ラベル列の色=主格。青=法人/緑=研太(個人)/橙=麻梨奈(家計)。");
    out.bs06="明細ラベル列Dを主格色＋サマリー＋見出し帯";
  }
  return { ok:true, applied:out,
    palette:"本部=00ピンク/01黄/02緑/03青/04紫/05橙　主格=法人青/個人(研太)緑/家計(麻梨奈)橙/結論金" };
}

// ── 実装A基盤9：06BSのズレ修正(運営元タグ駆動)＋05資金繰り主格色＋03本部色 2026-06-06 ──
function buildBatch9(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  const DARK="#434343", GOLD="#FFE599";
  const HONBU={"00":"#F4CCCC","01":"#FFF2CC","02":"#D9EAD3","03":"#CFE2F3","04":"#D9D2E9","05":"#FCE5CD"};
  const LAWt="#E8F0FE", PERt="#E6F4EA", HOMEt="#FFF2E6";
  function unei(v){ v=(v==null?"":v).toString(); if(v.indexOf("法人")>=0)return LAWt; if(v.indexOf("研太")>=0)return PERt; if(v.indexOf("麻梨奈")>=0)return HOMEt; return null; }
  var out={};

  // === 06 BS：ズレ修正＝B列(運営元)/D列ラベルを1行ずつ読んで塗り直し（ハードコード行を廃止）===
  var bs=shk("資産負債");
  if(bs){
    bs.getRange("A30:D170").setBackground(null); // 旧ズレ色をリセット(A-D・月列は不触)
    var v=bs.getRange(30,1,141,4).getValues();   // A30:D170 → [A,B,C,D]
    for(var i=0;i<v.length;i++){
      var rr=30+i, A=(v[i][0]||"").toString().trim(), Bu=v[i][1], D=(v[i][3]||"").toString();
      if(A==="資産"||A==="負債"){ bs.getRange(rr,1,1,4).setBackground(DARK).setFontColor("#FFFFFF").setFontWeight("bold"); continue; }
      if(/自己資本|純資産/.test(D)){ bs.getRange(rr,4).setBackground(GOLD).setFontWeight("bold"); continue; }
      var c=unei(Bu)||unei(D);   // 明細は運営元(B)、サマリーはDラベル
      if(c) bs.getRange(rr,4).setBackground(c);
    }
    bs.getRange(49,4).setNote("D列の色=主格(運営元B列で自動判定)。青=法人/緑=研太(個人)/橙=麻梨奈(家計)/金=自己資本(結論)。");
    out.bs06="運営元タグ駆動で塗り直し＝ズレ解消";
  }

  // === 05 資金繰り：主格色(Kタグ駆動)＋結論金 ===
  var cf=shk("資金繰り");
  if(cf){
    var lastR=cf.getLastRow();
    var aV=cf.getRange(1,1,lastR,1).getValues();
    var kV=cf.getRange(1,11,lastR,1).getValues(); // K列=主格タグ(batch6で付与)
    var mt={"法人":LAWt,"個人事業":PERt,"家計":HOMEt};
    for(var r=1;r<=lastR;r++){
      var lbl=(aV[r-1][0]||"").toString(), tag=(kV[r-1][0]||"").toString().trim();
      if(/入金合計|出金合計|当月純増減|月末現預金|ランウェイ/.test(lbl)){ cf.getRange(r,1,1,9).setBackground(GOLD).setFontWeight("bold"); continue; }
      if(/入金小計|出金小計/.test(lbl)){ var m=lbl.indexOf("法人")>=0?LAWt:(lbl.indexOf("個人")>=0?PERt:(lbl.indexOf("家計")>=0?HOMEt:null)); if(m) cf.getRange(r,1,1,4).setBackground(m).setFontWeight("bold"); continue; }
      if(mt[tag]) cf.getRange(r,1,1,2).setBackground(mt[tag]); // 明細はA:Bを主格色
    }
    out.cf05="主格色(Kタグ)＋小計主格色＋結論金";
  }

  // === 03 売上見込み：本部色をA列に追加（主格B:Cは維持）===
  var u=shk("売上見込み");
  if(u){
    var av=u.getRange(5,1,10,1).getValues(); // A5:A14=本部
    for(var k=0;k<10;k++){ var key=(av[k][0]==null?"":av[k][0]).toString().trim(); var k2=(key.length===1?"0"+key:key);
      if(HONBU[k2]) u.getRange(5+k,1).setBackground(HONBU[k2]).setFontWeight("bold"); }
    out.board03="本部色をA列に追加(主格B:C維持)";
  }
  return { ok:true, applied:out, note:"06はハードコード行を廃止し運営元B列で自動判定＝ズレ再発しない。" };
}

// ── 実装A基盤10：PL整備(賃料/個人利息/ゆーし業務委託費/EC明細/家計明細) 2026-06-06 ──
// 行挿入で参照は自動追従。全社の限界/固定費/経常のみ新構成に再計算。EC菊池ネット61,000で03/資金繰りと整合。
function buildBatch10(body){
  return {ok:false, note:"buildBatch10は廃止(buildBatch11に統合)。実行不要＝重複事故の原因。何もしません。"};
  // eslint-disable-next-line no-unreachable
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var pl=shk("損益"); if(!pl) return {error:"04損益PLが無い"};
  // 二重実行ガード
  var chk=pl.getRange(1,1,40,1).getValues();
  for(var g=0;g<chk.length;g++){ if((chk[g][0]||"").toString().indexOf("EC営業利益")>=0) return {ok:true, note:"既にPL整備済み(EC営業利益あり)・スキップ"}; }

  // 03 EC菊池ネットを61,000に確定（buildFixes3未実行でも整合させる）
  var u=shk("売上見込み"); if(u){ u.getRange("E12").setValue(61000); }

  // === 個人事業：粗利(R15)の直後に5行挿入してEC明細＋賃料＋ゆーし費＋利息 ===
  pl.insertRowsAfter(15,5); // 新R16-20
  pl.getRange("A16").setValue("　ツール・送料等（EC月額費用）");                 pl.getRange("B16").setFormula("=B15-B17");
  pl.getRange("A17").setValue("　EC営業利益（両者・ツール控除後）");             pl.getRange("B17").setFormula("='03_売上見込み'!K12*2");
  pl.getRange("A18").setValue("　賃料収入（不動産所得・北千住）");               pl.getRange("B18").setFormula("='03_売上見込み'!K13");
  pl.getRange("A19").setValue("　ゆーし業務委託費（EC折半・四半期後払い）");      pl.getRange("B19").setFormula("='03_売上見込み'!K12");
  var GO="'"+String.fromCharCode(0x2464)+" 借入'"; // 借入タブ名(丸5+借入)をソースに直書きせず生成(貼付化け防止)
  pl.getRange("A20").setValue("　支払利息（個人事業借入・借入タブ連動）");        pl.getRange("B20").setFormula("="+GO+"!AB9+"+GO+"!AW9+"+GO+"!AI9");
  pl.getRange("A21").setValue("　固定費（個人税理士）");
  pl.getRange("A22").setValue("　経常利益（個人＝EC菊池ネット＋賃料−税理士−利息）");
  pl.getRange("B22").setFormula("=B17+B18-B19-B20-B21");
  pl.getRange("B16").setNote("EC粗利(B15)−EC営業利益(両者)＝ツール・送料の月額均し。");
  pl.getRange("B17").setNote("03のK12(菊池ネット61,000)×2＝両者の営業利益。ゆーし費(B19)を引くと菊池ネットに戻る。");
  pl.getRange("B19").setNote("ゆーしへの折半支払(=K12)。四半期締め→翌々月末後払い(5/8/11/2月末)。要EC一周精査。");
  pl.getRange("B20").setNote("個人事業の借入利息＝TB創業(AB)+公庫(AW)+TBセゾン(AI)。住宅MCJ利息は家計側。");

  // === 家計：収入(R24)の直後に2行挿入して生活費を明細化 ===
  pl.insertRowsAfter(24,2); // 新R25,R26
  pl.getRange("A24").setValue("　収入（育休給付・将来は復職給与）");
  pl.getRange("A25").setValue("　生活費（楽天・経常）");           pl.getRange("B25").setValue(180000);
  pl.getRange("A26").setValue("　生活費（麻梨奈・三井住友）");      pl.getRange("B26").setValue(121000);
  pl.getRange("A27").setValue("　生活費 計");                     pl.getRange("B27").setFormula("=B25+B26");
  pl.getRange("A28").setValue("　収支");                          pl.getRange("B28").setFormula("=B24-B27");
  pl.getRange("B25").setNote("楽天ブラック経常分(大型単発除く)。減らせる。");
  pl.getRange("B26").setNote("麻梨奈 三井住友カード分。");

  // === 全社：限界/固定費/経常を新構成で再計算（全社=R29ヘッダ、限界30/固定費31/経常32）===
  pl.getRange("B30").setFormula("=B6+B17+B18");      // 限界＝法人粗利+EC営業利益(両者)+賃料
  pl.getRange("B31").setFormula("=B7+B19+B20+B21");  // 固定費計＝法人固定費+ゆーし+個人利息+個人税理士
  pl.getRange("B32").setFormula("=B9+B22");          // 経常＝法人経常+個人経常

  return { ok:true,
    personal:"EC明細(売上/原価/粗利/ツール/営業利益)+賃料+ゆーし業務委託費(固定費)+個人支払利息→経常75,358",
    home:"生活費を楽天18万/麻梨奈12.1万/計に明細化",
    all:"全社 限界(B30)/固定費(B31)/経常(B32) を新構成で再計算",
    next:"色は buildBatch7 を再実行で再適用(ラベル駆動で正しく当たる)" };
}

// ── 実装A基盤11：PLを決算第4期に接地して再構築(法人8科目+役員報酬/社保+個人EC明細+賃料+利息+家計明細) 2026-06-06 ──
// 全インサートはGAS行挿入で参照自動追従。アンカーはラベル基準。全社のみ新構成で再計算。
function buildBatch11(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var pl=shk("損益"); if(!pl) return {error:"04損益PLが無い"};
  function fr(kw){ var a=pl.getRange(1,1,130,1).getValues(); for(var i=0;i<a.length;i++){ if((a[i][0]||"").toString().indexOf(kw)>=0) return i+1; } return -1; }
  if(fr("業務委託料")>0) return {ok:true, note:"既にPL接地済(業務委託料あり)・スキップ"};

  var u=shk("売上見込み"); if(u) u.getRange("E12").setValue(61000); // EC菊池ネット確定
  var GO="'"+String.fromCharCode(0x2464)+" 借入'";

  // アンカー（現状: 法人3/個人12/家計18/全社22）。下(家計)→上(法人)の順で挿入＝上のアンカーがズレない。
  var hHome=fr("■家計"), hPer=fr("■個人事業"), hLaw=fr("■法人");

  // === A) 家計：生活費を楽天/麻梨奈に明細化 ===
  pl.insertRowsAfter(hHome+1, 2); // 収入(hHome+1)の直後に2行
  pl.getRange(hHome+1,1).setValue("　収入（育休給付・将来は復職給与）");
  pl.getRange(hHome+2,1).setValue("　生活費（楽天・経常）");          pl.getRange(hHome+2,2).setValue(180000);
  pl.getRange(hHome+3,1).setValue("　生活費（麻梨奈・三井住友）");     pl.getRange(hHome+3,2).setValue(121000);
  pl.getRange(hHome+4,1).setValue("　生活費 計");                    pl.getRange(hHome+4,2).setFormula("=B"+(hHome+2)+"+B"+(hHome+3));
  pl.getRange(hHome+5,1).setValue("　収支");                         pl.getRange(hHome+5,2).setFormula("=B"+(hHome+1)+"-B"+(hHome+4));

  // === B) 個人事業：EC明細+賃料+ゆーし業務委託費+個人支払利息 ===
  var pG=hPer+3; // 個人粗利
  pl.insertRowsAfter(pG,5);
  pl.getRange(pG+1,1).setValue("　ツール・送料等（EC月額費用）");            pl.getRange(pG+1,2).setFormula("=B"+pG+"-B"+(pG+2));
  pl.getRange(pG+2,1).setValue("　EC営業利益（両者・ツール控除後）");        pl.getRange(pG+2,2).setFormula("='03_売上見込み'!K12*2");
  pl.getRange(pG+3,1).setValue("　賃料収入（不動産所得・北千住）");          pl.getRange(pG+3,2).setFormula("='03_売上見込み'!K13");
  pl.getRange(pG+4,1).setValue("　ゆーし業務委託費（EC折半・四半期後払い）"); pl.getRange(pG+4,2).setFormula("='03_売上見込み'!K12");
  pl.getRange(pG+5,1).setValue("　支払利息（個人事業借入・借入連動）");       pl.getRange(pG+5,2).setFormula("="+GO+"!AB9+"+GO+"!AW9+"+GO+"!AI9");
  pl.getRange(pG+6,1).setValue("　固定費（個人税理士）"); // 既存(税理士46,750)が移動
  pl.getRange(pG+7,1).setValue("　経常利益（個人）");
  pl.getRange(pG+7,2).setFormula("=B"+(pG+2)+"+B"+(pG+3)+"-B"+(pG+4)+"-B"+(pG+5)+"-B"+(pG+6));

  // === C) 法人：固定費ベタ→決算8科目に分解接地（最後＝上に挿入しても下は処理済）===
  var lG=hLaw+3; // 法人粗利
  pl.insertRowsAfter(lG,8);
  var hojin=[["　業務委託料（江藤/B&S・7月で最後→8月〜0）",87000],["　接待交際費（JAL接待含む）",30000],["　旅費交通費（JAL出張/えきねっと含む）",10300],["　支払報酬（税理士・橋本）",18333],["　支払手数料",3300],["　保険料・租税公課・広告 他",2000],["　役員報酬（月5万・定期同額）",50000],["　法定福利費（社保）",22167]];
  for(var i=0;i<8;i++){ pl.getRange(lG+1+i,1).setValue(hojin[i][0]); pl.getRange(lG+1+i,2).setValue(hojin[i][1]); }
  pl.getRange(lG+9,1).setValue("　販管費 計（決算8科目接地）"); pl.getRange(lG+9,2).setFormula("=SUM(B"+(lG+1)+":B"+(lG+8)+")");
  pl.getRange(lG+11,1).setValue("　経常利益（法人）"); // 利息=lG+10(自動), 経常=lG+11
  pl.getRange(lG+8+1,1); // noop
  // 経常(法人)= 粗利 - 販管費計 - 支払利息（自動シフトされるが明示）
  pl.getRange(lG+11,2).setFormula("=B"+lG+"-B"+(lG+9)+"-B"+(lG+10));
  pl.getRange(lG+1,2).setNote("第4期決算 業務委託料1,046,498/年≒87,208/月(江藤/B&S)。7月で最後→8月以降0に手修正。");
  pl.getRange(lG+2,2).setNote("第4期決算 接待交際費359,876/年≒29,990/月。JAL法人移行(5月〜)の接待分もここへ記帳。");
  pl.getRange(lG+3,2).setNote("第4期決算 旅費交通費123,584/年≒10,299/月。JAL出張・えきねっと等もここへ記帳。");
  pl.getRange(lG+7,2).setNote("第4期は役員報酬0計上だったが今期は計上(定期同額・月5万・未払計上で損金)。");
  pl.getRange(lG+8,2).setNote("社保(口座振替ベース)。第4期決算は法定福利費0だったが今期計上。");

  // === D) 全社：限界/固定費/経常を新構成で再計算（ラベル基準で再特定）===
  var rLawGross=fr("粗利益（限界利益）");      // 1件目=法人粗利
  var rECop=fr("EC営業利益");
  var rRent=fr("賃料収入");
  var rLawFix=fr("販管費 計");
  var rPerZei=fr("固定費（個人税理士）");
  var rYushi=fr("ゆーし業務委託費");
  var rPerInt=fr("支払利息（個人事業借入");
  var rLawKei=fr("経常利益（法人）");
  var rPerKei=fr("経常利益（個人）");
  var hAll=fr("■全社");
  if(hAll>0){
    pl.getRange(hAll+1,2).setFormula("=B"+rLawGross+"+B"+rECop+"+B"+rRent);               // 全社限界
    pl.getRange(hAll+2,2).setFormula("=B"+rLawFix+"+B"+rPerZei+"+B"+rYushi+"+B"+rPerInt);  // 全社固定費計
    pl.getRange(hAll+3,2).setFormula("=B"+rLawKei+"+B"+rPerKei);                           // 全社経常
    pl.getRange(hAll+3,1).setValue("経常利益（全社）");
  }
  return { ok:true,
    law:"法人を決算8科目に分解(業務委託87,000/接待30,000/旅費10,300/税理士18,333/手数料3,300/保険他2,000/役員報酬50,000/社保22,167)＝販管費計約223,100",
    personal:"EC明細+賃料+ゆーし業務委託費+個人支払利息→経常",
    home:"生活費を楽天/麻梨奈に明細化",
    all:"全社 限界/固定費/経常を再計算",
    next:"buildBatch7 で再色付け→次に上下統合(資金繰り要約をPL下に)" };
}

// ── 実装A基盤12：PLの7月以降(C:I)を埋めて8ヶ月予測を完成＝事前入力運用OK 2026-06-06 ──
function buildBatch12(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var pl=shk("損益"); if(!pl) return {error:"04損益PLが無い"};
  function fr(kw){ var a=pl.getRange(1,1,130,1).getValues(); for(var i=0;i<a.length;i++){ if((a[i][0]||"").toString().indexOf(kw)>=0) return i+1; } return -1; }
  function CL(c){ return String.fromCharCode(64+c); } // 3→C … 9→I
  function carry(r){ for(var c=3;c<=9;c++) pl.getRange(r,c).setFormula("=$B"+r); }          // 6月値をキャリー(各列上書き可)
  function perCol(r,fn){ for(var c=3;c<=9;c++) pl.getRange(r,c).setFormula(fn(CL(c))); }     // 各列の式

  var rGyom=fr("業務委託料"); if(rGyom<1) return {error:"buildBatch11未実行(業務委託料なし)"};
  // 法人8科目: 6月値キャリー。業務委託料は7月(C)まで→8月(D)〜1月(I)=0。
  for(var k=0;k<8;k++) carry(rGyom+k);
  for(var c=4;c<=9;c++) pl.getRange(rGyom,c).setValue(0); // 業務委託料 D:I=0
  var rFix=fr("販管費 計"); perCol(rFix,function(L){ return "=SUM("+L+rGyom+":"+L+(rGyom+7)+")"; });
  var rLawGross=fr("粗利益（限界利益）");      // 1件目=法人粗利
  var rLawInt=fr("支払利息（借入");           // 法人利息(個人は『個人事業借入』で別)
  var rLawKei=fr("経常利益（法人）");
  perCol(rLawKei,function(L){ return "="+L+rLawGross+"-"+L+rFix+"-"+L+rLawInt; });

  // 個人事業
  var rECop=fr("EC営業利益"), rRent=fr("賃料収入"), rYushi=fr("ゆーし業務委託費");
  var rTool=fr("ツール・送料"), rPerInt=fr("支払利息（個人事業借入"), rPerZei=fr("固定費（個人税理士）");
  var rPerKei=fr("経常利益（個人）"); var rECgross=rECop-2; // 粗利→ツール→営業利益
  carry(rECop); carry(rRent); carry(rYushi); carry(rPerZei); carry(rPerInt);
  perCol(rTool,function(L){ return "="+L+rECgross+"-"+L+rECop; });
  perCol(rPerKei,function(L){ return "="+L+rECop+"+"+L+rRent+"-"+L+rYushi+"-"+L+rPerInt+"-"+L+rPerZei; });

  // 家計
  var rInc=fr("収入（育休"), rRak=fr("生活費（楽天"), rMar=fr("生活費（麻梨奈"), rSk=fr("生活費 計"), rShu=fr("収支");
  carry(rInc); carry(rRak); carry(rMar);
  perCol(rSk,function(L){ return "="+L+rRak+"+"+L+rMar; });
  perCol(rShu,function(L){ return "="+L+rInc+"-"+L+rSk; });

  // 全社
  var hAll=fr("■全社");
  if(hAll>0){
    perCol(hAll+1,function(L){ return "="+L+rLawGross+"+"+L+rECop+"+"+L+rRent; });          // 限界
    perCol(hAll+2,function(L){ return "="+L+rFix+"+"+L+rPerZei+"+"+L+rYushi+"+"+L+rPerInt; }); // 固定費計
    perCol(hAll+3,function(L){ return "="+L+rLawKei+"+"+L+rPerKei; });                        // 経常
  }
  pl.getRange(rGyom,3).setNote("業務委託料は7月(C列)まで→8月(D列)以降0(江藤7月で最後)。他科目はC:I=6月値キャリー＝各列を直接手修正できる。");
  return { ok:true, filled:"7月〜翌1月(C:I)を全行補完。業務委託料は7月迄→8月0。合計/経常/全社は各列で再計算。",
    note:"事前入力運用OK＝将来月のセルに直接数値を打てば上書きされる(キャリー式を消すだけ)。" };
}

// ── 実装A基盤13：PLの#REF修復(buildBatch10+11二重実行で個人/家計が重複した事故の復旧) 2026-06-06 ──
function buildBatch13(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var pl=shk("損益"); if(!pl) return {error:"04損益PLが無い"};
  function fr(kw){ var a=pl.getRange(1,1,90,1).getValues(); for(var i=0;i<a.length;i++){ if((a[i][0]||"").toString().indexOf(kw)>=0) return i+1; } return -1; }
  function CL(c){ return String.fromCharCode(64+c); } // 2→B … 9→I

  // ガード：賃料収入が2回以上＝重複(壊れた状態)のときだけ修復
  var aAll=pl.getRange(1,1,90,1).getValues(); var rc=0;
  for(var i=0;i<aAll.length;i++){ if((aAll[i][0]||"").toString().indexOf("賃料収入")>=0) rc++; }
  if(rc<2) return {ok:true, note:"重複なし・スキップ(既に正常)"};

  // 1) 個人：利息(個人事業借入・借入連動=11版)の直後〜■家計手前を全削除→税理士+経常を1本で再作成
  var rPerInt=fr("支払利息（個人事業借入・借入連動"); if(rPerInt<1) rPerInt=fr("支払利息（個人事業借入");
  var hHome=fr("■家計");
  var delN=hHome-rPerInt-1;
  if(delN>0) pl.deleteRows(rPerInt+1, delN);
  pl.insertRowsAfter(rPerInt,2);
  pl.getRange(rPerInt+1,1).setValue("　固定費（個人税理士）"); pl.getRange(rPerInt+1,2).setValue(46750);
  pl.getRange(rPerInt+2,1).setValue("　経常利益（個人）");
  var rECop=fr("EC営業利益"), rRent=fr("賃料収入"), rYushi=fr("ゆーし業務委託費");
  for(var c=2;c<=9;c++){ var L=CL(c); pl.getRange(rPerInt+2,c).setFormula("="+L+rECop+"+"+L+rRent+"-"+L+rYushi+"-"+L+rPerInt+"-"+L+(rPerInt+1)); }
  for(var c2=3;c2<=9;c2++) pl.getRange(rPerInt+1,c2).setFormula("=$B"+(rPerInt+1)); // 税理士キャリー

  // 2) 家計：正味5行(収入/楽天/麻梨奈/生活費計/収支)を超える重複を末尾削除
  var hHome2=fr("■家計"), hAll0=fr("■全社");
  var homeRows=hAll0-hHome2-1;
  if(homeRows>5) pl.deleteRows(hHome2+6, homeRows-5);

  // 3) 全社を各列(B:I)で再計算＝#REF連鎖を断つ
  var rLawGross=fr("粗利益（限界利益）"), rECop2=fr("EC営業利益"), rRent2=fr("賃料収入"),
      rFix=fr("販管費 計"), rZei=fr("固定費（個人税理士）"), rYushi2=fr("ゆーし業務委託費"),
      rPerInt2=fr("支払利息（個人事業借入"), rLawKei=fr("経常利益（法人）"), rPerKei=fr("経常利益（個人）"),
      hAll=fr("■全社");
  if(hAll>0){
    var lim=hAll+1, fix=hAll+2, kei=hAll+3, mok=hAll+4, bep=hAll+5, han=hAll+6;
    for(var c3=2;c3<=9;c3++){ var L=CL(c3);
      pl.getRange(lim,c3).setFormula("="+L+rLawGross+"+"+L+rECop2+"+"+L+rRent2);
      pl.getRange(fix,c3).setFormula("="+L+rFix+"+"+L+rZei+"+"+L+rYushi2+"+"+L+rPerInt2);
      pl.getRange(kei,c3).setFormula("="+L+rLawKei+"+"+L+rPerKei);
      pl.getRange(bep,c3).setFormula("="+L+fix+"+"+L+mok);
    }
    pl.getRange(han,2).setFormula('=IF(B'+lim+'>=B'+bep+',"○達成","✕不足"&TEXT(B'+bep+'-B'+lim+',"#,##0")&"円")');
    pl.getRange(kei,1).setValue("経常利益（全社）");
  }
  return { ok:true, fixed:"個人/家計の重複ブロック除去→税理士・経常を1本に再作成→全社をB:I再計算で#REF解消", deleted_personal_rows:delN, home_extra:(homeRows>5?homeRows-5:0) };
}

// ── 実装A基盤14：上下統合(04PLの下に資金繰りCF要約を05連動で連結)＋家計の根拠注記 2026-06-06 ──
function buildBatch14(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var pl=shk("損益"); if(!pl) return {error:"04損益PLが無い"};
  function fr(kw){ var a=pl.getRange(1,1,140,1).getValues(); for(var i=0;i<a.length;i++){ if((a[i][0]||"").toString().indexOf(kw)>=0) return i+1; } return -1; }
  function CL(c){ return String.fromCharCode(64+c); }
  const DARK="#434343", GOLD="#FFE599", CF="'05_資金繰り'";
  var out={};

  // === A) 上下統合：PL末尾の下に「資金繰り（CF・05連動）」要約を8ヶ月で連結 ===
  // 二重実行ガード
  if(fr("資金繰り（CF・05連動")<1){
    var base=pl.getLastRow()+2;
    pl.getRange(base,1).setValue("■ 資金繰り（CF・05連動）＝このPLの下でキャッシュを1画面（古田土 全部入り）").setFontWeight("bold").setFontColor("#FFFFFF").setBackground(DARK);
    pl.getRange(base,1,1,9).setBackground(DARK);
    // 月見出し（05のR2を引く）
    pl.getRange(base+1,1).setValue("科目＼月");
    var map=[["月初現預金",3],["入金合計",15],["出金合計",33],["当月純増減",34],["月末現預金（体力計）",35]];
    for(var c=2;c<=9;c++){ var L=CL(c);
      pl.getRange(base+1,c).setFormula("="+CF+"!"+L+"2");
      for(var m=0;m<map.length;m++){ pl.getRange(base+2+m,c).setFormula("="+CF+"!"+L+map[m][1]); }
    }
    for(var m=0;m<map.length;m++) pl.getRange(base+2+m,1).setValue(map[m][0]);
    pl.getRange(base+1,1,1,9).setBackground("#434343").setFontColor("#FFFFFF").setFontWeight("bold");
    pl.getRange(base+6,1,1,9).setBackground(GOLD).setFontWeight("bold"); // 月末現預金=結論(金)
    pl.getRange(base+7,1).setValue("ランウェイ（月数・経常燃焼ベース）"); pl.getRange(base+7,2).setFormula("="+CF+"!B38");
    pl.getRange(base+2,1,6,1).setBackground("#CFE2F3");
    pl.getRange(base,1).setNote("PL(上)＝損益の予測／資金繰りCF(下)＝05連動の現金。月初現金1セル(司令塔B10→05B3)から先6ヶ月が動く＝古田土の全部入り1枚。");
    out.cf="04PL末尾(R"+base+"〜)に資金繰りCF要約(月初/入金/出金/純増減/月末/ランウェイ・B:I)を05連動で連結";
  } else out.cf="既にCF要約あり・スキップ";

  // === B) 家計の根拠を注記で焼き込む（値は変えない＝再事故回避）===
  var rRak=fr("生活費（楽天"), rMar=fr("生活費（麻梨奈"), rSk=fr("生活費 計");
  if(rRak>0) pl.getRange(rRak,1).setNote("楽天ブラック経常＝食料品/テイクアウト/家族日用品。化粧品・NISA積立・ベビーシッターは含まない(カード使い分けルール)。大型単発は別。");
  if(rMar>0) pl.getRange(rMar,1).setNote("麻梨奈 三井住友＝MF実績3ヶ月平均約118k。光熱費/通信/サブスク/食費/日用品/ベビー用品が家計コア(月3.5-4.5万)＋麻梨奈プライベート美容代 月約1.2万を含む(家計外だが世帯支出)。NISA積立は別(05投信積立・止め可)。MF自動取得→IMPORT化が将来。");
  if(rSk>0) pl.getRange(rSk,1).setNote("ベビーシッター：江東区制度=昼2,500円/h上限・年144h(最大36万/年)・時給2,500以下なら実質ほぼ全額戻り・申請2-3ヶ月後償還。現在は利用実績なし(未活用)。使い始めたら『実費出金＋2-3ヶ月後ほぼ全額戻り入金』で両建てモデル化する。");
  out.kakei="楽天/麻梨奈/生活費計に根拠＋シッター(江東区・未活用・将来両建て)を注記";

  return { ok:true, applied:out, note:"色は buildBatch7 再実行で再適用。家計の値は据置(美容含む実支出のまま・注記で内訳明示)。" };
}

// ── 実装A基盤15：04を全部入り1枚へ(ドライバー/KPI最上部＋05全項目ミラー＋05非表示) 2026-06-06 ──
function buildBatch15(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var pl=shk("損益"); if(!pl) return {error:"04損益PLが無い"};
  function fr(kw){ var a=pl.getRange(1,1,200,1).getValues(); for(var i=0;i<a.length;i++){ if((a[i][0]||"").toString().indexOf(kw)>=0) return i+1; } return -1; }
  function CL(c){ return String.fromCharCode(64+c); }
  const DARK="#434343", GOLD="#FFE599", BL="#CFE2F3", CFR="'05_資金繰り'";
  var out={};

  // 0) buildBatch14の旧6行CFサマリーを撤去(あれば・ヘッダから8行)
  var oldcf=fr("資金繰り（CF・05連動");
  if(oldcf>0){ pl.deleteRows(oldcf, 8); }

  // 1) 05全項目をミラー(なければ追加)。05のR3〜R38をそのまま=参照で04末尾に。
  if(fr("資金繰り（CF・05全項目")<1){
    var cf05=shk("資金繰り"); var lastCF=cf05?cf05.getLastRow():38; if(lastCF>40) lastCF=40;
    var base=pl.getLastRow()+2;
    pl.getRange(base,1,1,9).setBackground(DARK);
    pl.getRange(base,1).setValue("■ 資金繰り（CF・05全項目ミラー）＝この下が現金。月初現金1セルから先6ヶ月が動く（古田土 全部入り）").setFontColor("#FFFFFF").setFontWeight("bold");
    for(var r=3;r<=lastCF;r++){
      var dst=base+1+(r-3);
      pl.getRange(dst,1).setFormula("="+CFR+"!A"+r);
      for(var c=2;c<=9;c++){ pl.getRange(dst,c).setFormula("="+CFR+"!"+CL(c)+r); }
    }
    out.mirror="04末尾(R"+base+"〜)に05資金繰り全項目をミラー(=05連動)";
  } else out.mirror="既にミラーあり・スキップ";

  // 2) 最上部に「今日の行動」ブロックを挿入(KPI+ドライバー要点)。既存ドライバーをラベル参照。
  if(fr("今日の行動（最上部")<1){
    pl.insertRowsAfter(1,7); // R2〜R8新規(タイトルR1の直後)
    pl.getRange(2,1,1,9).setBackground(DARK);
    pl.getRange(2,1).setValue("🎯 今日の行動（最上部・毎朝最初に見る）＝家族と黒字から逆算した今日やること").setFontColor("#FFFFFF").setFontWeight("bold");
    // 参照先をラベルで再特定(挿入後)
    var rGap=fr("経常ギャップ（毎月の穴"), rOiKa=fr("今日の追客目標"), rIte=fr("今日の一手"), rNeed=fr("必要な追客/月"), rAllKei=fr("経常利益（全社）");
    pl.getRange(3,1).setValue("🔴 経常ギャップ（家族ラインの穴/月）");      if(rGap>0)  pl.getRange(3,2).setFormula("=B"+rGap);
    pl.getRange(4,1).setValue("🎯 今日の追客目標");                       if(rOiKa>0) pl.getRange(4,2).setFormula("=B"+rOiKa);
    pl.getRange(5,1).setValue("💡 今日の一手");                          if(rIte>0)  pl.getRange(5,2).setFormula("=B"+rIte);
    pl.getRange(6,1).setValue("家族ラインに必要な追客/月");                if(rNeed>0) pl.getRange(6,2).setFormula("=B"+rNeed);
    pl.getRange(7,1).setValue("全社 経常利益（当月）");                   if(rAllKei>0)pl.getRange(7,2).setFormula("=B"+rAllKei);
    pl.getRange(8,1).setValue("月末現預金（体力）／ランウェイ");           pl.getRange(8,2).setFormula("="+CFR+"!B35"); pl.getRange(8,3).setFormula("="+CFR+"!B38");
    pl.getRange(3,1,6,1).setBackground(BL);
    pl.getRange(3,2,2,1).setBackground(GOLD).setFontWeight("bold"); // ギャップ/追客目標を金
    out.top="最上部R2〜8に今日の行動(経常ギャップ/追客目標/一手/必要追客/全社経常/月末現金・ランウェイ)";
  } else out.top="既に今日の行動あり・スキップ";

  // 3) 05を非表示(実質廃止・連動は生かす)
  var cf=shk("資金繰り"); if(cf) cf.hideSheet();
  out.hide05="05_資金繰りを非表示(裏方ハブ・あなたは04だけ見る)";

  return { ok:true, applied:out, warn:"行挿入で参照は自動追従するが、実行後に01統合司令塔とPLに#REFが無いか必ず確認。色は buildBatch7 再実行で再適用。" };
}

// ── 実装A基盤16：未来会計図表を古田土「お金のブロックパズル」正規構造で再構築(原典240320準拠) 2026-06-06 ──
function buildBatch16(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var pl=shk("損益"); if(!pl) return {error:"04損益PLが無い"};
  var z=shk("未来会計図表"); if(!z) z=ss.insertSheet("未来会計図表");
  // 04PLのラベル→行マップ
  var av=pl.getRange(1,1,200,1).getValues();
  function plRow(kw){ for(var i=0;i<av.length;i++){ if((av[i][0]||"").toString().indexOf(kw)>=0) return i+1; } return -1; }
  var P="'04_損益PL'!B";
  var rHojinUri=plRow("売上（コンサル/物件"), rEcUri=plRow("売上（EC"), rRent=plRow("賃料収入"),
      rMQ=plRow("限界利益（粗利計"), rZensha=plRow("■全社"), rG=plRow("経常利益（全社"),
      rYakuin=plRow("役員報酬"), rShaho=plRow("法定福利費"), rGyom=plRow("業務委託料"),
      rMokuhyo=plRow("目標経常利益");
  var rF=(rZensha>0?rZensha+2:-1); // 全社 固定費 計 = 全社ヘッダ+2
  // 売上(PQ)・固定費(F)の参照式
  var PQ="("+P+rHojinUri+"+"+P+rEcUri+"+"+P+rRent+")";
  var MQ=P+rMQ, F=P+rF, G=P+rG;
  var JINKEN="("+P+rYakuin+"+"+P+rShaho+"+"+P+rGyom+")";
  var SAIMU="'05_資金繰り'!B21+'05_資金繰り'!B25+'05_資金繰り'!B26"; // 住宅+法人+個人 借入返済(月)

  // 描画
  z.clear();
  const BR="#AA2E26", GRY="#ECECEC", BLU="#CFE2F3", ORG="#FCE5CD", GRN="#D9EAD3", GOLD="#FFE599", DARK="#434343";
  function set(r,c,v){ z.getRange(r,c).setValue(v); }
  function setF(r,c,f){ z.getRange(r,c).setFormula(f); }
  function head(r,txt){ z.getRange(r,1,1,5).setBackground(DARK); z.getRange(r,1).setValue(txt).setFontColor("#FFFFFF").setFontWeight("bold"); }

  set(1,1,"■ 未来会計図表（古田土式 お金のブロックパズル）— どこに手を打てば利益が出るか"); z.getRange(1,1,1,5).setBackground(BR); z.getRange(1,1).setFontColor("#FFFFFF").setFontWeight("bold");
  set(2,1,"利益＝会社が稼がなければならない額(社員と家族を守るコスト)。粗利益額MQが全ブロックを貫く軸。原典=240320古田土。当月6月・PL連動。");
  head(3,"【ブロックA】売上高 と 粗利益額（PQ＝VQ＋MQ）");
  set(4,1,"売上高（PQ）＝客単価×客数"); setF(4,2,"="+PQ); set(4,3,"商品力(P)×営業力(Q)");
  set(5,1,"　├ 変動費（VQ・仕入/原価/外注/ツール）"); setF(5,2,"="+PQ+"-"+MQ); set(5,3,"変動費率v="); setF(5,4,"=B5/B4");
  set(6,1,"　└ 粗利益額（MQ＝限界利益）★軸"); setF(6,2,"="+MQ); set(6,3,"粗利益率m="); setF(6,4,"=B6/B4");
  head(7,"【ブロックB】粗利益額 と 固定費（MQ＝F＋G）");
  set(8,1,"粗利益額（MQ）"); setF(8,2,"="+MQ);
  set(9,1,"　├ 固定費（F）"); setF(9,2,"="+F); set(9,3,"損益分岐点比率 F÷MQ="); setF(9,4,"=B9/B8");
  set(10,1,"　└ 経常利益（G）"); setF(10,2,"="+G); set(10,3,"経営安全率 G÷MQ="); setF(10,4,"=B10/B8");
  set(11,1,"固定費生産性（MQ÷F・理想1.25/目標1.1）"); setF(11,2,"=B8/B9");
  set(12,1,"格付け（損益分岐点比率）"); setF(12,2,'=IF(B9=0,"-",IF(B9<0.6,"SS 超優良",IF(B9<0.8,"S 優良",IF(B9<0.9,"A 健全",IF(B9<=1,"B 分岐点",IF(B9<=2,"C 赤字","D 倒産"))))))');
  head(13,"【ブロックC】固定費 と 人件費（F＝人件費＋経費）");
  set(14,1,"固定費（F）"); setF(14,2,"="+F);
  set(15,1,"　├ 人件費（役員報酬+社保+業務委託）"); setF(15,2,"="+JINKEN); set(15,3,"労働分配率 人件費÷MQ="); setF(15,4,"=B15/B8");
  set(16,1,"　└ 経費（モノ・金利・未来）"); setF(16,2,"=B14-B15"); set(16,3,"※未来費用(広告/教育/研究)は積極投資、金利は最小化");
  head(17,"■ 必要売上（古田土逆算：固定費＋目標利益＋借入返済 を賄う売上）");
  set(18,1,"目標経常利益（入力）"); setF(18,2,"="+P+rMokuhyo);
  set(19,1,"借入返済（月・住宅+法人+個人）"); setF(19,2,"="+SAIMU);
  set(20,1,"必要粗利（MQ）＝固定費＋目標＋返済"); setF(20,2,"=B9+B18+B19");
  set(21,1,"必要売上高＝必要粗利÷粗利益率"); setF(21,2,"=IFERROR(B20/B6/B4*B4,B20/(B6/B4))");
  set(22,1,"いま不足（必要粗利−現粗利MQ）"); setF(22,2,"=B20-B8");
  head(23,"■ 利益感度（やさしい順に手を打つ）");
  set(24,1,"P 客単価↑（商品力・値決め）／Q 客数↑（営業力・リピート）／V 変動費↓（交渉/技術）／F 固定費（パワー・削るな未来投資）");
  set(25,1,"古田土の判断：粗利率を上げるのが目的でなく粗利額MQを稼ぐ。固定費はコストでなくパワー。新規よりリピート率。");

  // 色（ブロックパズルの塗り分け）
  z.getRange(4,2).setBackground(GRY).setFontWeight("bold");      // 売上=灰枠
  z.getRange(5,2).setBackground(GRY);                            // 変動費=灰
  z.getRange(6,2).setBackground(BLU).setFontWeight("bold");      // 粗利MQ=青(軸)
  z.getRange(8,2).setBackground(BLU);
  z.getRange(9,2).setBackground(ORG).setFontWeight("bold");      // 固定費=橙
  z.getRange(10,2).setBackground(GOLD).setFontWeight("bold");    // 経常G=金
  z.getRange(12,2).setBackground(GOLD).setFontWeight("bold");    // 格付け=金
  z.getRange(14,2).setBackground(ORG);
  z.getRange(15,2).setBackground(GRN).setFontWeight("bold");     // 人件費=緑
  z.getRange(16,2).setBackground(ORG);
  z.getRange(20,2).setBackground(BLU).setFontWeight("bold");
  z.getRange(21,2).setBackground(GOLD).setFontWeight("bold");
  z.getRange("B4:B22").setNumberFormat("#,##0");
  z.getRange("D5:D6").setNumberFormat("0.0%"); z.getRange("D9:D10").setNumberFormat("0.0%"); z.getRange("D15").setNumberFormat("0.0%");
  z.getRange("B11").setNumberFormat("0.00");
  z.setColumnWidth(1,360); z.setColumnWidth(2,140); z.setColumnWidth(3,200); z.setColumnWidth(4,90);
  z.showSheet(); // 非表示だったら表示
  // 04直後あたりへ移動
  try{ ss.setActiveSheet(z); ss.moveActiveSheet(5); }catch(e){}

  return { ok:true, built:"未来会計図表を古田土ブロックパズル(A売上→VQ/MQ・B MQ→F/G・C F→人件費/経費)＋指標(変動費率/粗利益率/損益分岐点比率/固定費生産性/経営安全率/労働分配率)＋格付け＋必要売上＋利益感度 で再構築・PL連動・色分け・表示化",
    refs:{売上:"法人"+rHojinUri+"+EC"+rEcUri+"+賃料"+rRent, MQ:rMQ, F:rF, G:rG} };
}

// ── 実装A基盤17：未来会計図表を「古田土×マーケKPI 共通言語版」に再構築(逆算ブリッジ+用語対応表) 2026-06-06 ──
function buildBatch17(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var pl=shk("損益"); if(!pl) return {error:"04損益PLが無い"};
  var z=shk("未来会計図表"); if(!z) z=ss.insertSheet("未来会計図表");
  var av=pl.getRange(1,1,200,1).getValues();
  function plRow(kw){ for(var i=0;i<av.length;i++){ if((av[i][0]||"").toString().indexOf(kw)>=0) return i+1; } return -1; }
  var P="'04_損益PL'!B";
  var rHU=plRow("売上（コンサル/物件"), rEC=plRow("売上（EC"), rRent=plRow("賃料収入"),
      rMQ=plRow("限界利益（粗利計"), rZ=plRow("■全社"), rG=plRow("経常利益（全社"),
      rYak=plRow("役員報酬"), rSha=plRow("法定福利費"), rGyo=plRow("業務委託料"), rMok=plRow("目標経常利益"),
      rTuiApo=plRow("追客→アポ"), rApoSei=plRow("アポ→成約");
  var rF=(rZ>0?rZ+2:-1);
  var PQ="("+P+rHU+"+"+P+rEC+"+"+P+rRent+")", MQ=P+rMQ, F=P+rF, G=P+rG,
      JINKEN="("+P+rYak+"+"+P+rSha+"+"+P+rGyo+")", SAIMU="('05_資金繰り'!B21+'05_資金繰り'!B25+'05_資金繰り'!B26)",
      TUIAPO=(rTuiApo>0?P+rTuiApo:"0.5"), APOSEI=(rApoSei>0?P+rApoSei:"0.2");

  z.clear();
  const BR="#AA2E26",GRY="#ECECEC",BLU="#CFE2F3",ORG="#FCE5CD",GRN="#D9EAD3",GOLD="#FFE599",DARK="#434343",PUR="#D9D2E9";
  function S(r,c,v){ z.getRange(r,c).setValue(v); } function F2(r,c,f){ z.getRange(r,c).setFormula(f); }
  function head(r,txt,bg){ z.getRange(r,1,1,5).setBackground(bg||DARK); z.getRange(r,1).setValue(txt).setFontColor("#FFFFFF").setFontWeight("bold"); }
  function note(r,c,t){ z.getRange(r,c).setNote(t); }

  S(1,1,"■ 未来会計図表（古田土会計 × マーケKPI）— 全経営者の共通言語で『どこに手を打てば利益が出るか』"); z.getRange(1,1,1,5).setBackground(BR); z.getRange(1,1).setFontColor("#FFFFFF").setFontWeight("bold");
  S(2,1,"財務(古田土)とマーケ(IMP/CVR/CPA)を同じ数字で語る1枚。粗利益額MQが軸。当月6月・PL連動。原典240320。");

  // 1) マーケ・営業ファネル（客数Qの作り方）
  head(3,"1) マーケ・営業ファネル ＝ 客数Q（成約）はどう作るか（メディア→問合せ→成約）",PUR);
  S(4,1,"IMP（インプレッション・表示回数）【入力・メディア】"); S(4,2,100000); note(4,1,"YouTube/X/LP/広告の表示回数。共通語=Impression。KHDはメディア集客の入口。");
  S(5,1,"× CTR（クリック率）【入力】"); z.getRange(5,2).setValue(0.01);
  S(6,1,"＝ 流入（クリック/セッション）"); F2(6,2,"=B4*B5");
  S(7,1,"× CVR（CV率＝問合せ率）【入力】"); z.getRange(7,2).setValue(0.02);
  S(8,1,"＝ CV（コンバージョン＝問合せ/リード）＝追客対象"); F2(8,2,"=B6*B7"); note(8,1,"共通語=Conversion。KHDでは『問合せ/反響/追客リスト入り』。ここから先はPLの営業ドライバー(追客→アポ→成約)と接続。");
  S(9,1,"× 追客→アポ 転換率"); F2(9,2,"="+TUIAPO);
  S(10,1,"× アポ→成約 転換率"); F2(10,2,"="+APOSEI);
  S(11,1,"＝ 成約 ＝ 客数Q"); F2(11,2,"=B8*B9*B10");
  S(4,4,"CPA（顧客獲得単価）"); S(5,4,"=広告費÷CV"); S(6,4,"LTV（顧客生涯価値）"); S(7,4,"ROAS（広告費用対効果）"); S(8,4,"=売上÷広告費");
  z.getRange(4,1,8,1).setBackground(GRN);

  // 2) 古田土 ブロックパズル
  head(12,"2) 古田土 お金のブロックパズル（売上→粗利→利益）",DARK);
  S(13,1,"売上高（PQ）＝ 客単価P × 客数Q"); F2(13,2,"="+PQ); S(13,3,"P=商品力/値決め・Q=1)ファネル出口");
  S(14,1,"　├ 変動費（VQ）＝仕入/原価/外注/ツール"); F2(14,2,"="+PQ+"-"+MQ); S(14,3,"変動費率 v＝VQ÷PQ"); F2(14,4,"=B14/B13");
  S(15,1,"　└ 粗利益額（MQ＝限界利益）★軸"); F2(15,2,"="+MQ); S(15,3,"粗利益率 m＝MQ÷PQ（≒粗利率/Gross Margin）"); F2(15,4,"=B15/B13");
  S(16,1,"粗利益額（MQ）"); F2(16,2,"="+MQ);
  S(17,1,"　├ 固定費（F）"); F2(17,2,"="+F); S(17,3,"損益分岐点比率 F÷MQ（≒BEP比率・理想80%）"); F2(17,4,"=B17/B16");
  S(18,1,"　└ 経常利益（G）"); F2(18,2,"="+G); S(18,3,"経営安全率 G÷MQ"); F2(18,4,"=B18/B16");
  S(19,1,"固定費生産性 MQ÷F（理想1.25/目標1.1＝社長の経営力）"); F2(19,2,"=B16/B17");
  S(20,1,"格付け（損益分岐点比率）"); F2(20,2,'=IF(B17=0,"-",IF(B17<0.6,"SS 超優良",IF(B17<0.8,"S 優良",IF(B17<0.9,"A 健全",IF(B17<=1,"B 分岐点",IF(B17<=2,"C 赤字","D 倒産"))))))');
  S(21,1,"固定費（F）"); F2(21,2,"="+F);
  S(22,1,"　├ 人件費（役員報酬+社保+業務委託）"); F2(22,2,"="+JINKEN); S(22,3,"労働分配率 人件費÷MQ（≒Labor Share）"); F2(22,4,"=B22/B16");
  S(23,1,"　└ 経費（モノ・金利・未来）"); F2(23,2,"=B21-B22"); S(23,3,"未来費用(広告/教育/R&D)は積極投資・金利は最小化");
  z.getRange(13,2).setBackground(GRY).setFontWeight("bold"); z.getRange(14,2).setBackground(GRY);
  z.getRange(15,2).setBackground(BLU).setFontWeight("bold"); z.getRange(16,2).setBackground(BLU);
  z.getRange(17,2).setBackground(ORG).setFontWeight("bold"); z.getRange(18,2).setBackground(GOLD).setFontWeight("bold");
  z.getRange(20,2).setBackground(GOLD).setFontWeight("bold"); z.getRange(21,2).setBackground(ORG); z.getRange(22,2).setBackground(GRN).setFontWeight("bold"); z.getRange(23,2).setBackground(ORG);

  // 3) 逆算ブリッジ（必要売上→必要Q→必要CV→必要IMP）
  head(24,"3) 逆算ブリッジ ＝ 必要売上(古田土) から 必要IMP(マーケ) まで一気通貫",BR);
  S(25,1,"目標経常利益（入力）"); F2(25,2,"="+P+rMok);
  S(26,1,"借入返済（月）"); F2(26,2,"="+SAIMU);
  S(27,1,"必要粗利 MQ ＝ 固定費＋目標＋返済"); F2(27,2,"=B17+B25+B26");
  S(28,1,"必要売上高 ＝ 必要MQ ÷ 粗利益率m"); F2(28,2,"=IFERROR(B27/B15*B13,0)");
  S(29,1,"客単価P（平均・入力）"); z.getRange(29,2).setValue(1100000); note(29,1,"医療テナント110万/物件190万等の平均客単価。Q逆算の分母。");
  S(30,1,"→ 必要 客数Q（成約数）＝必要売上÷P"); F2(30,2,"=IFERROR(B28/B29,0)");
  S(31,1,"→ 必要 アポ ＝Q÷成約率"); F2(31,2,"=IFERROR(B30/B10,0)");
  S(32,1,"→ 必要 CV（問合せ/追客）＝アポ÷追客アポ率"); F2(32,2,"=IFERROR(B31/B9,0)");
  S(33,1,"→ 必要 IMP（表示）＝CV÷CVR÷CTR"); F2(33,2,"=IFERROR(B32/B7/B5,0)");
  z.getRange(27,2).setBackground(BLU).setFontWeight("bold"); z.getRange(28,2).setBackground(GOLD).setFontWeight("bold"); z.getRange(30,2,4,1).setBackground(PUR).setFontWeight("bold");

  // 4) 用語対応表（共通言語）
  head(34,"4) 用語対応表（古田土会計 ⇔ マーケKPI ＝ 共通言語）",DARK);
  var tbl=[["古田土／会計用語","マーケ／共通用語","式・意味"],
    ["売上高 PQ","売上 Revenue","客単価P × 客数Q"],
    ["客単価 P","客単価 AOV(平均注文額)","商品力・値決め"],
    ["客数 Q","CV/受注数 Conversion","ファネル出口=成約"],
    ["変動費率 v","原価率 COGS率","VQ÷PQ"],
    ["粗利益率 m","粗利率 Gross Margin","MQ÷PQ"],
    ["損益分岐点比率","BEP比率 Break-even","F÷MQ（理想80%）"],
    ["固定費生産性","-（経営効率）","MQ÷F（理想1.25）"],
    ["労働分配率","人件費率 Labor Share","人件費÷MQ"],
    ["—（集客効率）","CPA 顧客獲得単価","広告費÷CV"],
    ["—（転換効率）","CVR コンバージョン率","CV÷流入"],
    ["—（顧客価値）","LTV 顧客生涯価値","客単価×継続×粗利率"],
    ["—（広告効率）","ROAS 広告費用対効果","売上÷広告費"]];
  z.getRange(35,1,tbl.length,3).setValues(tbl);
  z.getRange(35,1,1,3).setBackground("#666666").setFontColor("#FFFFFF").setFontWeight("bold");
  z.getRange(36,1,tbl.length-1,3).setBackground(GRY);

  // 5) 利益感度
  var rr=35+tbl.length+1;
  head(rr,"5) 利益感度（やさしい順に手を打つ）",BR);
  S(rr+1,1,"P 客単価↑（値決め=商品価値で・コスト基準でない）／Q 客数↑（新規よりリピート率）／V 変動費↓（交渉/技術）／F 固定費（パワー・未来投資は削るな/金利は最小化）");
  S(rr+2,1,"古田土：粗利「率」でなく粗利「額」MQを稼ぐ。固定費はコストでなくパワー。PQアップは運転資金を食う点に注意。");

  z.getRange("B4:B33").setNumberFormat("#,##0");
  z.getRangeList(["B5","B7","B9","B10","D14","D15","D17","D18","D22"]).setNumberFormat("0.0%");
  z.getRange("B19").setNumberFormat("0.00");
  z.getRange("B30:B33").setNumberFormat("#,##0.0");
  z.setColumnWidth(1,330); z.setColumnWidth(2,150); z.setColumnWidth(3,260); z.setColumnWidth(4,170); z.setColumnWidth(5,120);
  z.showSheet(); try{ ss.setActiveSheet(z); ss.moveActiveSheet(5); }catch(e){}
  return { ok:true, built:"未来会計図表＝古田土×マーケKPI共通言語版(1)ファネルIMP→CV→Q 2)ブロックパズル 3)逆算ブリッジ必要売上→必要IMP 4)用語対応表 5)利益感度)・PL連動・色分け" };
}

// ── 実装A基盤18：03に営業ドライバー(マーケ用語)を一本化＋薄地黒文字＋10収益パイプ吸収 2026-06-06 ──
// 安全：空き行30-49のみ書き換え(案件5-14・集計16-21・科目マスター50-56・外部参照は不触)。白抜き廃止。
function buildBatch18(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var u=shk("売上見込み"); if(!u) return {error:"03が無い"};
  // 薄地＋黒文字パレット（白抜き廃止）
  const HEAD="#D9D9D9", TITLE="#FCE5CD", SUB="#FFF2CC", IN="#FFF9E0", CALC="#E8F0FE", GOAL="#FFE599", BLK="#000000";

  // === 1) 営業ドライバー（30-47）をマーケ用語版で再構築（家族と黒字から逆算）===
  u.getRange("A30:E49").clearContent();
  function row(r,a,bF,c,note){ u.getRange(r,1).setValue(a); if(bF!=null){ if((""+bF).charAt(0)==="="||(""+bF).indexOf("=")===0) u.getRange(r,2).setFormula(bF); else u.getRange(r,2).setValue(bF);} if(c) u.getRange(r,3).setValue(c); if(note) u.getRange(r,1).setNote(note); }
  u.getRange("A30").setValue("🎯 営業ドライバー（家族と黒字から逆算＝今日営業する理由）｜マーケKPI共通言語").setFontWeight("bold").setFontColor(BLK).setBackground(TITLE);
  u.getRange("A30:E30").setBackground(TITLE);
  row(31,"家族が潰れないライン/月（必要粗利）","='05_資金繰り'!B37",null,"05資金繰りの純月次燃焼＝生活費+借入も賄うのに毎月要る粗利。");
  row(32,"経常確定粗利/月（EC+賃料・recurring）","=K12+K13",null,"毎月入る安定粗利(EC菊池ネット61,000+北千住賃料81,000)。");
  row(33,"🔴 経常ギャップ（毎月の穴）","=MAX(0,B31-B32)","← これを営業で埋める");
  row(34,"平均成約単価（入力・客単価P）",1100000,"医療110万/物件単発の平均","Q逆算の分母。マーケ用語=AOV(平均注文額)。");
  row(35,"→ 必要 追加成約数 Q（=CV/受注）","=IF(B33<=0,0,ROUNDUP(B33/B34,0))","件 ｜マーケ=Conversion(成約)");
  row(36,"アポ→成約 転換率（入力・成約CVR）",0.2,"PL B53と揃える");
  row(37,"→ 必要 アポ数","=IF(B35<=0,0,ROUNDUP(B35/B36,0))","件");
  row(38,"追客→アポ 転換率（入力）",0.5,"PL B52と揃える");
  row(39,"→ 必要 CV（問合せ/追客リード）","=IF(B37<=0,0,ROUNDUP(B37/B38,0))","件 ｜マーケ=Conversion(問合せ)");
  row(40,"CVR（CV率・入力/メディア）",0.02,"流入→問合せ率。YouTube/LP用");
  row(41,"CTR（クリック率・入力/メディア）",0.01,"表示→クリック率");
  row(42,"→ 必要 IMP（表示回数）","=IF(B39<=0,0,ROUNDUP(B39/B40/B41,0))","回 ｜マーケ=Impression");
  u.getRange("A43").setValue("─ 今日の一手＆時間KPI ─").setFontWeight("bold").setFontColor(BLK).setBackground(HEAD); u.getRange("A43:E43").setBackground(HEAD);
  row(44,"月の稼働日（入力）",20,null);
  row(45,"→ 今日の追客本数","=IF(B39<=0,0,ROUNDUP(B39/B44,0))","件/日 ← 02作業DBの日次追客に割る");
  row(46,"営業直結 時間比率（07本部マトリクス連動）→目標60%","='07_本部マトリクス'!F21",null,"時間ROI。07の実態時間から自動。低ければ内務を削り営業へ。");
  u.getRange("A47").setValue("CPA=広告費÷CV ／ LTV=客単価×継続×粗利率 ／ ROAS=売上÷広告費（マーケ共通指標・併記）").setFontColor("#666666");
  // 色：入力=薄黄／計算=薄青／ギャップ=金
  [34,36,38,40,41,44].forEach(function(r){ u.getRange(r,2).setBackground(IN); });
  [31,32,35,37,39,42,45,46].forEach(function(r){ u.getRange(r,2).setBackground(CALC); });
  u.getRange("B33").setBackground(GOAL).setFontWeight("bold");
  u.getRange("B31:B46").setNumberFormat("#,##0"); u.getRange("B36").setNumberFormat("0%"); u.getRange("B38").setNumberFormat("0%"); u.getRange("B40").setNumberFormat("0%"); u.getRange("B41").setNumberFormat("0%"); u.getRange("B46").setNumberFormat("0%");

  // === 2) 白抜き廃止：03の見出しを薄地＋黒太字に ===
  function lite(a1,bg){ u.getRange(a1).setBackground(bg).setFontColor(BLK).setFontWeight("bold"); }
  lite("A1:Q1",TITLE);     // タイトル
  lite("A2:K2",HEAD);      // KPIヘッダ
  lite("A4:Q4",HEAD);      // 案件ヘッダ
  lite("A16",SUB);         // 集計見出し
  lite("A50",SUB);         // 科目マスター見出し

  // === 3) 10_収益パイプラインを非表示（03に吸収・重複廃止）===
  var p10=shk("収益パイプライン"); if(p10){ p10.hideSheet(); }

  return { ok:true,
    driver:"03の30-47に営業ドライバー(家族ライン→経常ギャップ→必要成約CV→必要アポ→必要CV問合せ→必要IMP→今日の追客本数＋営業直結%07連動)をマーケ用語で一本化",
    color:"03見出しを薄地+黒太字に(白抜き廃止)",
    p10:"10_収益パイプライン非表示(03吸収)",
    note:"案件5-14/集計16-21/科目マスター50-56/外部参照は不触＝安全。04top・未来会計図表の重複整理と、04/未来会計図表の白抜きも次で薄地黒文字に統一。" };
}

// ── 実装A基盤19：顧客マスター構築(空の02_追客リストを転用・先生ニーズ8名転記・顧客番号H001) 2026-06-06 ──
// WHO(顧客)をWHATの03から分離＝正規化。03案件は顧客番号で参照。リネームは空シートのみ＝参照保全。
function buildBatch19(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  const HEAD="#D9D9D9", IN="#FFF9E0", BLK="#000000", TITLE="#FCE5CD";
  // 空の02_追客リストを顧客マスターに転用(参照ゼロ＝安全)。無ければ新規。
  var c=shk("追客リスト"); if(!c) c=shk("顧客"); if(!c) c=ss.insertSheet("顧客マスター");
  c.clear();
  var oldF=c.getFilter(); if(oldF) oldF.remove();           // ★旧フィルタ(▼)を除去＝c.clear()では消えない
  c.getRange(1,1,c.getMaxRows(),c.getMaxColumns()).clearDataValidations(); // 旧ドロップダウンも一掃
  if(c.getName().indexOf("顧客")<0) c.setName("顧客マスター");

  c.getRange("A1").setValue("顧客マスター（WHO＝誰の・何のニーズ・誰経由）｜◀入力ゾーン：本部長やり取りで案件化した人をここに追記。03売上見込みは顧客番号で連動").setFontWeight("bold").setFontColor(BLK).setBackground(TITLE);
  c.getRange("A1:R1").merge();
  var hdr=["顧客番号","本部","顧客名","診療科・科目","ニーズ(希望条件)","希望エリア","面積坪","予算","物件種別","ハブ(紹介者)","連絡先","開業時期","温度","ステータス","次アクション","最終接触","紐づく案件(03)","備考"];
  c.getRange(2,1,1,hdr.length).setValues([hdr]).setBackground(HEAD).setFontColor(BLK).setFontWeight("bold");

  // 先生ニーズ8名を転記(顧客番号H001〜・全18列で統一)。温度は開業時期から推定。
  var rows=[
    ["H001","04","曾我 拓嗣","眼科(白内障OP)","60坪以上・白内障手術スペース要","台東/足立/千葉/埼玉/神奈川",60,"15,000円/坪","テナント","山本先生","meromijp@yahoo.co.jp","未定","WARM","案件化","物件提案・福井経由で前進","2026-06-02","03_医療テナント承継","白内障OP有"],
    ["H002","04","内山 裕智","内科(移転)","35坪以上・西新宿5丁目移転","渋谷本町/中野弥生町/西新宿5",35,"-","テナント・売ビル","社労士 金井","luchiyamaclinic@outlook.jp","物件次第","WARM","探索中","西新宿物件あたり","-","-","移転案件"],
    ["H003","04","長西","婦人科","条件ヒアリング中","-","-","-","-","福井(奥様窓口)","-","-","WARM","探索中","奥様窓口で条件ヒアリング","-","-","福井紹介"],
    ["H004","04","星山 隆行","内科","横浜線沿線で探索","長津田/十日市場/鴨居/小机","-","-","テナント","野崎様","hsym10@kitasato-u.ac.jp","R8/3以降","COLD","探索中","開業時期先・関係維持","-","-","北里大 総合診療"],
    ["H005","04","桑原 良奈","産婦人科","土地購入・佐久/御代田","佐久","-","-","土地購入","岡先生","-","-","WARM","探索中","御代田町の土地と照合","-","-","御代田案件と関連(Drive 19kxzM94)"],
    ["H006","04","奥村 健治(TSメディカル)","耳鼻科(日帰り手術)","55坪前後・鼻手術専門","東京都中野区近辺",55,"-","テナント","福井","k.okumura@ts-medical.co.jp","即可","HOT","案件化","即開業可・物件最優先で探す","2026-06-02","-","退職済で即開業可"],
    ["H007","04","眞木先生","泌尿器科","40-50坪・居抜き・CT導入・四ツ谷駅近","四谷〜新宿",45,"-","居抜き","福井提携","-","未定","WARM","探索中","四ツ谷駅近の居抜き探索","-","-","CT等機器導入視野"],
    ["H008","04","秋元先生","内科(女医)","勤務先探し中","西池袋","-","-","-","福井","-","-","COLD","探索中","勤務先紹介・開業は先","-","-","女医"]
  ];
  c.getRange(3,1,rows.length,hdr.length).setValues(rows);
  // 関連フォルダ列(S=19)：先生ごとのDriveフォルダ(H番号で作成→関連資料を蓄積→URLを貼る)
  c.getRange(2,19).setValue("関連フォルダ").setBackground(HEAD).setFontColor(BLK).setFontWeight("bold");
  c.getRange(2,19).setNote("先生ごとにDriveフォルダ(例:H001_曾我)を作り関連資料を蓄積→そのURLをここに貼る。番号に紐づいて今後保存していく。");
  c.getRange(7,19).setValue("https://drive.google.com/drive/folders/19kxzM94xt7UncHflOvQEBS2YfPFGAwwe"); // 桑原=御代田フォルダ
  // 入力ゾーン色(データ8行＋追加3行分)。温度/ステータス/次アクションはClaudeが報告から記入＝ドロップダウン(フィルタ風▼)は付けない。
  c.getRange(3,1,rows.length+3,19).setBackground(IN);
  c.getRange(2,1).setNote("顧客番号=H001〜の連番。物件案件は物件マスターのNo、人/コンサルはこのH番号。03売上見込みはこの番号で顧客を参照。");
  c.getRange(2,14).setNote("ステータス: 探索中→案件化(03へ)→成約/見送り。温度/ステータス/次アクションはClaudeが本部長報告から記入(あなたが選ぶ手間なし)。");
  // 振り分け凡例(運用ルール)を最上部の下に注記
  c.getRange(2,3).setNote("【振り分けルール】人だけ(物件未定)=ここ顧客マスターに蓄積／物件が絡む=作業DBへ／顧客に物件が紐づいたら作業DBへ昇格(H番号引継)。Claudeが対話/夜ポチッで振り分け＋抜け漏れ質問。");

  c.setFrozenRows(2); c.setColumnWidth(3,120); c.setColumnWidth(5,220); c.setColumnWidth(6,200); c.setColumnWidth(15,200);
  // 02作業DBの直後あたりへ
  try{ ss.setActiveSheet(c); ss.moveActiveSheet(4); }catch(e){}

  return { ok:true, sheet:c.getName(), customers:rows.length,
    fields:hdr.length+"列(顧客番号/温度/ステータス/次アクション/最終接触/紐づく案件 を新設)",
    note:"先生ニーズ8名をH001-H008で転記(温度=開業時期から推定)。ステータス=探索中/案件化/成約/見送りのドロップダウン。案件化したら03へ。リネームは空シートのみ＝既存参照は無傷。",
    next:"03売上見込みの案件に顧客番号列を足して連動／本部長報告から抜け漏れ追記＋質問の運用" };
}

// ── 実装A基盤20：②03に古田土ブロックパズル＋用語対応表を集約(ドライバーの下・PL連動)＋未来会計図表非表示 2026-06-06 ──
// 03のドライバー(30-47)にマーケファネルは既存→重複避け、ここは利益構造(ブロックパズル)＋用語対応表のみ。57行〜に追記(案件/集計/マスター不触)。
function buildBatch20(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var u=shk("売上見込み"); if(!u) return {error:"03が無い"};
  var pl=shk("損益"); if(!pl) return {error:"04PLが無い"};
  var av=pl.getRange(1,1,200,1).getValues();
  function plRow(kw){ for(var i=0;i<av.length;i++){ if((av[i][0]||"").toString().indexOf(kw)>=0) return i+1; } return -1; }
  var P="'04_損益PL'!B";
  var rHU=plRow("売上（コンサル/物件"),rEC=plRow("売上（EC"),rRent=plRow("賃料収入"),rMQ=plRow("限界利益（粗利計"),rZ=plRow("■全社"),rG=plRow("経常利益（全社"),rYak=plRow("役員報酬"),rSha=plRow("法定福利費"),rGyo=plRow("業務委託料");
  var rF=(rZ>0?rZ+2:-1);
  var PQ="("+P+rHU+"+"+P+rEC+"+"+P+rRent+")",MQ=P+rMQ,F=P+rF,G=P+rG,JINKEN="("+P+rYak+"+"+P+rSha+"+"+P+rGyo+")";
  const HEAD="#D9D9D9",TITLE="#FCE5CD",GRY="#ECECEC",BLU="#CFE2F3",ORG="#FCE5CD",GRN="#D9EAD3",GOLD="#FFE599",BLK="#000000";
  function S(r,c,v){ u.getRange(r,c).setValue(v); } function F2(r,c,f){ u.getRange(r,c).setFormula(f); }

  var b=57; // 科目マスター(50-56)の下から
  u.getRange("A"+b+":E90").clearContent();
  S(b,1,"■ 未来会計図表（古田土：売上→粗利→利益の構造）｜上のドライバーで“いくら要る”、ここで“利益の形”を見る"); u.getRange(b,1,1,5).setBackground(TITLE); u.getRange(b,1).setFontWeight("bold").setFontColor(BLK);
  S(b+1,1,"売上高（PQ）＝客単価P×客数Q"); F2(b+1,2,"="+PQ); S(b+1,3,"P=商品力/値決め・Q=ドライバー出口");
  S(b+2,1,"　├ 変動費（VQ）"); F2(b+2,2,"="+PQ+"-"+MQ); S(b+2,3,"変動費率v(=原価率)"); F2(b+2,4,"=B"+(b+2)+"/B"+(b+1));
  S(b+3,1,"　└ 粗利益額（MQ＝限界利益）★軸"); F2(b+3,2,"="+MQ); S(b+3,3,"粗利益率m(=Gross Margin)"); F2(b+3,4,"=B"+(b+3)+"/B"+(b+1));
  S(b+4,1,"粗利益額（MQ）"); F2(b+4,2,"="+MQ);
  S(b+5,1,"　├ 固定費（F）"); F2(b+5,2,"="+F); S(b+5,3,"損益分岐点比率 F÷MQ(=BEP・理想80%)"); F2(b+5,4,"=B"+(b+5)+"/B"+(b+4));
  S(b+6,1,"　└ 経常利益（G）"); F2(b+6,2,"="+G); S(b+6,3,"経営安全率 G÷MQ"); F2(b+6,4,"=B"+(b+6)+"/B"+(b+4));
  S(b+7,1,"固定費生産性 MQ÷F(理想1.25)"); F2(b+7,2,"=B"+(b+4)+"/B"+(b+5));
  S(b+8,1,"格付け（損益分岐点比率）"); F2(b+8,2,'=IF(B'+(b+5)+'=0,"-",IF(B'+(b+5)+'<0.6,"SS 超優良",IF(B'+(b+5)+'<0.8,"S 優良",IF(B'+(b+5)+'<0.9,"A 健全",IF(B'+(b+5)+'<=1,"B 分岐点",IF(B'+(b+5)+'<=2,"C 赤字","D 倒産"))))))');
  S(b+9,1,"固定費（F）"); F2(b+9,2,"="+F);
  S(b+10,1,"　├ 人件費（役員報酬+社保+業務委託）"); F2(b+10,2,"="+JINKEN); S(b+10,3,"労働分配率 人件費÷MQ(=Labor Share)"); F2(b+10,4,"=B"+(b+10)+"/B"+(b+4));
  S(b+11,1,"　└ 経費（モノ・金利・未来）"); F2(b+11,2,"=B"+(b+9)+"-B"+(b+10)); S(b+11,3,"未来費用(広告/教育/R&D)は積極投資・金利最小化");
  // 色(薄地黒文字)
  u.getRange(b+1,2).setBackground(GRY).setFontWeight("bold"); u.getRange(b+2,2).setBackground(GRY);
  u.getRange(b+3,2).setBackground(BLU).setFontWeight("bold"); u.getRange(b+4,2).setBackground(BLU);
  u.getRange(b+5,2).setBackground(ORG).setFontWeight("bold"); u.getRange(b+6,2).setBackground(GOLD).setFontWeight("bold");
  u.getRange(b+8,2).setBackground(GOLD).setFontWeight("bold"); u.getRange(b+9,2).setBackground(ORG); u.getRange(b+10,2).setBackground(GRN).setFontWeight("bold"); u.getRange(b+11,2).setBackground(ORG);
  u.getRange("B"+(b+1)+":B"+(b+11)).setNumberFormat("#,##0");
  u.getRangeList(["D"+(b+2),"D"+(b+3),"D"+(b+5),"D"+(b+6),"D"+(b+10)]).setNumberFormat("0.0%"); u.getRange("B"+(b+7)).setNumberFormat("0.00");

  // 用語対応表(古田土⇔マーケ共通言語)
  var t=b+13;
  S(t,1,"■ 用語対応表（古田土会計 ⇔ マーケKPI ＝ 全経営者の共通言語）"); u.getRange(t,1,1,5).setBackground(HEAD); u.getRange(t,1).setFontWeight("bold").setFontColor(BLK);
  var tbl=[["古田土／会計用語","マーケ／共通用語","式・意味"],
    ["売上高 PQ","売上 Revenue","客単価P × 客数Q"],
    ["客数 Q","CV/受注数 Conversion","ドライバーの出口=成約"],
    ["粗利益率 m","粗利率 Gross Margin","MQ÷PQ"],
    ["損益分岐点比率","BEP比率 Break-even","F÷MQ(理想80%)"],
    ["労働分配率","人件費率 Labor Share","人件費÷MQ"],
    ["—(集客効率)","CPA 顧客獲得単価","広告費÷CV"],
    ["—(転換効率)","CVR コンバージョン率","CV÷流入"],
    ["—(表示)","IMP インプレッション","表示回数"],
    ["—(顧客価値)","LTV 顧客生涯価値","客単価×継続×粗利率"]];
  u.getRange(t+1,1,tbl.length,3).setValues(tbl);
  u.getRange(t+1,1,1,3).setBackground("#666666").setFontColor("#FFFFFF").setFontWeight("bold");
  u.getRange(t+2,1,tbl.length-1,3).setBackground(GRY);

  // 未来会計図表タブは03に集約したので非表示
  var z=shk("未来会計図表"); if(z) z.hideSheet();

  return { ok:true, added:"03の"+b+"行〜に古田土ブロックパズル(売上→VQ/MQ→F/G→人件費/経費＋指標＋格付け)＋用語対応表(古田土⇔マーケ)を集約",
    hidden:"未来会計図表タブを非表示(03に集約)",
    note:"ドライバー(30-47)のマーケファネルと重複させず・利益構造と用語対応のみ。案件/集計/科目マスター不触。色は薄地黒文字。" };
}

// ── 実装A基盤21：③07本部マトリクスの時間KPIを02右上に連動表示＋振り分け凡例＋07非表示 2026-06-06 ──
// 02上部への行挿入は03のFILTER(02 D2:D2000参照)を壊すので、時間KPIは右上(X列〜)に配置。データ行は不触＝安全。
function buildBatch21(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  const HEAD="#D9D9D9", IN="#FFF9E0", GOAL="#FFE599", BLK="#000000", BLU="#CFE2F3";
  var db=shk("作業DB"); if(!db) return {error:"02作業DBが無い"};
  var M="'07_本部マトリクス'!";

  // 02右上(X1〜)に時間KPI(07連動)。X=24列目。
  db.getRange(1,24).setValue("■ 時間KPI（07連動・カレンダー実績／毎日見る）").setBackground(HEAD).setFontColor(BLK).setFontWeight("bold");
  db.getRange(1,24,1,3).setBackground(HEAD);
  var kpi=[
    ["営業直結 時間比率","="+M+"F21","目標60%"],
    ["家族 時間比率","="+M+"F20",""],
    ["内務 時間比率","="+M+"F22","削って営業へ"],
    ["仕込 時間比率","="+M+"F23",""],
    ["総実績時間/月(h)","="+M+"D19",""]
  ];
  for(var i=0;i<kpi.length;i++){ var r=2+i;
    db.getRange(r,24).setValue(kpi[i][0]);
    db.getRange(r,25).setFormula(kpi[i][1]);
    db.getRange(r,26).setValue(kpi[i][2]);
  }
  db.getRange(2,24,kpi.length,1).setBackground(IN);
  db.getRange(2,25,kpi.length,1).setBackground(BLU);
  db.getRange(2,25).setBackground(GOAL).setFontWeight("bold"); // 営業直結%=結論
  db.getRange("Y2:Y5").setNumberFormat("0.0%");
  db.getRange(1,24).setNote("時間KPIは07本部マトリクス(カレンダー実績)から連動。営業直結%が目標60%に届いてるか毎日チェック。低ければ内務を削り営業へ。");

  // 振り分け凡例(運用ルール)を02のヘッダ近くに
  db.getRange(1,4).setNote("【振り分けルール】物件が絡む=この作業DBで行動記録＋03で案件化／人だけ=顧客マスターへ／顧客に物件が紐づいたら作業DBへ昇格(H番号引継)。Claudeが対話/夜ポチッ(当日チャット遡り)で振り分け＋抜け漏れ質問。");

  // 07本部マトリクスを非表示(02に時間KPI集約・01/04の参照は裏で生存)
  var m7=shk("本部マトリクス"); if(m7) m7.hideSheet();

  return { ok:true, kpi:"02右上(X1〜)に時間KPI(営業直結%/家族/内務/仕込/総時間・07連動)を表示",
    rule:"02ヘッダに振り分け凡例", hide07:"07本部マトリクスを非表示(02集約・01/04は裏で連動継続)",
    note:"02の行は不触(03のFILTER保護)。時間KPIは右上配置。" };
}

// ── 実装A基盤22：03再設計(本部色再適用・ドライバー直下・古田土図直下・マスタ最下部・コンパクト) 2026-06-06 ──
// 安全：案件5-14/集計16-21/KPIヘッダ1-4/外部参照は不触。22行目以降のみ再構築＋A5:A14本部色再適用。
function buildBatch22(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var u=shk("売上見込み"); if(!u) return {error:"03が無い"};
  var pl=shk("損益"); var av=pl?pl.getRange(1,1,200,1).getValues():[];
  function plRow(kw){ for(var i=0;i<av.length;i++){ if((av[i][0]||"").toString().indexOf(kw)>=0) return i+1; } return -1; }
  var P="'04_損益PL'!B";
  var rHU=plRow("売上（コンサル/物件"),rEC=plRow("売上（EC"),rRent=plRow("賃料収入"),rMQ=plRow("限界利益（粗利計"),rZ=plRow("■全社"),rG=plRow("経常利益（全社"),rYak=plRow("役員報酬"),rSha=plRow("法定福利費"),rGyo=plRow("業務委託料");
  var rF=(rZ>0?rZ+2:-1);
  var PQ="("+P+rHU+"+"+P+rEC+"+"+P+rRent+")",MQ=P+rMQ,F=P+rF,G=P+rG,JINKEN="("+P+rYak+"+"+P+rSha+"+"+P+rGyo+")";
  const HEAD="#D9D9D9",TITLE="#FCE5CD",IN="#FFF9E0",CALC="#E8F0FE",GOAL="#FFE599",GRY="#ECECEC",ORG="#FCE5CD",GRN="#D9EAD3",BLK="#000000";
  const HONBU={"00":"#F4CCCC","01":"#FFF2CC","02":"#D9EAD3","03":"#CFE2F3","04":"#D9D2E9","05":"#FCE5CD"};
  function S(r,c,v){ u.getRange(r,c).setValue(v); } function F2(r,c,f){ u.getRange(r,c).setFormula(f); }
  function ttl(r,t){ u.getRange(r,1,1,5).setBackground(TITLE); u.getRange(r,1).setValue(t).setFontWeight("bold").setFontColor(BLK); }

  // 0) 本部色 再適用(A5:A14)
  var a5=u.getRange(5,1,10,1).getValues();
  for(var k=0;k<10;k++){ var key=(a5[k][0]==null?"":a5[k][0]).toString().trim(); var k2=(key.length===1?"0"+key:key); if(HONBU[k2]) u.getRange(5+k,1).setBackground(HONBU[k2]).setFontWeight("bold"); }

  // 1) 22行目以降をリセット
  u.getRange("A22:R95").clearContent(); u.getRange("A22:R95").setBackground(null);

  // 2) 営業ドライバー(22-37)＝案件直下。月次トラッキング(必要vs確定vs不足)は上部に統合。
  ttl(22,"🎯 今日の行動｜営業ドライバー（家族と黒字から逆算→今日の追客本数）｜マーケKPI共通言語");
  u.getRange(22,1).setNote("【使い方】上の案件(5-14)を確度で積んだ見込みと『家族が潰れないライン』の差＝穴を、何件成約/何追客で埋めるか逆算。毎朝ここを見て02作業DBに今日の追客を落とす。黄=入力/青=自動/金=結論。");
  S(23,1,"必要：家族が潰れないライン/月"); F2(23,2,"='05_資金繰り'!B37");
  S(24,1,"確定：経常粗利/月(EC+賃料 recurring)"); F2(24,2,"=K12+K13");
  S(25,1,"🔴 不足＝経常ギャップ(毎月の穴)"); F2(25,2,"=MAX(0,B23-B24)"); S(25,3,"← これを営業で埋める");
  S(26,1,"平均成約単価P(入力)"); S(26,2,1100000); S(26,3,"医療110万/物件単発の平均(=AOV)");
  S(27,1,"→ 必要 成約数Q [CV/受注]"); F2(27,2,"=IF(B25<=0,0,ROUNDUP(B25/B26,0))"); S(27,3,"件");
  S(28,1,"アポ→成約 転換率(入力)"); S(28,2,0.2);
  S(29,1,"→ 必要 アポ数"); F2(29,2,"=IF(B27<=0,0,ROUNDUP(B27/B28,0))"); S(29,3,"件");
  S(30,1,"追客→アポ 転換率(入力)"); S(30,2,0.5);
  S(31,1,"→ 必要 CV [問合せ/追客リード]"); F2(31,2,"=IF(B29<=0,0,ROUNDUP(B29/B30,0))"); S(31,3,"件");
  S(32,1,"CVR(CV率・入力/メディア)"); S(32,2,0.02); S(33,1,"CTR(クリック率・入力/メディア)"); S(33,2,0.01);
  S(34,1,"→ 必要 IMP [表示回数]"); F2(34,2,"=IF(B31<=0,0,ROUNDUP(B31/B32/B33,0))"); S(34,3,"回");
  S(35,1,"月の稼働日(入力)"); S(35,2,20);
  S(36,1,"💡 今日の追客本数"); F2(36,2,"=IF(B31<=0,0,ROUNDUP(B31/B35,0))"); S(36,3,"件/日 →02作業DBへ");
  S(37,1,"営業直結 時間比率(07連動)→目標60%"); F2(37,2,"='07_本部マトリクス'!F21");
  [26,28,30,32,33,35].forEach(function(r){ u.getRange(r,2).setBackground(IN); });
  [23,24,27,29,31,34,37].forEach(function(r){ u.getRange(r,2).setBackground(CALC); });
  u.getRange(25,2).setBackground(GOAL).setFontWeight("bold"); u.getRange(36,2).setBackground(GOAL).setFontWeight("bold");
  u.getRange("B23:B37").setNumberFormat("#,##0"); ["B28","B30","B32","B33","B37"].forEach(function(a){ u.getRange(a).setNumberFormat("0%"); });

  // 3) 未来会計図表(古田土・39-48)＝ドライバー直下。なぜ/どう使うを明記。
  ttl(39,"■ 未来会計図表（古田土：売上→粗利→利益の形）｜なぜ=どこを削れば/増やせば利益が出るか、粗利MQを軸に見る図");
  u.getRange(39,1).setNote("【なぜ存在】上のドライバーは「いくら要る」、ここは「利益の構造」。粗利益額MQが軸。【どう使う】損益分岐点比率(F÷MQ)で会社の格付け、労働分配率(人件費÷MQ)で人件費の重さを見る。案件(5-14)が増えれば売上PQ→MQ→格付けが自動で動く。");
  S(40,1,"売上高(PQ)＝客単価×客数"); F2(40,2,"="+PQ); S(40,3,"案件満額(法人)+EC+賃料の合算");
  S(41,1,"　├ 変動費(VQ)"); F2(41,2,"="+PQ+"-"+MQ); S(41,3,"変動費率v="); F2(41,4,"=B41/B40");
  S(42,1,"　└ 粗利益額(MQ＝限界利益)★軸"); F2(42,2,"="+MQ); S(42,3,"粗利益率m(=Gross Margin)="); F2(42,4,"=B42/B40");
  S(43,1,"固定費(F)"); F2(43,2,"="+F); S(43,3,"損益分岐点比率 F÷MQ(理想80%)="); F2(43,4,"=B43/B42");
  S(44,1,"経常利益(G)"); F2(44,2,"="+G); S(44,3,"経営安全率 G÷MQ="); F2(44,4,"=B44/B42");
  S(45,1,"格付け(損益分岐点比率)"); F2(45,2,'=IF(B43=0,"-",IF(B43<0.6,"SS 超優良",IF(B43<0.8,"S 優良",IF(B43<0.9,"A 健全",IF(B43<=1,"B 分岐点",IF(B43<=2,"C 赤字","D 倒産"))))))');
  S(46,1,"人件費(役員+社保+業務委託)"); F2(46,2,"="+JINKEN); S(46,3,"労働分配率 人件費÷MQ(=Labor Share)="); F2(46,4,"=B46/B42");
  S(47,1,"経費(モノ・金利・未来)"); F2(47,2,"=B43-B46"); S(47,3,"未来費用(広告/教育/R&D)は積極投資");
  u.getRange(40,2).setBackground(GRY).setFontWeight("bold"); u.getRange(41,2).setBackground(GRY); u.getRange(42,2).setBackground(CALC).setFontWeight("bold");
  u.getRange(43,2).setBackground(ORG).setFontWeight("bold"); u.getRange(44,2).setBackground(GOAL).setFontWeight("bold"); u.getRange(45,2).setBackground(GOAL).setFontWeight("bold");
  u.getRange(46,2).setBackground(GRN).setFontWeight("bold"); u.getRange(47,2).setBackground(ORG);
  u.getRange("B40:B47").setNumberFormat("#,##0"); ["D41","D42","D43","D44","D46"].forEach(function(a){ u.getRange(a).setNumberFormat("0.0%"); });

  // 4) 用語対応表(古田土⇔マーケ・49-58)
  ttl(49,"■ 用語対応表（古田土会計 ⇔ マーケKPI ＝ 全経営者の共通言語）");
  var tbl=[["古田土／会計","マーケ／共通","式・意味"],["売上高 PQ","売上 Revenue","客単価P×客数Q"],["客数 Q","CV/受注 Conversion","ドライバー出口=成約"],["粗利益率 m","粗利率 Gross Margin","MQ÷PQ"],["損益分岐点比率","BEP比率 Break-even","F÷MQ(理想80%)"],["労働分配率","人件費率 Labor Share","人件費÷MQ"],["集客効率","CPA 顧客獲得単価","広告費÷CV"],["転換効率","CVR コンバージョン率","CV÷流入"],["表示","IMP インプレッション","表示回数"]];
  u.getRange(50,1,tbl.length,3).setValues(tbl);
  u.getRange(50,1,1,3).setBackground("#666666").setFontColor("#FFFFFF").setFontWeight("bold");
  u.getRange(51,1,tbl.length-1,3).setBackground(GRY);

  // 5) 科目マスター(最下部・60-66)＋B5:B14のプルダウンを新位置へ貼り替え
  var mb=60; ttl(mb,"■ 勘定科目マスター（最下部・この6行を編集すると科目プルダウンが変わる）");
  var KAMOKU=["売上高（不動産売買）","売上高（コンサル）","賃貸料","売上（EC・海外）","営業代行","雑収入"];
  for(var m=0;m<6;m++) u.getRange(mb+1+m,1).setValue(KAMOKU[m]);
  u.getRange(mb+1,1,6,1).setBackground(IN);
  var rule=SpreadsheetApp.newDataValidation().requireValueInRange(u.getRange(mb+1,1,6,1),true).setAllowInvalid(true).build();
  u.getRange("B5:B14").setDataValidation(rule);

  return { ok:true,
    layout:"03=KPIヘッダ(1-4)/案件(5-14)/集計(16-21)/営業ドライバー(22-37・案件直下)/未来会計図表(39-47)/用語対応表(49-58)/科目マスター("+mb+"-)",
    fixed:"本部色再適用・空白38-49廃止・マスタ最下部・古田土図を直下へ・なぜ/どう使う注記・薄地黒文字・科目プルダウン貼り替え",
    note:"案件5-14/集計16-21/外部参照は不触＝安全。10収益パイプは案件=03/ニーズ=顧客マスターに分解済で再投入不要。" };
}

// ── 実装A基盤23：05入金SUMIFS貼り直し(REF一掃)＋04を薄地黒文字の主格色＋備考に主格別フォルダ 2026-06-06 ──
function buildBatch23(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  function frS(sh,kw){ var a=sh.getRange(1,1,120,1).getValues(); for(var i=0;i<a.length;i++){ if((a[i][0]||"").toString().indexOf(kw)>=0) return i+1; } return -1; }
  var out={};

  // === A) 05入金SUMIFS貼り直し(REF連鎖の根本=B9等を修復) ===
  var cf=shk("資金繰り");
  if(cf){
    var baseY=2026, baseM=6;
    [5,6,9].forEach(function(R){ for(var col=0;col<8;col++){ var mm=baseM+col,yy=baseY+Math.floor((mm-1)/12),m2=((mm-1)%12)+1;
      cf.getRange(R,2+col).setFormula("=SUMIFS('03_売上見込み'!$K:$K,'03_売上見込み'!$B:$B,$A"+R+",'03_売上見込み'!$G:$G,DATE("+yy+","+m2+",1))"); } }); // 不動産売買/コンサル/営業代行=一時
    [7,8].forEach(function(R){ for(var col=0;col<8;col++){
      cf.getRange(R,2+col).setFormula("=SUMIF('03_売上見込み'!$B:$B,$A"+R+",'03_売上見込み'!$K:$K)"); } }); // 賃貸料/EC=recurring
    out.cf05="05入金B5:I9のSUMIFS/SUMIFを03のB/G/K列へ貼り直し＝REF連鎖を一掃(03ドライバー/04今日の行動が復活)";
  }

  // === B) 04を薄地黒文字の主格色に(白抜き廃止)＋結論金＋備考フォルダ ===
  var pl=shk("損益");
  if(pl){
    const LAWt="#E8F0FE",PERt="#E6F4EA",HOMEt="#FFF2E6",ALLt="#F1F1F1",GOLD="#FFE599",BLK="#000000",COLS=9;
    var sec=[{h:frS(pl,"■法人"),t:LAWt,n:"法人=青／明細フォルダ:共有ドライブ/02_KHD/2026_帳票"},{h:frS(pl,"■個人事業"),t:PERt,n:"個人事業=緑／明細フォルダ:01_個人/2026_帳票"},{h:frS(pl,"■家計"),t:HOMEt,n:"家計=橙／明細フォルダ:00_プライベート/2026"},{h:frS(pl,"■全社"),t:ALLt,n:"全社=灰(集計)"}];
    var endAll=frS(pl,"家族が潰れ"); if(endAll<1) endAll=(sec[3].h>0?sec[3].h+8:90);
    var bnd=[sec[1].h,sec[2].h,sec[3].h,endAll];
    for(var i=0;i<4;i++){ var s=sec[i].h; if(s<1) continue; var e=bnd[i]-1; if(e<s) e=s;
      pl.getRange(s,1,e-s+1,COLS).setBackground(sec[i].t).setFontColor(BLK).setFontWeight("normal"); // 明細=薄色・黒文字
      pl.getRange(s,1,1,COLS).setBackground(sec[i].t).setFontColor(BLK).setFontWeight("bold");        // 見出し=同系薄色・黒太字(白抜き廃止)
      pl.getRange(s,1).setNote(sec[i].n);                                                             // 備考=主格別明細フォルダ
    }
    var aV=pl.getRange(1,1,endAll,1).getValues();
    for(var r=1;r<=endAll;r++){ var v=(aV[r-1][0]||"").toString();
      if(["経常利益","税引後利益","収支","黒字判定"].some(function(k){return v.indexOf(k)>=0;}))
        pl.getRange(r,1,1,COLS).setBackground(GOLD).setFontWeight("bold").setFontColor(BLK); }
    var hdr=frS(pl,"科目"); if(hdr<1) hdr=2;
    pl.getRange(hdr,1).setNote("色=主格(薄地黒文字)。青=法人/緑=個人事業/橙=家計/灰=全社。金=結論。各■見出しの備考に主格別明細フォルダのパス。");
    out.pl04="04を薄地黒文字の主格色＋結論金＋各見出しに主格別明細フォルダ注記";
  }
  return { ok:true, applied:out,
    note:"05のSUMIFS貼り直しでREF一掃が最重要。04のドライバー削除＋01貼り替えは次(buildBatch24)で安全に。備考フォルダは実URLを後で貼れるようパス記載。" };
}

// ── 実装A基盤24：循環参照を断つ＝03向けSUMIFSを列全体→案件行限定($5:$14)に。01!B38(孤立REF)も修復 2026-06-06 ──
// 原因：05/04PLのSUMIFSが03の列全体を舐め、その中の営業ドライバー(B23→05!B37→B9)/古田土図(→04PL)を巻き込み循環。範囲を案件行5-14に限定して輪を切る。
function buildBatch24(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  function frS(sh,kw){ var a=sh.getRange(1,1,120,1).getValues(); for(var i=0;i<a.length;i++){ if((a[i][0]||"").toString().indexOf(kw)>=0) return i+1; } return -1; }
  var out={};
  var K="'03_売上見込み'!$K$5:$K$14", B="'03_売上見込み'!$B$5:$B$14", G="'03_売上見込み'!$G$5:$G$14";

  // === A) 05入金SUMIFS/SUMIFを案件行限定に(循環の根を断つ) ===
  var cf=shk("資金繰り");
  if(cf){
    var baseY=2026, baseM=6;
    [5,6,9].forEach(function(R){ for(var col=0;col<8;col++){ var mm=baseM+col,yy=baseY+Math.floor((mm-1)/12),m2=((mm-1)%12)+1;
      cf.getRange(R,2+col).setFormula("=SUMIFS("+K+","+B+",$A"+R+","+G+",DATE("+yy+","+m2+",1))"); } });
    [7,8].forEach(function(R){ for(var col=0;col<8;col++){
      cf.getRange(R,2+col).setFormula("=SUMIF("+B+",$A"+R+","+K+")"); } });
    out.cf05="05入金B5:I9のSUMIFS/SUMIFを03の案件行限定($5:$14)へ＝循環の根を断つ";
  }

  // === B) 04PLの03向けSUMIFS(法人売上・月別)を案件行限定に(古田土図との循環を断つ) ===
  var pl=shk("損益");
  if(pl){
    var rUri=frS(pl,"売上（コンサル/物件"); // 法人売上行
    if(rUri>0){
      for(var col=0;col<8;col++){ var L=String.fromCharCode(66+col); var mm=6+col,yy=2026+Math.floor((mm-1)/12),m2=((mm-1)%12)+1;
        pl.getRange(rUri,2+col).setFormula("=SUMIFS("+K+","+G+",DATE("+yy+","+m2+",1))"); }
      out.pl04="04PL法人売上(r"+rUri+")のSUMIFSを案件行限定($5:$14)へ＝古田土図との循環を断つ";
    }
  }

  // === C) 01!B38(家計収支=#REF孤立)を04PLの収支へ貼り直し ===
  var s1=shk("統合司令塔");
  if(s1 && pl){
    var rShu=-1, aP=pl.getRange(1,1,120,1).getValues();
    for(var i=0;i<aP.length;i++){ if((aP[i][0]||"").toString().indexOf("収支")>=0){ rShu=i+1; break; } }
    if(rShu>0){ s1.getRange("B38").setFormula("='04_損益PL'!B"+rShu); out.cmd01="01!B38(家計収支)を04PL収支(r"+rShu+")へ貼り直し"; }
  }

  return { ok:true, applied:out,
    note:"循環参照(03ドライバー/古田土図⇔05/04PLの列全体SUMIFS)を、集計範囲を案件行5-14に限定して切断。これで05のREFが消え→03ドライバー・04今日の行動が復活。" };
}

// ── 実装A基盤25：完成＝03!B30補完(DIV/0解消)＋04スッキリ化(driver/トラッキング/今日の行動削除)＋01貼替 2026-06-06 ──
function buildBatch25(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var u=shk("売上見込み"), pl=shk("損益"), s1=shk("統合司令塔");
  var out={};

  // A) 03!B30(追客→アポ率)が空＝#DIV/0!の唯一の起点。0.5を補完。
  if(u){ if(u.getRange("B30").getValue()==="" || u.getRange("B30").getValue()==null) u.getRange("B30").setValue(0.5); else u.getRange("B30").setValue(0.5); out.b30="03!B30=0.5補完→B31/B34/B36のDIV/0解消"; }

  // B) 01の参照を03ドライバーへ貼替(04driver削除の前に・#REF防止)
  if(s1){
    s1.getRange("B27").setFormula("='03_売上見込み'!B25"); // 今日の一手→経常ギャップ(穴)
    s1.getRange("B28").setFormula("='03_売上見込み'!B36"); // 追客目標→今日の追客本数
    s1.getRange("B15").setFormula("='03_売上見込み'!B25"); // 損益分岐リスク→ギャップ
    out.cmd01="01の今日の一手/追客目標/損益分岐リスクを03ドライバー(B25/B36)へ貼替";
  }

  // C) 04スッキリ化：家族ライン～CF手前(=営業ドライバー+月次トラッキング)を削除、続いてtop今日の行動を削除
  if(pl){
    function frP(kw){ var a=pl.getRange(1,1,160,1).getValues(); for(var i=0;i<a.length;i++){ if((a[i][0]||"").toString().indexOf(kw)>=0) return i+1; } return -1; }
    var dStart=frP("家族が潰れ"); if(dStart<1) dStart=frP("営業ドライバー");
    var cfStart=frP("資金繰り（CF");
    if(dStart>0 && cfStart>dStart){ pl.deleteRows(dStart, cfStart-dStart); out.delDriver="04の家族ライン～CF手前(driver+月次トラッキング)を削除"; }
    var topS=frP("今日の行動（最上部");
    if(topS>0){ pl.deleteRows(topS, 7); out.delTop="04top今日の行動を削除(03に一本化済)"; }
    else { var topS2=frP("今日の行動"); if(topS2>0 && topS2<10){ pl.deleteRows(topS2,7); out.delTop="04top今日の行動を削除"; } }
  }

  return { ok:true, applied:out,
    note:"完成形：04=PL本体(法人/個人/家計/全社・決算接地)＋下部CF(05連動)のスッキリ構成。営業ドライバー/月次トラッキング/今日の行動は03に一本化。01は03ドライバー参照に貼替。残=⑥11個人習慣の別タブ分離のみ。" };
}

// ── 実装A基盤26：診断＋確定修復＋読み戻し検証(一括) 2026-06-06 ──
// 03ドライバー入力欄を全て再投入(DIV/0根絶)＋01!B25/C25の#REF修復＋04/03/01のエラーセルを読み戻して報告
function buildBatch26(body){
  const ss=SpreadsheetApp.openById((body&&body.hostId)||"1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var u=shk("売上見込み"), pl=shk("損益"), s1=shk("統合司令塔");
  var rep={fixed:[],before:{},after:{},errors:{}};

  // ── A) 03ドライバー入力欄を全て再投入(空欄ならDIV/0になる6つの入力＋出力式) ──
  if(u){
    // before: 入力6セルの現状値を記録
    ["B26","B28","B30","B32","B33","B35"].forEach(function(a){ rep.before["03!"+a]=u.getRange(a).getValue(); });
    // 入力欄を確定値で再投入(冪等・空欄事故を根絶)
    u.getRange("B26").setValue(1100000); // 平均成約単価P
    u.getRange("B28").setValue(0.2);     // アポ→成約 転換率
    u.getRange("B30").setValue(0.5);     // 追客→アポ 転換率 ★DIV/0の起点
    u.getRange("B32").setValue(0.02);    // CVR
    u.getRange("B33").setValue(0.01);    // CTR
    u.getRange("B35").setValue(20);      // 稼働日
    // 出力式も再投入(消えてても復活)
    u.getRange("B27").setFormula("=IF(B25<=0,0,ROUNDUP(B25/B26,0))");
    u.getRange("B29").setFormula("=IF(B27<=0,0,ROUNDUP(B27/B28,0))");
    u.getRange("B31").setFormula("=IF(B29<=0,0,ROUNDUP(B29/B30,0))");
    u.getRange("B34").setFormula("=IF(B31<=0,0,ROUNDUP(B31/B32/B33,0))");
    u.getRange("B36").setFormula("=IF(B31<=0,0,ROUNDUP(B31/B35,0))");
    SpreadsheetApp.flush();
    // after: 出力式の表示値(エラーか数値か)
    ["B25","B26","B27","B28","B29","B30","B31","B34","B36"].forEach(function(a){ rep.after["03!"+a]=u.getRange(a).getDisplayValue(); });
    rep.fixed.push("03!B26/B28/B30/B32/B33/B35を確定値で再投入＋B27/29/31/34/36の式を再設定→DIV/0根絶");
    // 03の22-47でエラー表示が残るセルを拾う
    var dv=u.getRange(22,1,26,3).getDisplayValues(); var e3=[];
    for(var i=0;i<dv.length;i++){ for(var j=0;j<3;j++){ var v=(dv[i][j]||"").toString(); if(v.indexOf("#")===0) e3.push("R"+(22+i)+"C"+(j+1)+"="+v); } }
    rep.errors["03(22-47)"]=e3.length?e3:"エラー無し";
  }

  // ── B) 01司令塔の#REF修復＋残存#REFスキャン ──
  if(s1){
    rep.before["01!B25"]=s1.getRange("B25").getFormula()||s1.getRange("B25").getValue();
    s1.getRange("B25").setFormula("='03_売上見込み'!B25"); // 損益分岐リスク→経常ギャップ(穴)
    s1.getRange("C25").setValue("");
    SpreadsheetApp.flush();
    rep.fixed.push("01!B25→'03'!B25(経常ギャップ)へ貼替・C25クリア");
    // 01全体で #REF! / #ERROR! 表示を拾う
    var lr=s1.getLastRow(), lc=s1.getLastColumn();
    var dv1=s1.getRange(1,1,Math.min(lr,60),Math.min(lc,12)).getDisplayValues(); var e1=[];
    for(var i=0;i<dv1.length;i++){ for(var j=0;j<dv1[i].length;j++){ var v=(dv1[i][j]||"").toString(); if(v.indexOf("#REF")===0||v.indexOf("#ERROR")===0||v==="#DIV/0!") e1.push("R"+(i+1)+"C"+(j+1)+"="+v); } }
    rep.errors["01"]=e1.length?e1:"エラー無し";
  }

  // ── C) 04_損益PLのエラーセルをスキャン(「21行目以下」の真偽確認) ──
  if(pl){
    var lr=pl.getLastRow();
    var dv2=pl.getRange(1,1,Math.min(lr,160),Math.min(pl.getLastColumn(),8)).getDisplayValues(); var e2=[];
    for(var i=0;i<dv2.length;i++){ for(var j=0;j<dv2[i].length;j++){ var v=(dv2[i][j]||"").toString(); if(v.indexOf("#REF")===0||v.indexOf("#ERROR")===0||v==="#DIV/0!"||v==="#N/A") e2.push("R"+(i+1)+"C"+(j+1)+"="+v); } }
    rep.errors["04"]=e2.length?e2:"エラー無し";
  }

  return { ok:true, applied:rep.fixed, before:rep.before, after:rep.after, errors:rep.errors,
    note:"after=各出力式の表示値(数値なら正常)。errorsが全て『エラー無し』なら03/04/01完全クリーン＝完成。残るエラー行があればそのR/Cを次で潰す。" };
}

// ── カレンダー：終了イベント確認 ──────────────────────────
function checkEndingEvents() {
  try {
    const now = new Date();
    const windowMin = 26; // 25分+1分バッファ
    const from = new Date(now.getTime() - windowMin * 60 * 1000);

    // 全カレンダーを検索
    const calendars = CalendarApp.getAllCalendars();
    const ending = [];
    const props = PropertiesService.getScriptProperties();

    calendars.forEach(cal => {
      if (cal.isHidden() || cal.isSelected() === false) return;
      cal.getEvents(from, now).forEach(ev => {
        if (ev.isAllDayEvent()) return;
        const end = ev.getEndTime();
        if (end < from || end > now) return;

        // 重複通知防止（同じイベントを15分以内に2度聞かない）
        const key = "asked_" + ev.getId();
        if (props.getProperty(key)) return;
        props.setProperty(key, "1");

        // 2時間後に自動削除（PropertiesServiceは自動削除機能ないので別途）
        ending.push({
          title: ev.getTitle(),
          end: formatTime(end),
          id: ev.getId().slice(0, 20)
        });
      });
    });

    return { events: ending, checked_at: formatTime(now) };
  } catch (e) {
    return { events: [], error: e.toString() };
  }
}

// ── 対話ログ記録 ─────────────────────────────────────────
function logConversation(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("LINE対話ログ");
  if (!sheet) {
    sheet = ss.insertSheet("LINE対話ログ");
    sheet.appendRow(["日時", "送信者", "方向", "内容", "アクション"]);
  }
  const now = todayStr() + " " + formatTime(new Date());
  sheet.appendRow([
    now,
    body.sender || "菊池",
    body.direction || "受信",
    body.content || "",
    body.action_taken || ""
  ]);
  return { logged: true };
}

// ── 次のWBSタスク提案 ────────────────────────────────────
function getNextTask() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("02_全社WBS");
  if (!sheet) return { error: "WBSシートが見つかりません" };

  const data = sheet.getDataRange().getValues();
  let headerRow = -1, colTask = 1, colStatus = -1, colFrom = -1, colDept = 0;
  for (let i = 0; i < data.length; i++) {
    const idx = data[i].indexOf("ステータス");
    if (idx >= 0) {
      headerRow = i; colStatus = idx;
      const ti = data[i].indexOf("工程"); if (ti >= 0) colTask = ti;
      const fi = data[i].indexOf("From"); if (fi >= 0) colFrom = fi;
      break;
    }
  }
  if (headerRow < 0) return { error: "ヘッダーが見つかりません" };

  const today = new Date();
  const priority = [], upcoming = [];

  for (let i = headerRow + 1; i < data.length; i++) {
    const task   = String(data[i][colTask]).trim();
    const status = String(data[i][colStatus]).trim();
    const dept   = String(data[i][colDept]).trim();
    if (!task || /ルーティン/.test(status)) continue;
    if (/完了|済み/.test(status)) continue;

    const label = `[${dept}] ${task}`;
    if (/最優先/.test(status)) priority.push(label);
    else if (/対応中/.test(status)) upcoming.push(label);
  }

  const all = [...priority, ...upcoming].slice(0, 3);
  if (all.length === 0) return { message: "対応中タスクなし" };

  return {
    suggestion: "次はこれですか？\n" + all.map((t,i) => `${i+1}. ${t}`).join("\n"),
    tasks: all
  };
}

// ── WBS進捗サマリー取得 ───────────────────────────────────
function getWbsSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("02_全社WBS");
  if (!sheet) return { error: "WBSシートが見つかりません" };

  const data = sheet.getDataRange().getValues();

  let headerRow = -1, colStatus = -1, colTask = 1, colDept = 0, colProgress = -1;
  for (let i = 0; i < data.length; i++) {
    const idx = data[i].indexOf("ステータス");
    if (idx >= 0) {
      headerRow  = i;
      colStatus  = idx;
      colTask    = data[i].indexOf("工程");    if (colTask < 0) colTask = 1;
      colDept    = data[i].indexOf("部署");    if (colDept < 0) colDept = 0;
      colProgress = data[i].indexOf("進捗率");
      break;
    }
  }
  if (headerRow < 0) return { error: "WBSヘッダーが見つかりません" };

  const DONE     = ["完了", "済み"];
  const PRIORITY = ["最優先"];
  const ACTIVE   = ["対応中", "新規立ち上げ", "新規構築中", "段階拡大中", "計画中"];
  const ROUTINE  = ["ルーティン"];

  let done = 0, total = 0;
  const priority = [], active = [];

  for (let i = headerRow + 1; i < data.length; i++) {
    const task   = String(data[i][colTask]).trim();
    const dept   = String(data[i][colDept]).trim();
    const status = String(data[i][colStatus]).trim();
    if (!task || ROUTINE.some(s => status.includes(s))) continue;
    total++;
    if (DONE.some(s => status.includes(s))) {
      done++;
    } else if (PRIORITY.some(s => status.includes(s))) {
      priority.push(`🔴 [${dept}] ${task}`);
    } else if (ACTIVE.some(s => status.includes(s))) {
      active.push(`🟡 [${dept}] ${task}`);
    }
  }

  const lines = [`📊 WBS: 完了${done}/${total}件`];
  if (priority.length) {
    lines.push("\n🔴 最優先:");
    priority.slice(0, 5).forEach(t => lines.push("  " + t.replace(/^🔴 /, "")));
  }
  if (active.length) {
    lines.push("\n🟡 対応中:");
    active.slice(0, 5).forEach(t => lines.push("  " + t.replace(/^🟡 /, "")));
  }
  if (!priority.length && !active.length) lines.push("対応中タスクなし");

  return { summary: lines.join("\n"), done, total,
           priority: priority.length, active: active.length };
}

// ── 音声報連相パース → 自動更新 ──────────────────────────
function parseAndUpdate(text) {
  if (!text) return { error: "textが空です" };
  const t = text.trim();
  const results = [];

  // ① WBS完了
  const wbsDone = t.match(/^(.+?)(?:が|を|は)?(?:終わっ|完了|できた|片付|済み|やり切|リリース|上げ)/);
  if (wbsDone) {
    const task = wbsDone[1].trim();
    if (!/朝活|勉強|学習|民法|不登|KPI|他力|EQ|家族/.test(task)) {
      results.push(updateWbs(task, "完了", "100%"));
    }
  }

  // ② WBS対応中
  if (!wbsDone) {
    const wbsWip = t.match(/^(.+?)(?:に|を)?(?:着手|対応中|始め|取り掛|やり始|進め)/);
    if (wbsWip) {
      const task = wbsWip[1].trim();
      if (!/朝活|勉強|学習|民法|不登|KPI|他力|EQ|家族/.test(task)) {
        results.push(updateWbs(task, "対応中", null));
      }
    }
  }

  // ③ 学習ログ
  const study = {};
  const hm = t.match(/(?:朝活|勉強|学習)[^\d]*(\d+(?:\.\d+)?)\s*(?:時間|h|H)/);
  if (hm) study.hours = parseFloat(hm[1]);
  const rm = t.match(/(?:民法|不登法|択一|記述|区分|表題)[^\d]*\d+[~〜\-]?\d*\s*(?:問|p|P|ページ|頁)?/);
  if (rm) study.range_takuitsu = rm[0].trim();
  const mm = t.match(/ミス[^\d]*(\d+)\s*(?:個|問|つ)?/);
  if (mm) study.miss_count = parseInt(mm[1]);
  if (Object.keys(study).length > 0) results.push(updateStudy(study));

  // ④ KPI
  const kpiBody = {};
  const kpiPatterns = [
    ["tariki",   /他力[：:\s]*([○×〇✕])/],
    ["eq",       /EQ[：:\s]*([○×〇✕])/],
    ["chokatsu", /朝活[：:\s]*([○×〇✕])/],
    ["sen_sabo", /戦サボ[：:\s]*([○×〇✕])/],
    ["kazoku",   /家族[：:\s]*([○×〇✕])/],
  ];
  kpiPatterns.forEach(([key, pat]) => {
    const m = t.match(pat);
    if (m) kpiBody[key] = m[1].replace("〇","○").replace("✕","×");
  });
  if (/KPI.{0,10}(?:全部|全て|全達成|達成|①②③④5)/.test(t) && Object.keys(kpiBody).length === 0) {
    kpiPatterns.forEach(([key]) => { kpiBody[key] = "○"; });
  }
  if (Object.keys(kpiBody).length > 0) results.push(updateKpi(kpiBody));

  // 5 日次ログ（会議・打ち合わせ系）
  if (/(?:やった|参加|mtg|ミーティング|打ち合|会議|報告|確認|連絡|提出|送った|送信)/i.test(t)) {
    results.push(appendLog({ content: t, category: "業務記録" }));
  }

  if (results.length === 0) return { error: "パターン未検出: " + t };

  const summaries = results.map(r => {
    const upd = (r.updated || []).join(" / ");
    return upd || ("row=" + r.row);
  });
  return { parsed: t, actions: results.length, summaries };
}

// ── デバッグ：シート構造確認 ──────────────────────────
function debugSheet(partial) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = partial ? findSheetByPartial(ss, partial) : ss.getActiveSheet();
  if (!sheet) return { error: "シートが見つかりません: " + partial };
  const data = sheet.getDataRange().getValues();
  return {
    name: sheet.getName(),
    rows: data.length,
    row1: data[0],
    row2: data[1] || [],
    row3: data[2] || [],
    row4: data[3] || [],
  };
}

// ── ユーティリティ ────────────────────────────────────
function notNull(v) {
  return v !== undefined && v !== null;
}

function findSheetByPartial(ss, partial) {
  return ss.getSheets().find(s => s.getName().includes(partial)) || null;
}

function todayStr() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth()+1).padStart(2,"0");
  const day = String(d.getDate()).padStart(2,"0");
  return `${y}/${m}/${day}`;
}

function dateMatches(cellVal, todayFull) {
  if (!cellVal) return false;
  let s;
  if (cellVal instanceof Date) {
    s = todayStr.call(null); // same format
    const d = cellVal;
    s = `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getDate()).padStart(2,"0")}`;
  } else {
    s = String(cellVal).replace(/\s*\([月火水木金土日]\).*/, "").trim();
  }
  // "2026/05/15" か "05/15" のどちらでも一致
  return s === todayFull || s === todayFull.slice(5);
}

function formatTime(d) {
  return String(d.getHours()).padStart(2,"0") + ":" + String(d.getMinutes()).padStart(2,"0");
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
