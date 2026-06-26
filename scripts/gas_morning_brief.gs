/**
 * KHD 朝ブリーフ 自動通知（毎朝6:30）
 * ------------------------------------------------------------
 * ①統合司令塔（今日の優先順位/一手/追客/時間/損益分岐/谷）＋
 * 🎯追客リスト（今日の推奨追客）＋ 今日のカレンダー予定 を
 * 1通のメールにまとめて ec-seller@kikuchi-hd.net へ毎朝送る。
 *
 * 【設置】
 *  1. 本体スプシ(2026_KHD PJ一覧_v2)で 拡張機能>Apps Script
 *  2. このコードを貼り付けて保存
 *  3. 拡張機能>Apps Script>サービス(＋) で「Tasks API」を追加（追客の自動タスク化に必要）
 *  4. 関数 setupMorningTrigger を1回だけ実行（毎朝6:30トリガー登録）
 *  5. 初回は権限承認（スプレッドシート/カレンダー/メール/タスク）
 *  ※テスト送信は sendMorningBrief を手動実行
 */

var SID = "1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc";
var MAIL_TO = "ec-seller@kikuchi-hd.net";

/** 毎朝6:30トリガーを登録（1回だけ実行） */
function setupMorningTrigger() {
  // 既存の同名トリガーを掃除
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "sendMorningBrief") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sendMorningBrief").timeBased().atHour(6).nearMinute(30).everyDays(1).create();
  Logger.log("✅ 毎朝6:30トリガー登録完了");
}

/** ラベル(A列)に keyword を含む行の C列値を返す */
function findByLabel(rows, keyword) {
  for (var i = 0; i < rows.length; i++) {
    var a = String(rows[i][0] || "");
    if (a.indexOf(keyword) !== -1) return String(rows[i][2] || rows[i][1] || "");
  }
  return "";
}

/** 朝ブリーフ本文を生成して送信 */
function sendMorningBrief() {
  var ss = SpreadsheetApp.openById(SID);
  var tz = ss.getSpreadsheetTimeZone() || "Asia/Tokyo";
  var today = Utilities.formatDate(new Date(), tz, "yyyy/MM/dd (E)");

  // --- ①統合司令塔 を読む（行ズレに強いラベル検索）---
  var cmd = ss.getSheetByName("①統合司令塔");
  var rows = cmd ? cmd.getRange("A1:C40").getValues() : [];
  var ranway   = findByLabel(rows, "ランウェイ");
  var tani     = findByLabel(rows, "谷");
  var bep      = findByLabel(rows, "損益分岐リスク");
  var ritsu    = findByLabel(rows, "営業直結比率");
  var hitote   = findByLabel(rows, "今日の一手");
  var tsuikyaku= findByLabel(rows, "追客目標");
  var jikan    = findByLabel(rows, "時間の一手");

  // 優先順位①〜⑤
  var prios = [];
  rows.forEach(function (r) {
    var a = String(r[0] || "");
    if (/^[①②③④⑤]/.test(a)) prios.push("  " + a + (r[2] ? "  … " + r[2] : ""));
  });

  // --- 🎯追客リスト 今日の推奨 ---
  var tl = ss.getSheetByName("🎯追客リスト");
  var pick = tl ? String(tl.getRange("C3").getValue() || "") : "";

  // --- 今日のカレンダー予定 ---
  var evs = [];
  try {
    CalendarApp.getDefaultCalendar().getEventsForDay(new Date()).forEach(function (e) {
      var t = e.isAllDayEvent() ? "終日" : Utilities.formatDate(e.getStartTime(), tz, "HH:mm");
      evs.push("  " + t + " " + e.getTitle());
    });
  } catch (err) { evs.push("  (カレンダー取得不可)"); }

  // --- 本文 ---
  var b = [];
  b.push("☀️ KHD 朝ブリーフ  " + today);
  b.push("");
  b.push("【今日の予定】");
  b.push(evs.length ? evs.join("\n") : "  予定なし");
  b.push("");
  b.push("【📌 今日の優先順位（この順に・考えない）】");
  b.push(prios.length ? prios.join("\n") : "  (司令塔参照)");
  b.push("");
  b.push("🎯 今日の推奨追客 → " + (pick || "(追客リスト参照)"));
  b.push("🎯 追客目標 → " + tsuikyaku);
  b.push("💡 今日の一手 → " + hitote);
  b.push("⏱ 時間の一手 → " + jikan + "（営業直結 現状 " + ritsu + " / 目標60%を死守）");
  b.push("");
  b.push("【お金のリスク】");
  b.push("  🚨 " + bep);
  b.push("  通期の谷: " + tani + "（ランウェイ " + ranway + "）");
  b.push("");
  b.push("— 数字は『自分の動く量』の羅針盤。客には信頼でGIVE。緊急以外は冷静に順に処理。");
  b.push("司令塔: https://docs.google.com/spreadsheets/d/" + SID + "/edit#gid=214595378");

  // 今日やる追客 上位3件を Google Tasks に自動投入
  var taskMsg = createTodayTasks(ss);
  if (taskMsg) b.push(""), b.push("【Google Tasksに投入した今日の追客】"), b.push(taskMsg);

  MailApp.sendEmail(MAIL_TO, "☀️ KHD朝ブリーフ " + today, b.join("\n"));
  Logger.log("✅ 朝ブリーフ送信: " + MAIL_TO);
}

/** 追客リストの優先スコア上位3件(進行中)をGoogle Tasksに投入し、本文用の要約を返す */
function createTodayTasks(ss) {
  var tl = ss.getSheetByName("🎯追客リスト");
  if (!tl) return "";
  var rows = tl.getRange("A6:K35").getValues();
  // [A優先スコア,B本部,C対象名,D種類,E確度,F温度,G次アク,H期限,I最終,J状態,K出所]
  var cand = rows.filter(function (r) {
    return r[2] && String(r[9]).indexOf("完了") === -1 && String(r[9]).indexOf("客付済") === -1;
  });
  cand.sort(function (a, b) { return (Number(b[0]) || 0) - (Number(a[0]) || 0); });
  var top = cand.slice(0, 3);
  var lines = [];
  var due = new Date(); due.setHours(23, 0, 0, 0);
  top.forEach(function (r) {
    var title = "🎯追客: " + r[2];
    var notes = String(r[6] || "") + "（確度" + r[4] + "% / " + r[1] + " / スコア" + r[0] + "）";
    try {
      Tasks.Tasks.insert({ title: title, notes: notes, due: due.toISOString() }, "@default");
      lines.push("  ・" + r[2] + " → " + r[6]);
    } catch (e) {
      lines.push("  ・" + r[2] + "（Tasks API未有効?: " + e + "）");
    }
  });
  return lines.join("\n");
}
