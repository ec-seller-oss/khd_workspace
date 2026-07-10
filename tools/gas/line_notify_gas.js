// ============================================================
// 毎日朝5時 WBS未完了タスク → LINE通知スクリプト
// Google Apps Script (GAS) にコピー&ペーストして使用
// ============================================================

var CONFIG = {
  SPREADSHEET_ID: "1OLNAcbPIMpdHFr8b-nYQosVuOKzzhXuuGDPsEwSk2i4",
  SHEET_GID: 1170356369,
  LINE_NOTIFY_TOKEN: "YOUR_LINE_NOTIFY_TOKEN_HERE", // ← ここを差し替える
};

// 本部長ごとの戦略・役割定義
var HONBUCHO_META = {
  "麻梨奈": {
    code: "00",
    emoji: "👶",
    area: "プライベート・ファミリー",
    strategy: "家族基盤を整えることで経営者メンタルを安定化。子育て・健康・ゴルフで心身のリフレッシュを確保する。",
    dailyFocus: "育児タスクは朝のうちに情報収集。長期計画は週1回30分で前進させる。",
  },
  "菊池": {
    code: "01",
    emoji: "📊",
    area: "経営管理",
    strategy: "MF記帳内製化で月次コスト削減、予実管理の見える化により意思決定スピードを上げる。株主総会・花巻支店登記で法的基盤を整備。",
    dailyFocus: "毎朝数字を確認し、異常値に即対応。管理業務は午前中に集中処理。",
  },
  "橋本": {
    code: "02",
    emoji: "💰",
    area: "資金調達",
    strategy: "AGP解約（6月末期限）とGWS確認（5月末期限）が最優先。運転資金1000万・岩銀融資を並行で進め、不要カード解約でコスト圧縮。キャッシュポジション強化が最重要ミッション。",
    dailyFocus: "期限付きタスクを毎日1件以上前進させる。金融機関との連絡は午前中に完結。",
  },
  "江藤": {
    code: "03",
    emoji: "🏢",
    area: "事業運営",
    strategy: "バイセル業務をルーティン化してオペレーションを安定させ、売上の底上げを図る。",
    dailyFocus: "ルーティン業務の品質維持。例外対応は記録して横展開。",
  },
  "福井": {
    code: "04",
    emoji: "📚",
    area: "コンサル・調査士",
    strategy: "2027年土地家屋調査士合格を目標に朝活3h死守。合格によりコンサル価値と案件単価を大幅向上させる。",
    dailyFocus: "朝活3時間は最優先。予定が入っても学習時間を死守する。",
  },
  "川股": {
    code: "05",
    emoji: "🏠",
    area: "物件調達",
    strategy: "追客・物件スクリーニングを継続的に実施し、パイプラインを太く維持。良質物件の早期確保が収益に直結。",
    dailyFocus: "毎日物件情報を確認し、有望案件は即アクション。追客リストを週次で更新。",
  },
};

// ステータスの優先順位（数値が小さいほど上に表示）
var STATUS_ORDER = {
  "🔴最優先": 0,
  "最優先":   1,
  "未着手":   2,
  "対応中":   3,
  "計画中":   4,
  "ルーティン": 5,
};

// メイン関数（トリガーで実行）
function sendDailyReport() {
  var tasks = extractUnfinishedTasks();
  var message = buildMessage(tasks);
  sendToLine(message);
}

// スプレッドシートから未完了タスクを取得
function extractUnfinishedTasks() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // gidからシートを特定
  var sheets = ss.getSheets();
  var sheet = null;
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === CONFIG.SHEET_GID) {
      sheet = sheets[i];
      break;
    }
  }
  if (!sheet) {
    // フォールバック: アクティブシート
    sheet = ss.getActiveSheet();
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  // ヘッダー行からカラムインデックスを特定
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var colMap = {
    task:     findCol(headers, ["工程", "タスク", "task"]),
    aim:      findCol(headers, ["狙い・完了条件", "完了条件", "aim"]),
    honbucho: findCol(headers, ["担当本部長", "本部長", "honbucho"]),
    person:   findCol(headers, ["担当者", "person"]),
    from:     findCol(headers, ["From", "FROM", "from", "開始"]),
    to:       findCol(headers, ["TO", "To", "to", "期限", "終了"]),
    progress: findCol(headers, ["進捗率", "進捗", "progress"]),
    status:   findCol(headers, ["ステータス", "status"]),
  };

  var result = {};

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var status = colMap.status >= 0 ? String(row[colMap.status]).trim() : "";
    var taskName = colMap.task >= 0 ? String(row[colMap.task]).trim() : "";

    // 空行・完了タスクはスキップ
    if (!taskName || taskName === "" || status === "完了") continue;

    var honbucho = colMap.honbucho >= 0 ? String(row[colMap.honbucho]).trim() : "未割当";
    if (!honbucho || honbucho === "") honbucho = "未割当";

    var progress = colMap.progress >= 0 ? String(row[colMap.progress]).trim() : "0%";
    var to = colMap.to >= 0 ? formatDate(row[colMap.to]) : "";
    var aim = colMap.aim >= 0 ? String(row[colMap.aim]).trim() : "";

    if (!result[honbucho]) result[honbucho] = [];
    result[honbucho].push({
      task: taskName,
      status: status || "未着手",
      progress: progress,
      deadline: to,
      aim: aim,
    });
  }

  // 各本部長のタスクをステータス順にソート
  Object.keys(result).forEach(function(name) {
    result[name].sort(function(a, b) {
      var oa = STATUS_ORDER[a.status] !== undefined ? STATUS_ORDER[a.status] : 99;
      var ob = STATUS_ORDER[b.status] !== undefined ? STATUS_ORDER[b.status] : 99;
      return oa - ob;
    });
  });

  return result;
}

// LINEメッセージを生成
function buildMessage(tasksByHonbucho) {
  var today = Utilities.formatDate(new Date(), "Asia/Tokyo", "MM/dd(E)");
  var lines = [];

  lines.push("━━━━━━━━━━━━━━━━━");
  lines.push("🌅 " + today + " 本部長WBS朝報");
  lines.push("━━━━━━━━━━━━━━━━━");

  // 定義順に出力
  var orderedNames = ["麻梨奈", "菊池", "橋本", "江藤", "福井", "川股"];

  orderedNames.forEach(function(name) {
    var meta = HONBUCHO_META[name];
    if (!meta) return;

    var tasks = tasksByHonbucho[name] || [];
    // 完了タスクのみ → スキップ (extractUnfinishedTasksで除外済み)
    // タスクが0件でも戦略は表示（継続的ミッション確認のため）

    lines.push("");
    lines.push(meta.emoji + "【" + meta.code + "_" + name + "】" + meta.area);
    lines.push("🎯 " + meta.strategy);
    lines.push("📌 今日のフォーカス: " + meta.dailyFocus);

    if (tasks.length > 0) {
      lines.push("");
      lines.push("📋 未完了タスク(" + tasks.length + "件):");
      tasks.forEach(function(t) {
        var flag = t.status === "🔴最優先" || t.status === "最優先" ? "🔴" : statusEmoji(t.status);
        var deadline = t.deadline ? " [期限:" + t.deadline + "]" : "";
        lines.push(flag + " " + t.task + " (" + t.progress + ")" + deadline);
        if (t.status === "🔴最優先" || t.status === "最優先") {
          lines.push("  └ 完了条件: " + (t.aim || "要確認"));
        }
      });
    } else {
      lines.push("✅ 未完了タスクなし");
    }

    lines.push("─────────────────");
  });

  // 未割当タスクがあれば末尾に
  if (tasksByHonbucho["未割当"] && tasksByHonbucho["未割当"].length > 0) {
    lines.push("");
    lines.push("⚠️【未割当タスク】");
    tasksByHonbucho["未割当"].forEach(function(t) {
      lines.push("• " + t.task + " (" + t.progress + ")");
    });
    lines.push("─────────────────");
  }

  lines.push("");
  lines.push("🚀 今日も前進あるのみ！");

  return lines.join("\n");
}

// LINE Notify にメッセージ送信
function sendToLine(message) {
  var token = CONFIG.LINE_NOTIFY_TOKEN;
  if (!token || token === "YOUR_LINE_NOTIFY_TOKEN_HERE") {
    Logger.log("LINE_NOTIFY_TOKEN が設定されていません");
    return;
  }

  var url = "https://notify-api.line.me/api/notify";
  var options = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    payload: "message=" + encodeURIComponent(message),
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    Logger.log("LINE送信結果: " + response.getResponseCode() + " " + response.getContentText());
  } catch (e) {
    Logger.log("LINE送信エラー: " + e.toString());
  }
}

// ============================================================
// セットアップ用: 朝5時トリガーを設定する関数
// 一度だけ手動で実行する
// ============================================================
function setupDailyTrigger() {
  // 既存トリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "sendDailyReport") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎日 5:00 JST に実行
  ScriptApp.newTrigger("sendDailyReport")
    .timeBased()
    .everyDays(1)
    .atHour(5)
    .inTimezone("Asia/Tokyo")
    .create();

  Logger.log("✅ 毎日05:00 JSTのトリガーを設定しました");
}

// ============================================================
// テスト用: 今すぐLINE送信を確認する関数
// ============================================================
function testSend() {
  sendDailyReport();
}

// ============================================================
// ユーティリティ関数
// ============================================================
function findCol(headers, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var idx = headers.indexOf(candidates[i]);
    if (idx >= 0) return idx;
  }
  return -1;
}

function formatDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Tokyo", "M/d");
  }
  return String(val).trim();
}

function statusEmoji(status) {
  switch (status) {
    case "対応中":   return "🔄";
    case "計画中":   return "📅";
    case "ルーティン": return "🔁";
    case "未着手":   return "⬜";
    default:         return "▪️";
  }
}
