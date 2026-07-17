// ============================================================
// KHD 入札・公売ウォッチャー v2（穴場サイト巡回 → スプシ記帳 → LINE通知）
// Google Apps Script (GAS) にコピー&ペーストして使用
// ------------------------------------------------------------
// v2の変更点:
//   - 検知した案件を「案件パイプライン」スプシに自動で行追加
//     （入札締切・現物確認日・想定売却額・入札上限・契約金額・想定利益を
// 　　 数式込みでセット。巡回→記帳→通知がスプシで完結）
//   - 「販売先マスタ」シートを自動生成（旧車王等の調査済み連絡先入り）
//   - 初回実行は既読登録のみ（通知の洪水を防止）
// 使い方:
//   1. Googleスプレッドシートを新規作成（名前例: KHD_入札案件パイプライン）
//      → URLの /d/ と /edit の間のIDを CONFIG.SPREADSHEET_ID に貼る
//   2. script.google.com → 本文を貼り付け → LINE_NOTIFY_TOKEN を設定
//   3. setupTrigger を1回実行（毎朝6時の自動巡回を登録）
//   4. run を1回手動実行 → 初回は「初期化完了」ログのみ（通知なし）
//   5. 以後、新着が出た朝だけ LINE が鳴り、スプシに行が増える
// ============================================================

var CONFIG = {
  LINE_NOTIFY_TOKEN: "YOUR_LINE_NOTIFY_TOKEN_HERE", // ← line_notify_gas.js と同じ値に差し替え
                                                    //    （またはスクリプトプロパティ LINE_NOTIFY_TOKEN に設定すれば以後貼り替え不要）
  SPREADSHEET_ID: "1_UPTbcKOvq9xVmsQZR6KAu5hkrREkYnsnn4BU6Eo7q8", // 03_入札案件パイプライン（設定済み）
  MAX_NOTIFY_PER_RUN: 10,
  USER_AGENT: "Mozilla/5.0 (Macintosh) KHD-Watcher/2.1",
};

var SITES = [
  { name: "国税庁 公売情報",            url: "https://www.koubai.nta.go.jp/",            type: "all",   memo: "税務署の差押財産。カメラ・時計・車・不動産まで全部出る" },
  { name: "KSI官公庁オークション",       url: "https://kankocho.jp/",                     type: "all",   memo: "自治体の公売・売払いが集約" },
  { name: "裁判所 不動産競売(BIT)",      url: "https://www.bit.courts.go.jp/",            type: "estate",memo: "競売物件。盛岡地裁管内を重点確認" },
  { name: "東北財務局 国有財産売却",     url: "https://lfb.mof.go.jp/tohoku/kokuyu/index.html", type: "estate", memo: "国有地・庁舎跡地" },
  { name: "関東財務局 国有財産売却",     url: "https://lfb.mof.go.jp/kantou/kokuyu/index.html", type: "estate", memo: "首都圏の国有財産" },
  { name: "日本年金機構 調達・売却情報", url: "https://www.nenkin.go.jp/info/chotatsu/index.html", type: "all", memo: "物品売払い・宿舎売却が稀に出る穴場" },
  { name: "岩手中部水道企業団",          url: "https://www.iwatetyubu-suido.jp/company/company-cat/co_cat_10/", type: "all", memo: "ランクル案件の主。入札情報カテゴリ直監視" },
  { name: "北上市",                      url: "https://www.city.kitakami.iwate.jp/",      type: "all",   memo: "圏域内（参加資格あり）" },
  { name: "花巻市",                      url: "https://www.city.hanamaki.iwate.jp/",      type: "all",   memo: "支店所在地・最優先" },
  { name: "紫波町",                      url: "https://www.town.shiwa.iwate.jp/",         type: "all",   memo: "圏域内" },
  { name: "岩手県",                      url: "https://www.pref.iwate.jp/",               type: "all",   memo: "県有財産・公用車売払い" },
  { name: "NEXCO東日本 調達",            url: "https://www.e-nexco.co.jp/procurement/",   type: "all",   memo: "維持管理車両・備品" },
  { name: "JR東日本 調達情報",           url: "https://www.jreast.co.jp/procurement/",    type: "all",   memo: "鉄道関連資産" },
  { name: "東京都主税局(公売)",          url: "https://www.tax.metro.tokyo.lg.jp/",       type: "all",   memo: "都税事務所公売。動産が多い" },
];

var KEYWORDS = ["売払", "売却", "公売", "入札", "オークション", "競売", "不用品", "不要品", "官公庁"];
var HOT = ["公用車", "車両", "カメラ", "レンズ", "時計", "不動産", "土地", "建物", "ランドクルーザー", "ハイエース"];

// ---- パイプラインの列定義（1行=1案件。巡回が自動で追加、以後は手で育てる）----
var PIPE_HEADERS = [
  "検知日","情報源","件名","URL","区分(車/物/不動産)","ステータス",         // A-F
  "入札締切","現物確認日","最低売却価格","想定売却額(保守)","想定売却額(中央)", // G-K
  "諸費用計","入札上限(自動)","入札額(決定)","契約金額(自動)","想定利益(自動)", // L-P
  "販売先候補","リスク・論点","次アクション","結果"                            // Q-T
];
// ステータスの推奨遷移: 新着 → 精査中 → 入札準備 → 入札済 → 落札/落選 → 売却済/撤退

function run() {
  var props = PropertiesService.getScriptProperties();
  var firstRun = !props.getProperty("initialized");
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  ensureSheets_(ss);
  var pipe = ss.getSheetByName("案件パイプライン");

  var notifications = [];
  SITES.forEach(function (site) {
    try {
      var html = UrlFetchApp.fetch(site.url, {
        muteHttpExceptions: true, followRedirects: true,
        headers: { "User-Agent": CONFIG.USER_AGENT },
      }).getContentText();
      var links = extractLinks_(html, site.url);
      var hits = links.filter(function (l) {
        return KEYWORDS.some(function (k) { return l.text.indexOf(k) >= 0; });
      });
      var seenKey = "seen_" + site.name;
      var seen = JSON.parse(props.getProperty(seenKey) || "[]");
      var fresh = hits.filter(function (l) { return seen.indexOf(hash_(l.text + l.href)) < 0; });
      if (!firstRun) {
        fresh.forEach(function (l) {
          var hot = HOT.some(function (k) { return l.text.indexOf(k) >= 0; });
          notifications.push({ site: site.name, text: l.text, href: l.href, hot: hot });
          appendPipelineRow_(pipe, site, l, hot);
        });
      }
      var newSeen = hits.map(function (l) { return hash_(l.text + l.href); }).concat(seen).slice(0, 300);
      props.setProperty(seenKey, JSON.stringify(newSeen));
    } catch (e) {
      Logger.log("ERROR " + site.name + ": " + e);
    }
  });

  if (firstRun) {
    props.setProperty("initialized", "1");
    Logger.log("初期化完了：既存記事を既読登録しました（通知なし）。明日以降の新着から通知します。");
    sendLine_("\n✅ 入札ウォッチャー初期化完了。14サイトの既存記事を既読登録しました。明日以降、新着があれば通知します。");
    return;
  }
  if (notifications.length === 0) { Logger.log("新着なし"); return; }
  notifications.sort(function (a, b) { return (b.hot ? 1 : 0) - (a.hot ? 1 : 0); });
  var top = notifications.slice(0, CONFIG.MAX_NOTIFY_PER_RUN);
  var lines = top.map(function (n) {
    return (n.hot ? "★" : "・") + "【" + n.site + "】" + n.text.slice(0, 60) + "\n" + n.href;
  });
  var msg = "\n🔔 入札・公売 新着 " + notifications.length + "件（スプシに記帳済み）\n" + lines.join("\n") +
            (notifications.length > top.length ? "\n…他" + (notifications.length - top.length) + "件" : "") +
            "\n📊 " + "https://docs.google.com/spreadsheets/d/" + CONFIG.SPREADSHEET_ID;
  sendLine_(msg);
}

// ---- スプシ整備 ----
function ensureSheets_(ss) {
  var pipe = ss.getSheetByName("案件パイプライン");
  if (!pipe) {
    pipe = ss.insertSheet("案件パイプライン");
    pipe.appendRow(PIPE_HEADERS);
    pipe.getRange(1, 1, 1, PIPE_HEADERS.length).setFontWeight("bold").setBackground("#F1ECE1");
    pipe.setFrozenRows(1);
    seedLandcruiser_(pipe); // 進行中のランクル案件を1行目に自動投入
  }
  var master = ss.getSheetByName("販売先マスタ");
  if (!master) {
    master = ss.insertSheet("販売先マスタ");
    var rows = [
      ["販売先","区分","得意分野","連絡先","メール/フォーム","担当者","メモ"],
      ["旧車王(カレント自動車)","車両","10年超の旧車専門・JPUC認定","045-476-1019","kaitori@currentmotor.co.jp","(取引時に確保)","ランクル実査定100〜200万の実績。引上げ無料"],
      ["フレックス・ドリーム","車両","ランクル・ハイエース専門","0078-6002-427583","https://www.flexdream.jp/","","買取本部=東京都調布市"],
      ["最強買取jp","車両","走行距離不問・全国出張","公式サイト経由","","","事故車修復歴もOKと明言"],
      ["向陽自販","車両","アフリカ・ケニア等輸出","要確認","","","老朽車の可否を要確認"],
      ["友人(ガリバー勤務)","車両","相場観の相談","LINE","","本人","写真レビュー協力。売却先というより顧問枠"],
      ["(不動産)地元仲介・買取再販業者","不動産","※案件発生時に個別選定","","","","競売・公売物件は瑕疵前提。転売先は物件次第"],
      ["(カメラ等動産)買取専門店・ヤフオク/メルカリ","物品","相場が公開されている","","","","古物商許可の取得が前提"],
    ];
    master.getRange(1, 1, rows.length, 7).setValues(rows);
    master.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#F1ECE1");
    master.setFrozenRows(1);
  }
}

function appendPipelineRow_(pipe, site, link, hot) {
  var r = pipe.getLastRow() + 1;
  var kind = hot ? guessKind_(link.text) : (site.type === "estate" ? "不動産" : "");
  pipe.getRange(r, 1, 1, 6).setValues([[
    Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd"),
    site.name, link.text.slice(0, 100), link.href, kind, "新着",
  ]]);
  // 数式列: M=入札上限(自動)=(J保守売却-30万利益-L諸費用)/1.1, O=契約金額=N*1.1, P=想定利益=K中央-O-L
  pipe.getRange(r, 13).setFormula("=IF(J" + r + ">0,(J" + r + "-300000-L" + r + ")/1.1,\"\")");
  pipe.getRange(r, 15).setFormula("=IF(N" + r + ">0,FLOOR(N" + r + "*1.1),\"\")");
  pipe.getRange(r, 16).setFormula("=IF(AND(K" + r + ">0,O" + r + ">0),K" + r + "-O" + r + "-L" + r + ",\"\")");
  // 定型のリスク・論点と次アクションを新着時に自動セット（手で上書きして育てる）
  pipe.getRange(r, 18).setValue("参加資格(地域要件)/現状渡し/一発入札/搬出期限/排ガス規制/修復歴 を精査");
  pipe.getRange(r, 19).setValue("公告PDFを取得→締切・最低価格を記入→勝ち筋10分判定");
}

// 進行中のランクル案件（岩手中部水道企業団・公告148号）を初期投入
function seedLandcruiser_(pipe) {
  var r = pipe.getLastRow() + 1;
  pipe.getRange(r, 1, 1, 12).setValues([[
    "2026-07-14", "岩手中部水道企業団",
    "公用車の売払い（ランクル100 バンVX ディーゼル5MT・H10年式・28万km・車検切れ）",
    "https://www.iwatetyubu-suido.jp/company/15406/",
    "車", "精査中",
    "2026-09-07 17:00", "2026-08-20 10:00-15:00",
    150000, 1000000, 1500000, 120000,
  ]]);
  pipe.getRange(r, 13).setFormula("=IF(J" + r + ">0,(J" + r + "-300000-L" + r + ")/1.1,\"\")");
  pipe.getRange(r, 15).setFormula("=IF(N" + r + ">0,FLOOR(N" + r + "*1.1),\"\")");
  pipe.getRange(r, 16).setFormula("=IF(AND(K" + r + ">0,O" + r + ">0),K" + r + "-O" + r + "-L" + r + ",\"\")");
  pipe.getRange(r, 17).setValue("旧車王(045-476-1019・実査定100〜200万)/フレックス・ドリーム/最強買取jp/向陽自販");
  pipe.getRange(r, 18).setValue("下回り40点が未評価(8/20が全て)/修復歴の有無/NOx・PM法で首都圏登録不可(保有なら花巻登録)/一発入札/9/15全額前払い/9/30搬出");
  pipe.getRange(r, 19).setValue("8/20現物確認の電話予約(0198-41-5315)＋整備士同行打診(コバック北上0197-71-1166)");
}

function guessKind_(text) {
  if (/(公用車|車両|ランドクルーザー|ハイエース|自動車)/.test(text)) return "車";
  if (/(土地|建物|不動産|宅地)/.test(text)) return "不動産";
  if (/(カメラ|レンズ|時計)/.test(text)) return "物";
  return "";
}

function extractLinks_(html, baseUrl) {
  var out = [];
  var re = /<a\s[^>]*href=["']([^"'#]+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  var m;
  while ((m = re.exec(html)) !== null) {
    var text = m[2].replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
    if (!text) continue;
    var href = m[1];
    if (href.indexOf("http") !== 0) {
      var mm = baseUrl.match(/^https?:\/\/[^\/]+/);
      if (!mm) continue;
      href = href.indexOf("/") === 0 ? mm[0] + href : baseUrl.replace(/\/[^\/]*$/, "/") + href;
    }
    out.push({ text: text, href: href });
  }
  return out;
}

function hash_(s) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ((b & 0xff) + 0x100).toString(16).slice(1); }).join("").slice(0, 16);
}

function sendLine_(message) {
  var token = CONFIG.LINE_NOTIFY_TOKEN;
  if (token.indexOf("YOUR_") === 0) {
    token = PropertiesService.getScriptProperties().getProperty("LINE_NOTIFY_TOKEN") || "";
  }
  if (!token) { Logger.log("LINEトークン未設定（通知スキップ・スプシ記帳は実行済み）"); return; }
  UrlFetchApp.fetch("https://notify-api.line.me/api/notify", {
    method: "post",
    headers: { Authorization: "Bearer " + token },
    payload: { message: message },
    muteHttpExceptions: true,
  });
}

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "run") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("run").timeBased().atHour(6).everyDays(1).create();
  Logger.log("毎朝6時のトリガーを設定しました");
}

// 動作テスト用: サイトごとの取得状況だけを確認（通知・記帳なし）
function testFetch() {
  SITES.forEach(function (site) {
    try {
      var res = UrlFetchApp.fetch(site.url, { muteHttpExceptions: true, followRedirects: true,
        headers: { "User-Agent": CONFIG.USER_AGENT } });
      var html = res.getContentText();
      var hits = extractLinks_(html, site.url).filter(function (l) {
        return KEYWORDS.some(function (k) { return l.text.indexOf(k) >= 0; });
      });
      Logger.log("OK " + site.name + " HTTP" + res.getResponseCode() + " 入札関連リンク" + hits.length + "件");
    } catch (e) {
      Logger.log("NG " + site.name + ": " + e);
    }
  });
}
