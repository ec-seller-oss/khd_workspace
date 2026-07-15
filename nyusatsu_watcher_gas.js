// ============================================================
// KHD 入札・公売ウォッチャー（穴場サイト巡回 → LINE通知）
// Google Apps Script (GAS) にコピー&ペーストして使用
// ------------------------------------------------------------
// 目的: 官公庁・インフラ機関の「売払い/公売/入札」情報を毎日巡回し、
//       新着があればLINEに通知する（車両・カメラ等備品・不動産）
// 使い方:
//   1. script.google.com → 新規プロジェクト → 本文を貼り付け
//   2. CONFIG.LINE_NOTIFY_TOKEN を設定（line_notify_gas.jsと同じトークンでOK）
//   3. 関数 setupTrigger を1回実行 → 毎朝6時に自動巡回
//   4. 初回は run を手動実行（既存記事を「既読」として記録）
// 運用: SITESに行を足すだけで巡回先を増やせる
// ============================================================

var CONFIG = {
  LINE_NOTIFY_TOKEN: "YOUR_LINE_NOTIFY_TOKEN_HERE", // ← line_notify_gas.js と同じでOK
  MAX_NOTIFY_PER_RUN: 10,       // 1回の通知上限（初回の洪水防止）
  USER_AGENT: "Mozilla/5.0 (Macintosh) KHD-Watcher/1.0",
};

// 巡回先。matchType: "car"=車両 / "goods"=カメラ・備品 / "estate"=不動産 / "all"=全部
var SITES = [
  // ---- 全国横断（まずここが本命） ----
  { name: "国税庁 公売情報",            url: "https://www.koubai.nta.go.jp/",            type: "all",
    memo: "税務署の差押財産。カメラ・時計・車・不動産まで全部出る" },
  { name: "KSI官公庁オークション",       url: "https://kankocho.jp/",                     type: "all",
    memo: "旧ヤフー官公庁オークションの後継。自治体の公売・売払いが集約" },
  { name: "裁判所 不動産競売(BIT) 東北", url: "https://www.bit.courts.go.jp/",            type: "estate",
    memo: "競売物件。花巻支店圏(盛岡地裁)を重点確認" },
  { name: "東北財務局 国有財産売却",     url: "https://lfb.mof.go.jp/tohoku/kokuyu/index.html", type: "estate",
    memo: "国有地・庁舎跡地等の一般競争入札" },
  { name: "関東財務局 国有財産売却",     url: "https://lfb.mof.go.jp/kantou/kokuyu/index.html", type: "estate",
    memo: "首都圏の国有財産" },

  // ---- 年金・共済系（穴場） ----
  { name: "日本年金機構 調達・売却情報", url: "https://www.nenkin.go.jp/info/chotatsu/index.html", type: "all",
    memo: "年金事務所の物品売払い・宿舎等の不動産売却が稀に出る" },

  // ---- 岩手（支店登記で参加資格がある地元圏・最優先） ----
  { name: "岩手中部水道企業団",          url: "https://www.iwatetyubu-suido.jp/company/company-cat/co_cat_10/", type: "all",
    memo: "今回のランクル案件の主。入札情報カテゴリを直接監視" },
  { name: "北上市 入札・契約",           url: "https://www.city.kitakami.iwate.jp/",      type: "all",
    memo: "圏域内。公用車・備品の売払い" },
  { name: "花巻市 入札・契約",           url: "https://www.city.hanamaki.iwate.jp/",      type: "all",
    memo: "支店所在地。最優先ウォッチ" },
  { name: "紫波町 入札情報",             url: "https://www.town.shiwa.iwate.jp/",         type: "all",
    memo: "圏域内" },
  { name: "岩手県 入札情報(売払い)",     url: "https://www.pref.iwate.jp/",               type: "all",
    memo: "県有財産・公用車の売払い" },

  // ---- インフラ機関（穴場） ----
  { name: "NEXCO東日本 入札・調達",      url: "https://www.e-nexco.co.jp/procurement/",   type: "all",
    memo: "維持管理車両・備品の売払いが稀に出る" },
  { name: "JR東日本 グループ調達情報",   url: "https://www.jreast.co.jp/procurement/",    type: "all",
    memo: "鉄道関連資産。頻度低いが出れば面白い" },
  { name: "東京都 公売・売払い(主税局)", url: "https://www.tax.metro.tokyo.lg.jp/",       type: "all",
    memo: "都税事務所の公売。カメラ・時計等の動産が多い" },
];

// 新着判定に使うキーワード（リンクテキストにこれが含まれたら「入札関連」とみなす）
var KEYWORDS = ["売払", "売却", "公売", "入札", "オークション", "競売", "不用品", "不要品", "官公庁"];
// 通知対象をさらに絞る強調キーワード（当たり案件の目印。含まれたら★付きで通知）
var HOT = ["公用車", "車両", "カメラ", "レンズ", "時計", "不動産", "土地", "建物", "ランドクルーザー", "ハイエース"];

function run() {
  var props = PropertiesService.getScriptProperties();
  var notifications = [];
  SITES.forEach(function (site) {
    try {
      var html = UrlFetchApp.fetch(site.url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: { "User-Agent": CONFIG.USER_AGENT },
      }).getContentText();
      var links = extractLinks_(html, site.url);
      var hits = links.filter(function (l) {
        return KEYWORDS.some(function (k) { return l.text.indexOf(k) >= 0; });
      });
      var seenKey = "seen_" + site.name;
      var seen = JSON.parse(props.getProperty(seenKey) || "[]");
      var fresh = hits.filter(function (l) { return seen.indexOf(hash_(l.text + l.href)) < 0; });
      fresh.forEach(function (l) {
        var hot = HOT.some(function (k) { return l.text.indexOf(k) >= 0; });
        notifications.push({ site: site.name, text: l.text, href: l.href, hot: hot });
      });
      // 既読を更新（最新300件だけ保持）
      var newSeen = hits.map(function (l) { return hash_(l.text + l.href); }).concat(seen).slice(0, 300);
      props.setProperty(seenKey, JSON.stringify(newSeen));
    } catch (e) {
      Logger.log("ERROR " + site.name + ": " + e);
    }
  });

  if (notifications.length === 0) { Logger.log("新着なし"); return; }
  // ★HOT優先で上限まで
  notifications.sort(function (a, b) { return (b.hot ? 1 : 0) - (a.hot ? 1 : 0); });
  var top = notifications.slice(0, CONFIG.MAX_NOTIFY_PER_RUN);
  var lines = top.map(function (n) {
    return (n.hot ? "★" : "・") + "【" + n.site + "】" + n.text.slice(0, 60) + "\n" + n.href;
  });
  var msg = "\n🔔 入札・公売 新着 " + notifications.length + "件\n" + lines.join("\n") +
            (notifications.length > top.length ? "\n…他" + (notifications.length - top.length) + "件" : "");
  sendLine_(msg);
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
      var origin = baseUrl.match(/^https?:\/\/[^\/]+/)[0];
      href = href.indexOf("/") === 0 ? origin + href : baseUrl.replace(/\/[^\/]*$/, "/") + href;
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
  UrlFetchApp.fetch("https://notify-api.line.me/api/notify", {
    method: "post",
    headers: { Authorization: "Bearer " + CONFIG.LINE_NOTIFY_TOKEN },
    payload: { message: message },
    muteHttpExceptions: true,
  });
}

// 毎朝6時に自動実行するトリガーを設定（1回だけ実行）
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "run") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("run").timeBased().atHour(6).everyDays(1).create();
  Logger.log("毎朝6時のトリガーを設定しました");
}
