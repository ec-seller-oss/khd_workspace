// ============================================================
// KHD 入札・公売ウォッチャー v3（穴場サイト巡回 → 詳細抽出 → スプシ記帳 → LINE通知）
// Google Apps Script (GAS) にコピー&ペーストして使用
// ------------------------------------------------------------
// v3の変更点（2026-07-18菊池さん指摘「列が空・ノイズ多い・出口リンクと利益を出せ」対応）:
//   - 検知した案件のリンク先(HTML)を自動で開き、入札締切・最低売却価格を
//     正規表現で抽出してG列・I列に自動記入（PDF直リンクは対象外＝次アクションに明記）
//   - 「売る側」の案件だけ拾うフィルタに変更（売払/公売/売却系のみ。
//     参加資格審査・FAQ・電子入札システム等の運用ページはNOISEとして除外）
//   - 同一巡回内の重複リンクを排除
//   - U列「売却出口リンク」を自動生成（車→グーネット相場/不動産→国交省 不動産情報
//     ライブラリ/物→オークファン。1クリックで売値の当たりを取れる）
//   - V列「粗利@最低価格(自動)」＝想定売却額(中央)−最低売却価格×1.1−諸費用 を
//     数式でセット。諸費用はデフォルト10万円を自動記入（手で上書き可）
//   - cleanupPipeline() を追加：ノイズで埋まった既存の初期スキャン行を一括削除
// v2.5の変更点（2026-07-18菊池さん指摘対応）:
//   - 「サイトを初めて巡回した時は全部だまって既読にする」設計をやめた。
//     初めて見るサイトは、その時点にある案件を「初期スキャン」として
//     全部パイプラインに記帳する＝起動直後から掘り出し物を拾える。
//   - 今後SITESに新しいサイトを追加した瞬間も、その1回目の巡回で
//     自動的に棚卸し（初期スキャン）される＝営業直結まで最短。
// v2の変更点:
//   - 検知した案件を「案件パイプライン」スプシに自動で行追加
//     （入札締切・現物確認日・想定売却額・入札上限・契約金額・想定利益を
// 　　 数式込みでセット。巡回→記帳→通知がスプシで完結）
//   - 「販売先マスタ」シートを自動生成（旧車王等の調査済み連絡先入り）
// 使い方:
//   1. Googleスプレッドシートを新規作成（名前例: KHD_入札案件パイプライン）
//      → URLの /d/ と /edit の間のIDを CONFIG.SPREADSHEET_ID に貼る
//   2. script.google.com → 本文を貼り付け → LINE_NOTIFY_TOKEN を設定
//   3. setupTrigger を1回実行（毎日6時・11時・17時の自動巡回=1日3回を登録）
//   4. run を1回手動実行 → 15サイト分の現在の案件が「初期スキャン」として
//      一気にスプシに記帳される（掘り出し物はここで探す）。LINEにも通知が届く
//   5. 以後は差分（新着）だけがスプシに追加され、LINEが鳴る
// ============================================================

var CONFIG = {
  LINE_NOTIFY_TOKEN: "YOUR_LINE_NOTIFY_TOKEN_HERE", // ← line_notify_gas.js と同じ値に差し替え
                                                    //    （またはスクリプトプロパティ LINE_NOTIFY_TOKEN に設定すれば以後貼り替え不要）
  SPREADSHEET_ID: "1_UPTbcKOvq9xVmsQZR6KAu5hkrREkYnsnn4BU6Eo7q8", // 03_入札案件パイプライン（設定済み）
  MAX_NOTIFY_PER_RUN: 10,
  DETAIL_FETCH_LIMIT: 12,   // 1回の巡回でリンク先本文を開く最大数（GAS実行時間対策）
  DEFAULT_COSTS: 100000,    // 諸費用のデフォルト値（陸送・登録等の仮置き。行ごとに上書き）
  USER_AGENT: "Mozilla/5.0 (Macintosh) KHD-Watcher/3.0",
};

// v2.2: 2026-07-17のtestFetchログを受けてURL修正
//   404だった4件を実在URLに差し替え／JR東(403=bot遮断)を国有財産売却情報サイトに交代／
//   北上・花巻は入札・契約ページ直リンクに変更／年金機構はスペル修正(chotatu)
var SITES = [
  { name: "国税庁 公売情報",            url: "https://www.koubai.nta.go.jp/",            type: "all",   memo: "✅動作確認済(37件)。税務署の差押財産" },
  { name: "KSI官公庁オークション",       url: "https://kankocho.jp/",                     type: "all",   memo: "⚠JS描画のため取得弱め。国税庁・自治体側でカバー" },
  { name: "裁判所 不動産競売(BIT)",      url: "https://www.bit.courts.go.jp/",            type: "estate",memo: "✅動作確認済(10件)。盛岡地裁管内を重点確認" },
  { name: "東北財務局 国有財産(公示中)", url: "https://lfb.mof.go.jp/tohoku/b6_baikyaku/kouji.html", type: "estate", memo: "現在公示中の一般競争入札(売却)" },
  { name: "東北財務局 国有財産(予定)",   url: "https://lfb.mof.go.jp/tohoku/b6_baikyaku/nyuusatsu_yotei.html", type: "estate", memo: "入札予定物件＝先回り用" },
  { name: "関東財務局 国有財産売却",     url: "https://lfb.mof.go.jp/kantou/kanzai/mokuji_00001.htm", type: "estate", memo: "現在公示中の売却物件" },
  { name: "日本年金機構 入札公告",       url: "https://www.nenkin.go.jp/chotatu/nyusatsu/index.html", type: "all", memo: "物品売払い・不用品処分が稀に出る穴場" },
  { name: "岩手中部水道企業団",          url: "https://www.iwatetyubu-suido.jp/company/company-cat/co_cat_10/", type: "all", memo: "⚠UAリトライ後も501。サーバ側WAFの可能性、要手動確認" },
  { name: "北上市 入札・契約",           url: "https://www.city.kitakami.iwate.jp/life/shisei/nyusatsu_keiyaku/index.html", type: "all", memo: "圏域内(参加資格あり)" },
  { name: "花巻市 入札・契約",           url: "https://www.city.hanamaki.iwate.jp/business/nyusatsu_keiyaku/index.html", type: "all", memo: "支店所在地・最優先" },
  { name: "紫波町",                      url: "https://www.town.shiwa.iwate.jp/",         type: "all",   memo: "圏域内" },
  { name: "岩手県",                      url: "https://www.pref.iwate.jp/",               type: "all",   memo: "✅動作確認済(20件)" },
  { name: "NEXCO東日本 調達",            url: "https://www.e-nexco.co.jp/bids/",          type: "all",   memo: "調達・お取引トップ" },
  { name: "国有財産売却情報サイト",      url: "https://kokuyuzaisan.akiya-athome.jp/",    type: "estate",memo: "⚠403(bot遮断)。自動巡回不可、要月1手動確認" },
  { name: "東京都主税局(公売)",          url: "https://www.tax.metro.tokyo.lg.jp/",       type: "all",   memo: "✅動作確認済(8件)。動産が多い" },
];

// 「役所が売る側」の案件を示す語（これに当たれば拾う）
var SELL_KEYWORDS = ["売払", "売り払い", "売却", "公売", "競売", "オークション", "不用品", "不要品", "不用物品"];
// 狙い目の品目語（これ＋「入札」の組合せでも拾う）
var HOT = ["公用車", "車両", "乗用車", "カメラ", "レンズ", "時計", "不動産", "土地", "建物", "ランドクルーザー", "ハイエース", "バイク", "重機"];
// 運用ページ・調達側のノイズ（当たったら除外）
var NOISE = ["参加資格", "審査申請", "電子入札", "指名停止", "入札結果", "契約結果", "入札・契約結果",
  "流れ", "FAQ", "よくある", "ガイドライン", "様式", "要領", "心得", "中止情報", "監視委員会",
  "メール配信", "委託", "役務", "工事", "コンサル", "調査業務", "検査", "名簿", "同等品",
  "委任状", "マニュアル", "ダウンロード", "検索", "照会", "ポータル", "システム", "配信サービス", "はこちら"];

function isTarget_(text) {
  if (NOISE.some(function (k) { return text.indexOf(k) >= 0; })) return false;
  if (SELL_KEYWORDS.some(function (k) { return text.indexOf(k) >= 0; })) return true;
  // 「入札」単体はノイズが多すぎるので、品目語(HOT)と同時に出た時だけ拾う
  var hasHot = HOT.some(function (k) { return text.indexOf(k) >= 0; });
  return hasHot && text.indexOf("入札") >= 0;
}

// ---- パイプラインの列定義（1行=1案件。巡回が自動で追加、以後は手で育てる）----
var PIPE_HEADERS = [
  "検知日","情報源","件名","URL","区分(車/物/不動産)","ステータス",         // A-F
  "入札締切","現物確認日","最低売却価格","想定売却額(保守)","想定売却額(中央)", // G-K
  "諸費用計","入札上限(自動)","入札額(決定)","契約金額(自動)","想定利益(自動)", // L-P
  "販売先候補","リスク・論点","次アクション","結果",                           // Q-T
  "売却出口リンク","粗利@最低価格(自動)"                                        // U-V
];
// ステータスの推奨遷移: 新着 → 精査中 → 入札準備 → 入札済 → 落札/落選 → 売却済/撤退

// 方針（2026-07-18菊池さん指摘で修正）：
// 「初回だけは全部だまって既読にする」は掘り出し物を見ずに握りつぶす設計ミスだった。
// → サイトごとに「このサイトを初めて見る巡回か」を判定し、初めてなら
//   その時点で出ている案件を"初期スキャン"としてそのままパイプラインに全部記帳する。
//   これにより①起動直後から狙い目を拾える②今後サイトを追加した瞬間もその場で棚卸しされる。
function run() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  ensureSheets_(ss);
  var pipe = ss.getSheetByName("案件パイプライン");

  var notifications = [];
  var initialScanCount = 0;
  var detailBudget = { left: CONFIG.DETAIL_FETCH_LIMIT };
  SITES.forEach(function (site) {
    try {
      var html = fetchSite_(site.url).getContentText();
      var links = extractLinks_(html, site.url);
      // 同一ページ内の重複リンク（メニューと本文の二重掲載など）を除去
      var inPage = {};
      var hits = links.filter(function (l) {
        if (!isTarget_(l.text)) return false;
        var h = hash_(l.text + l.href);
        if (inPage[h]) return false;
        inPage[h] = true;
        return true;
      });
      var seenKey = "seen_" + site.name;
      var seenRaw = props.getProperty(seenKey);
      var siteFirstScan = !seenRaw; // このサイトを初めて巡回する（=このサイトの掘り出し物をまだ一度も棚卸ししていない）
      var seen = JSON.parse(seenRaw || "[]");
      var fresh = hits.filter(function (l) { return seen.indexOf(hash_(l.text + l.href)) < 0; });
      fresh.forEach(function (l) {
        var hot = HOT.some(function (k) { return l.text.indexOf(k) >= 0; });
        notifications.push({ site: site.name, text: l.text, href: l.href, hot: hot, firstScan: siteFirstScan });
        appendPipelineRow_(pipe, site, l, hot, siteFirstScan, detailBudget);
        if (siteFirstScan) initialScanCount++;
      });
      var newSeen = hits.map(function (l) { return hash_(l.text + l.href); }).concat(seen).slice(0, 300);
      props.setProperty(seenKey, JSON.stringify(newSeen));
    } catch (e) {
      Logger.log("ERROR " + site.name + ": " + e);
    }
  });

  if (notifications.length === 0) { Logger.log("新着なし"); return; }
  // 初期スキャン(掘り出し物の棚卸し)を優先、次にHOTワード命中を優先
  notifications.sort(function (a, b) {
    return (b.firstScan ? 2 : 0) + (b.hot ? 1 : 0) - ((a.firstScan ? 2 : 0) + (a.hot ? 1 : 0));
  });
  var top = notifications.slice(0, CONFIG.MAX_NOTIFY_PER_RUN);
  var lines = top.map(function (n) {
    var mark = n.firstScan ? "🆕初期" : (n.hot ? "★" : "・");
    return mark + "【" + n.site + "】" + n.text.slice(0, 60) + "\n" + n.href;
  });
  var header = "\n🔔 入札・公売 検知 " + notifications.length + "件（スプシに記帳済み）";
  if (initialScanCount > 0) {
    header += "\n　うち🆕初期スキャン(このサイトを初めて棚卸し)" + initialScanCount + "件＝掘り出し物が眠っているかも、要チェック";
  }
  var msg = header + "\n" + lines.join("\n") +
            (notifications.length > top.length ? "\n…他" + (notifications.length - top.length) + "件（全件はスプシで）" : "") +
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
  } else {
    // 旧バージョンのシート（20列）にU/V列ヘッダーを追記
    var lastCol = pipe.getRange(1, 1, 1, pipe.getMaxColumns()).getValues()[0].filter(String).length;
    if (lastCol < PIPE_HEADERS.length) {
      pipe.getRange(1, lastCol + 1, 1, PIPE_HEADERS.length - lastCol)
          .setValues([PIPE_HEADERS.slice(lastCol)])
          .setFontWeight("bold").setBackground("#F1ECE1");
    }
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

function appendPipelineRow_(pipe, site, link, hot, siteFirstScan, detailBudget) {
  var r = pipe.getLastRow() + 1;
  var kind = hot ? guessKind_(link.text) : (site.type === "estate" ? "不動産" : "");
  var status = siteFirstScan ? "初期スキャン" : "新着";
  pipe.getRange(r, 1, 1, 6).setValues([[
    Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd"),
    site.name, link.text.slice(0, 100), link.href, kind, status,
  ]]);
  // 諸費用デフォルト（陸送・登録等の仮置き＝手で上書き前提。これでV列が即計算される）
  pipe.getRange(r, 12).setValue(CONFIG.DEFAULT_COSTS);
  // 数式列: M=入札上限(自動), O=契約金額(自動), P=想定利益(入札額ベース), V=粗利@最低価格
  pipe.getRange(r, 13).setFormula("=IF(J" + r + ">0,(J" + r + "-300000-L" + r + ")/1.1,\"\")");
  pipe.getRange(r, 15).setFormula("=IF(N" + r + ">0,FLOOR(N" + r + "*1.1),\"\")");
  pipe.getRange(r, 16).setFormula("=IF(AND(K" + r + ">0,N" + r + ">0),K" + r + "-FLOOR(N" + r + "*1.1)-L" + r + ",\"\")");
  pipe.getRange(r, 22).setFormula("=IF(AND(I" + r + ">0,K" + r + ">0),K" + r + "-FLOOR(I" + r + "*1.1)-L" + r + ",\"\")");
  // U列: 売却出口リンク（区分に応じた相場検索を1クリックで）
  pipe.getRange(r, 21).setFormula(exitLinkFormula_(kind, link.text));
  // 詳細抽出: リンク先本文から締切・最低売却価格を拾う（HTML限定・回数制限つき）
  var extractNote = "";
  if (/\.pdf(\?|$)/i.test(link.href)) {
    extractNote = "リンク先がPDFのため自動抽出不可→開いて締切・最低価格をG/I列に記入";
  } else if (detailBudget && detailBudget.left > 0) {
    detailBudget.left--;
    var d = extractDetail_(link.href);
    if (d.deadline) pipe.getRange(r, 7).setValue(d.deadline);
    if (d.minPrice) pipe.getRange(r, 9).setValue(d.minPrice);
    if (!d.deadline && !d.minPrice) extractNote = "本文から締切・価格を自動抽出できず→開いて確認";
  } else {
    extractNote = "詳細抽出の回数上限到達→開いて確認";
  }
  pipe.getRange(r, 18).setValue(siteFirstScan
    ? "初期スキャンで検出＝締切切れ・条件不一致の可能性もあり要一次選別"
    : "参加資格(地域要件)/現状渡し/一発入札/搬出期限/排ガス規制/修復歴 を精査");
  pipe.getRange(r, 19).setValue(
    (extractNote ? extractNote + " ／ " : "") +
    "想定売却額(K列)を出口リンクで調べて記入→V列に粗利が自動表示");
}

// 区分に応じた「売却出口」相場検索リンク（HYPERLINK数式）
function exitLinkFormula_(kind, title) {
  // 件名から検索語を作る（先頭の記号・組織名カッコ書きを軽く除去して30字まで）
  var q = title.replace(/[【】《》（）()「」]/g, " ").replace(/\s+/g, " ").trim().slice(0, 30);
  var enc = encodeURIComponent(q);
  var url, label;
  if (kind === "車") {
    url = "https://www.goo-net.com/php/search/summary.php?keyword=" + enc;
    label = "🚗グーネット相場";
  } else if (kind === "不動産") {
    url = "https://www.reinfolib.mlit.go.jp/";
    label = "🏠不動産情報ライブラリ(国交省・成約価格)";
  } else if (kind === "物") {
    url = "https://aucfan.com/search1/q-" + enc + "/";
    label = "📦オークファン相場";
  } else {
    url = "https://www.google.com/search?q=" + enc + "+%E7%9B%B8%E5%A0%B4+%E4%B8%AD%E5%8F%A4";
    label = "🔎相場を検索";
  }
  return '=HYPERLINK("' + url + '","' + label + '")';
}

// リンク先HTMLから 入札締切/最低売却価格 を抽出（ベストエフォート）
function extractDetail_(url) {
  var out = { deadline: "", minPrice: "" };
  try {
    var res = fetchSite_(url);
    if (res.getResponseCode() !== 200) return out;
    var text = res.getContentText().replace(/<[^>]+>/g, " ").replace(/\s+/g, " ");
    var pm = text.match(/(最低売却価格|最低入札価格|売却基準価額|予定価格|最低価格)[^0-9０-９]{0,20}([0-9０-９,，]+)\s*円/);
    if (pm) {
      var num = pm[2].replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 65248); }).replace(/[,，]/g, "");
      var v = parseInt(num, 10);
      if (v > 0) out.minPrice = v;
    }
    var dm = text.match(/(入札書?の?提出期限|入札期限|提出期限|受付期限|入札期間|受付期間|開札日)[^。]{0,30}?((令和|平成)\s*[0-9０-９]+\s*年\s*[0-9０-９]+\s*月\s*[0-9０-９]+\s*日|[0-9]{4}\s*年\s*[0-9]{1,2}\s*月\s*[0-9]{1,2}\s*日)/);
    if (dm) out.deadline = dm[1] + "：" + dm[2].replace(/\s+/g, "");
  } catch (e) {}
  return out;
}

// 【必要な時だけ実行】v2.5以前が書き込んだノイズ行を一括削除する。
// 対象＝ステータスが「初期スキャン」または「新着」のままで、新フィルタ(isTarget_)に
// 通らない行。手で育て始めた行（ステータス変更済み）は消さない。
function cleanupPipeline() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var pipe = ss.getSheetByName("案件パイプライン");
  if (!pipe) { Logger.log("シートなし"); return; }
  var data = pipe.getDataRange().getValues();
  var removed = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    var status = String(data[i][5] || "");
    var title = String(data[i][2] || "");
    if ((status === "初期スキャン" || status === "新着") && !isTarget_(title)) {
      pipe.deleteRow(i + 1);
      removed++;
    }
  }
  Logger.log("ノイズ行を " + removed + " 行削除しました（ステータスを手で変えた行は残しています）");
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
  pipe.getRange(r, 16).setFormula("=IF(AND(K" + r + ">0,N" + r + ">0),K" + r + "-FLOOR(N" + r + "*1.1)-L" + r + ",\"\")");
  pipe.getRange(r, 17).setValue("旧車王(045-476-1019・実査定100〜200万)/フレックス・ドリーム/最強買取jp/向陽自販");
  pipe.getRange(r, 21).setFormula('=HYPERLINK("https://www.goo-net.com/php/search/summary.php?keyword=' + encodeURIComponent("ランドクルーザー100 ディーゼル") + '","🚗グーネット相場")');
  pipe.getRange(r, 22).setFormula("=IF(AND(I" + r + ">0,K" + r + ">0),K" + r + "-FLOOR(I" + r + "*1.1)-L" + r + ",\"\")");
  pipe.getRange(r, 18).setValue("下回り40点が未評価(8/20が全て)/修復歴の有無/NOx・PM法で首都圏登録不可(保有なら花巻登録)/一発入札/9/15全額前払い/9/30搬出");
  pipe.getRange(r, 19).setValue("8/20現物確認の電話予約(0198-41-5315)＋整備士同行打診(コバック北上0197-71-1166)");
}

function guessKind_(text) {
  if (/(公用車|車両|ランドクルーザー|ハイエース|自動車)/.test(text)) return "車";
  if (/(土地|建物|不動産|宅地)/.test(text)) return "不動産";
  if (/(カメラ|レンズ|時計)/.test(text)) return "物";
  return "";
}

// UA起因のブロック(501等)対策: 非200なら標準UAなしで1回リトライ
function fetchSite_(url) {
  var res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true, followRedirects: true,
    headers: { "User-Agent": CONFIG.USER_AGENT },
  });
  if (res.getResponseCode() !== 200) {
    var retry = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (retry.getResponseCode() === 200) return retry;
  }
  return res;
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

// 1日3回巡回（6時・11時・17時）。更新が集中しやすい時間帯を挟んで
// 「まとめて読む」ペースに合わせる。時刻を変えたい場合は RUN_HOURS を編集。
var RUN_HOURS = [6, 11, 17];

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "run") ScriptApp.deleteTrigger(t);
  });
  RUN_HOURS.forEach(function (h) {
    ScriptApp.newTrigger("run").timeBased().atHour(h).everyDays(1).create();
  });
  Logger.log("トリガーを設定しました：毎日 " + RUN_HOURS.join("時・") + "時 に巡回します（計" + RUN_HOURS.length + "回/日）");
}

// 【1回だけ実行してください】旧v2.2/v2.3/v2.4は「初回は既読のみ」でscript
// propertiesに全案件を既読登録済みのため、コードを貼り替えただけでは
// 過去分（掘り出し物）が再表示されない。これを実行して既読状態をリセットしてから
// run() を実行すると、15サイト分の現在の案件が改めて「初期スキャン」として
// 全部パイプラインに記帳される。以後の通常運用では使わない。
function resetSeenState() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var cleared = 0;
  Object.keys(all).forEach(function (k) {
    if (k.indexOf("seen_") === 0 || k === "initialized") {
      props.deleteProperty(k);
      cleared++;
    }
  });
  Logger.log("既読状態をリセットしました（" + cleared + "件のプロパティを削除）。次に run() を実行してください。");
}

// 動作テスト用: サイトごとの取得状況だけを確認（通知・記帳なし）
function testFetch() {
  SITES.forEach(function (site) {
    try {
      var res = fetchSite_(site.url);
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
