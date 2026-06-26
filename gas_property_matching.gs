/**
 * ============================================================
 *  医療物件マッチング自動化 GAS（AI照合版 v3） 2026-06-01
 *  新着物件メール → Drive保存 → 物件一覧スプシ追記
 *                → 先生リストとAI照合 → 福井宛 下書き作成 → Googleタスク化
 *  ※ Google側で無人稼働（PC・Claude不要）。10分ごとに自動チェック。
 *  ※ APIキー未設定でもテンプレ下書きで動く（フォールバック設計）。
 * ============================================================
 *
 * 【初回セットアップ手順】（5分・1回だけ）
 *  1. https://script.google.com で新規プロジェクト作成、このコードを全文貼り付け
 *  2. 左メニュー「サービス +」→「Tasks API」を追加（識別子は Tasks のまま）
 *  3. ⚙️プロジェクトの設定 → スクリプトプロパティ → 「ANTHROPIC_API_KEY」を追加し
 *     値に Anthropic のAPIキーを貼る（※これで"賢いAI照合"が有効に。未設定でもテンプレで動く）
 *  4. 関数選択で「setupTrigger」を選び ▷実行 → 権限承認のダイアログで許可
 *     （Gmail / Drive / スプレッドシート / Tasks / 外部URL接続 へのアクセスを許可）
 *  5. 以後は10分ごとに自動でprocessPropertyEmailsが走ります
 *  ※ 動作テストは「processPropertyEmails」を手動▷実行 → 実行ログを確認
 */

// ===== 設定（IDは変更不要・必要なら差し替え） =====
var PARENT_FOLDER_ID = '1t_50_gV_xzYEGUt_Zbpz2W5yXrFrOlJA';            // 月別フォルダの親
var LIST_SHEET_ID    = '1a0w6K-fi_BpTGGAVmB1lHqAJYPjnM4M8fw8Rs25ghnc'; // 物件マッチング一覧スプシ
var SENSEI_SHEET_ID  = '1d7FaFIOJNqpb1JTN09H6RIers99S_ejQayHrPPT2fog'; // 先生ニーズリスト(AI照合用)
var FUKUI_EMAIL      = 'fukui@tenant-aw.jp';
var TASKLIST_ID      = '@default';                                     // Googleタスク「マイタスク」
var PROCESSED_LABEL  = '処理済み_物件';                                 // 二重処理防止ラベル

// === 医療物件の確実な配信元（送信元ホワイトリスト主軸・2026-06-01実データ）===
var MEDICAL_SENDERS = ['ph-k.co.jp', 'welcia-yakkyoku.co.jp', 'nicho.co.jp', 'iinkaigyo'];
// === 不動産業者の配信元 ===
var FUDOSAN_SENDERS = ['l-pp.co.jp', 'fudosan-site.jp', 'fuji-plan.net', 'christy.co.jp',
                       'morijuken.co.jp', 'miraiarc.jp', 'livable.jp', 'fountaintokyo.com',
                       'kenbiya.com', 'irios.co.jp', 'inovito.co.jp', 'life-tokyo.jp',
                       'trad-w.com', 'bf-estate.co.jp', 'goodrealtor.co.jp', 'jsc-k.com',
                       'firstraight.com'];
// === 除外（物件ではない・誤爆源）===
var EXCLUDE_SENDERS = ['funds.jp', 'timee.co.jp', 'co-ad.com', 'zenko-kyo.or.jp',
                       'tranbi.com', 'f-mikata.co.jp'];
// === 福井(tenant-aw.jp等)= お客(先生/クリニック)案件のみ拾う。純事務はスキップ ===
var FUKUI_SENDERS = ['tenant-aw.jp', 'fukui1584@icloud.com', 'medilink_fukui'];
var FUKUI_CASE_KW = ['物件', '診療圏', '事業計画', '概要書', '開業', '承継', '移転',
                     'クリニック', '先生', 'テナント', '坪', 'メディカル', '医院', '薬局'];
var FUKUI_SKIP_KW = ['タックシール', 'ロゴ', 'セミナー', '交通費', '農地', '固定資産税'];

// ===== メイン処理 =====
function processPropertyEmails() {
  var processedLabel = getOrCreateLabel(PROCESSED_LABEL);
  var medicalLabel   = getOrCreateLabel('医療物件');
  var fudosanLabel   = getOrCreateLabel('不動産');

  // 過去2日・未処理 を対象（in:inboxは付けない＝業者メールは自動アーカイブされinbox外のため）
  // 送信元ホワイトリストで絞るので、通知/請求等の非物件メールは classify でスキップされる
  var query = 'newer_than:2d -in:sent -in:trash -in:spam -in:chats -label:"' + PROCESSED_LABEL + '"';
  var threads = GmailApp.search(query, 0, 100);
  var count = 0;

  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var msgs = thread.getMessages();
    var msg = msgs[msgs.length - 1];
    var subject = msg.getSubject() || '(件名なし)';
    var from = msg.getFrom();
    var body = msg.getPlainBody();

    var category = classify(from, subject, body);
    if (!category) { continue; } // 対象外メールはスキップ（ラベルも付けない）

    // ① ラベル付与
    thread.addLabel(category === '医療物件' ? medicalLabel : fudosanLabel);

    // ② 添付をDrive当月フォルダへ保存
    var folder = getMonthFolder();
    var saved = [];
    var atts = msg.getAttachments();
    for (var a = 0; a < atts.length; a++) {
      var att = atts[a];
      if (att.getSize() === 0) { continue; }
      var file = folder.createFile(att.copyBlob());
      file.setName(att.getName());
      saved.push(file.getUrl());
    }

    // ③ 物件一覧スプシに1行追記
    var permalink = thread.getPermalink();
    var sheet = SpreadsheetApp.openById(LIST_SHEET_ID).getSheets()[0];
    sheet.appendRow([
      formatDate(msg.getDate()), // 受信日
      category,                  // 種別
      '',                        // 物件No
      subject,                   // 物件名(件名)
      '',                        // 所在
      '',                        // 募集科目
      '',                        // 面積
      '',                        // 賃料/坪
      '',                        // 最寄
      from,                      // 送信元
      saved.length ? saved.join('\n') : permalink, // 資料リンク
      '(要照合)',                // マッチ候補(先生)
      '要対応',                  // ステータス
      ''                         // 備考
    ]);

    // ④ 医療物件は福井宛 下書き作成（AI照合・該当先生がいる時のみ）
    var draftMade = false;
    if (category === '医療物件') {
      draftMade = createFukuiDraft(subject, permalink, saved, body);
    }

    // ⑤ Googleタスク作成
    createTask(category, subject, permalink, saved, draftMade);

    // 二重処理防止ラベル
    thread.addLabel(processedLabel);
    count++;
  }
  Logger.log('処理件数: ' + count);
}

// ===== 分類（送信元ホワイトリスト主軸・誤爆防止）=====
function classify(from, subject, body) {
  var text = subject + ' ' + body;

  // 0) 除外送信元（Funds/タイミー/メディカルアイ等）は即スキップ
  for (var e = 0; e < EXCLUDE_SENDERS.length; e++) {
    if (from.indexOf(EXCLUDE_SENDERS[e]) >= 0) { return null; }
  }

  // 1) 福井 = お客案件のみ医療物件、純事務(タックシール/ロゴ/農地等)はスキップ
  for (var f = 0; f < FUKUI_SENDERS.length; f++) {
    if (from.indexOf(FUKUI_SENDERS[f]) >= 0) {
      for (var s = 0; s < FUKUI_SKIP_KW.length; s++) {
        if (text.indexOf(FUKUI_SKIP_KW[s]) >= 0) { return null; }
      }
      for (var c = 0; c < FUKUI_CASE_KW.length; c++) {
        if (text.indexOf(FUKUI_CASE_KW[c]) >= 0) { return '医療物件'; }
      }
      return null; // 案件キーワードが無い福井メールはスルー
    }
  }

  // 2) 医療物件の配信元（アイリス/ウエルシア/日本調剤）
  for (var i = 0; i < MEDICAL_SENDERS.length; i++) {
    if (from.indexOf(MEDICAL_SENDERS[i]) >= 0) { return '医療物件'; }
  }

  // 3) 不動産業者の配信元 → 医療専用ツールのため拾わない（タスク化ノイズ防止 2026-06-02 菊池指示）
  for (var j = 0; j < FUDOSAN_SENDERS.length; j++) {
    if (from.indexOf(FUDOSAN_SENDERS[j]) >= 0) { return null; }
  }

  // 4) リスト外は対象外（キーワード単独では拾わない＝誤爆防止）
  return null;
}

// ===== 福井宛 下書き（AI照合つき・キー無しはテンプレにフォールバック）=====
// 戻り値: true=下書き作成 / false=該当先生なしで作らず
function createFukuiDraft(subject, permalink, saved, body) {
  var resourceLine = saved.length ? saved.join('\n　　　　　　') : '(元メール参照)';

  // --- AI照合（APIキーがあれば）---
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (apiKey) {
    var senseiList = getSenseiListText();
    var sys = 'あなたはKHD(医療テナント仲介)の秘書です。届いた物件情報を、開業場所を探している先生リストと照合します。'
            + '合致する先生がいれば、紹介元の福井さん宛に「この物件が○○先生のご希望に合いそうです、ご判断ください」という相談メールの本文だけを書いてください。'
            + 'エリア・科目・面積・予算の合致点を箇条書きで明記。売り込まず相手目線で、押し付けない丁寧な文面に。署名は「菊池」。'
            + '合致する先生が一人もいなければ、本文ではなく「該当なし」とだけ返してください。';
    var usr = '【届いた物件メール】\n件名: ' + subject + '\n本文(抜粋):\n' + (body || '').slice(0, 2500)
            + '\n\n【物件資料リンク】\n' + resourceLine
            + '\n\n【開業場所を探している先生リスト（1行=1名）】\n' + senseiList;
    var aiText = callClaude(sys, usr);
    if (aiText) {
      if (aiText.replace(/\s/g, '').indexOf('該当なし') >= 0 && aiText.length < 40) {
        return false; // 合致先生なし→下書き作らない
      }
      var aiBody = aiText + '\n\n――――――\n■ 物件資料：' + resourceLine + '\n■ 元メール：' + permalink;
      GmailApp.createDraft(FUKUI_EMAIL, 'Re: ' + subject + '（先生紹介のご相談）', aiBody);
      return true;
    }
  }

  // --- フォールバック：従来テンプレ ---
  var lines = [
    '福井さん', '',
    'お疲れ様です。菊池です。', '',
    '下記の医療テナント物件の資料が届きましたので共有いたします。',
    'お預かりしている先生のご希望に合致しそうか、ご確認いただけますでしょうか。', '',
    '■ 件名：' + subject,
    '■ 物件資料：' + resourceLine,
    '■ 元メール：' + permalink, '',
    'よろしくお願いいたします。', '', '菊池'
  ];
  GmailApp.createDraft(FUKUI_EMAIL, 'Fw: ' + subject + '（物件共有）', lines.join('\n'));
  return true;
}

// ===== 先生ニーズリストをテキスト化（AIへ渡す照合軸）=====
function getSenseiListText() {
  try {
    var sheet = SpreadsheetApp.openById(SENSEI_SHEET_ID).getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var lines = [];
    for (var r = 0; r < data.length; r++) {
      lines.push(data[r].join(' | '));
    }
    return lines.join('\n');
  } catch (e) {
    Logger.log('先生リスト読込エラー: ' + e);
    return '(先生リスト取得失敗)';
  }
}

// ===== Claude API 呼び出し（キー未設定/失敗時は null）=====
function callClaude(systemPrompt, userPrompt) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) { return null; }
  var payload = {
    model: 'claude-sonnet-4-6',
    max_tokens: 1500,
    system: systemPrompt,
    messages: [{ role: 'user', content: userPrompt }]
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    var json = JSON.parse(resp.getContentText());
    if (json.content && json.content[0] && json.content[0].text) {
      return json.content[0].text;
    }
    Logger.log('Claude応答異常: ' + resp.getContentText().slice(0, 300));
  } catch (e) {
    Logger.log('Claude API エラー: ' + e);
  }
  return null;
}

// ===== Googleタスク作成 =====
function createTask(category, subject, permalink, saved, draftMade) {
  if (category !== '医療物件') { return; } // 医療物件のみタスク化（不動産【確認】ノイズ防止 2026-06-02 菊池指示）
  var notes = [
    '■ 種別：' + category,
    '■ 元メール：' + permalink,
    draftMade
      ? '■ 福井宛の下書き作成済み → Gmailの「下書き」から確認して送信'
      : '■ 下書きなし（不動産／必要なら手動対応）',
    '■ 物件資料：' + (saved.length ? saved.join('  ') : 'なし'),
    '■ 一覧スプシ：https://docs.google.com/spreadsheets/d/' + LIST_SHEET_ID + '/edit'
  ].join('\n');
  var prefix = (category === '医療物件') ? '【要送信】福井宛 ' : '【確認】';
  Tasks.Tasks.insert({ title: prefix + subject, notes: notes }, TASKLIST_ID);
}

// ===== ヘルパー =====
function getMonthFolder() {
  var now = new Date();
  var yy = ('' + now.getFullYear()).slice(2);
  var mm = ('0' + (now.getMonth() + 1)).slice(-2);
  var name = yy + mm; // 例: 2606
  var parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function getOrCreateLabel(name) {
  var label = GmailApp.getUserLabelByName(name);
  return label ? label : GmailApp.createLabel(name);
}

function formatDate(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

// ===== トリガー設定（初回1回だけ実行）=====
function setupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processPropertyEmails') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('processPropertyEmails').timeBased().everyMinutes(10).create();
  Logger.log('トリガー設定完了：10分ごとに自動実行されます');
}
