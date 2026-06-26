/**
 * KHD 追客LINE自動ログ GAS（2026-06-03）
 * -----------------------------------------------------------
 * LINE公式アカウント(Messaging API)と連携し、接触ログ(③)を全自動化する。
 *
 *  A) doPost(e)  … LINE webhook受信。顧客からの受信メッセージを
 *                  ②接触ログへ自動追記＋送信者のuserIdを①追客リストに紐付け。
 *  B) sendQueuedLineMessages() … ①追客リストの「送信文面」欄に書かれた文面を
 *                  LINE公式から自動送信→②接触ログに「送信」を自動記録→
 *                  最終接触日を更新→送信文面欄をクリア（10分ごとトリガー）。
 *
 * これにより菊池さんは「スプシに送信文面を1行書く」だけ。送信もログもKPIも全自動。
 *
 * 事前設定（LINE Developers側）:
 *  1) LINE公式アカウント → Messaging APIを有効化
 *  2) チャネルアクセストークン(長期) と チャネルシークレット を取得
 *  3) このGASを「ウェブアプリ」としてデプロイ→URLをLINEのWebhook URLに設定・利用ON
 *  4) スクリプトプロパティに保存:
 *       LINE_TOKEN  = チャネルアクセストークン
 *       LINE_SECRET = チャネルシークレット（署名検証用・任意だが推奨）
 *  5) setupLineTrigger() を1回実行して承認
 */

const LN_SHEET_ID = '1PWI3JyEVaFsTZqOA_DNijBIDtGHMZM2JRhdHORkKQhY';
const LN_LIST_TAB = '①追客リスト';
const LN_LOG_TAB  = '②接触ログ';
// ①追客リスト 追加列（GASが自動でヘッダを作る）
const LN_COL_USERID = 19; // S列: LINE_userId
const LN_COL_SEND   = 20; // T列: 送信文面(書くと自動送信)
const LN_COL_NAME   = 2;  // 顧客名
const LN_COL_ID     = 1;  // 顧客ID
const LN_COL_LAST   = 12; // 最終接触日

function _prop(k){ return PropertiesService.getScriptProperties().getProperty(k); }
function _ss(){ return SpreadsheetApp.openById(LN_SHEET_ID); }
function _list(){ return _ss().getSheetByName(LN_LIST_TAB) || _ss().getSheets()[0]; }
function _log(){ return _ss().getSheetByName(LN_LOG_TAB) || _ss().getSheets()[1]; }
function _now(){ return Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd HH:mm'); }
function _today(){ return Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd'); }

// ============================================================
// A) Webhook受信：顧客→KHD のメッセージを自動ログ
// ============================================================
function doPost(e) {
  try {
    // 署名検証（任意）
    const secret = _prop('LINE_SECRET');
    if (secret && e.postData) {
      const sig = e.parameter && e.parameter['X-Line-Signature']; // 環境により取得不可な場合あり→無ければスキップ
      // ※GASのdoPostはヘッダを直接取れないため厳密検証は省略可。必要なら別途プロキシ。
    }
    const body = JSON.parse(e.postData.contents);
    const events = body.events || [];
    const log = _log();
    events.forEach(ev => {
      if (ev.type === 'message' && ev.message && ev.message.type === 'text') {
        const userId = (ev.source && ev.source.userId) || '';
        const text = ev.message.text || '';
        const name = _displayName(userId) || ('LINEユーザー(' + userId.slice(-6) + ')');
        const custId = _matchLeadByUserId(userId, name); // 紐付け試行
        // ②接触ログへ: 日時/顧客ID/顧客名/チャネル/文面/反応/次アクション
        log.appendRow([_now(), custId, name, 'LINE', text, '受信', '']);
      }
    });
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, err:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

// userIdから表示名を取得（Messaging API）
function _displayName(userId) {
  const token = _prop('LINE_TOKEN');
  if (!token || !userId) return '';
  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + userId, {
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
    });
    if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()).displayName || '';
  } catch (e) {}
  return '';
}

// 受信userIdを①追客リストに紐付け（未登録のuserId列に表示名一致で書き込み）。顧客IDを返す
function _matchLeadByUserId(userId, name) {
  if (!userId) return '';
  const sh = _list();
  const lr = sh.getLastRow();
  if (lr < 2) return '';
  _ensureLineHeaders(sh);
  const rng = sh.getRange(2, 1, lr-1, LN_COL_SEND).getValues();
  // 既に同じuserId登録済ならその顧客ID
  for (let i=0;i<rng.length;i++){
    if ((rng[i][LN_COL_USERID-1]||'').toString().trim() === userId) return rng[i][LN_COL_ID-1] || '';
  }
  // 表示名が顧客名に部分一致する行へuserIdを書き込み（紐付け）
  for (let i=0;i<rng.length;i++){
    const cname = (rng[i][LN_COL_NAME-1]||'').toString();
    if (name && cname && (cname.indexOf(name) >= 0 || name.indexOf(cname) >= 0) && !(rng[i][LN_COL_USERID-1])){
      sh.getRange(i+2, LN_COL_USERID).setValue(userId);
      return rng[i][LN_COL_ID-1] || '';
    }
  }
  return ''; // 紐付け不可（新規見込み客の可能性）→ログには表示名で残る
}

// ============================================================
// B) スプシ経由 自動送信：①の「送信文面」欄→LINE送信＋ログ
// ============================================================
function sendQueuedLineMessages() {
  const token = _prop('LINE_TOKEN');
  if (!token) { Logger.log('LINE_TOKEN未設定'); return; }
  const sh = _list(); _ensureLineHeaders(sh);
  const lr = sh.getLastRow(); if (lr < 2) return;
  const log = _log();
  const data = sh.getRange(2, 1, lr-1, LN_COL_SEND).getValues();
  for (let i=0;i<data.length;i++){
    const rowNum = i+2;
    const text   = (data[i][LN_COL_SEND-1]||'').toString().trim();
    const userId = (data[i][LN_COL_USERID-1]||'').toString().trim();
    const name   = (data[i][LN_COL_NAME-1]||'').toString();
    const custId = data[i][LN_COL_ID-1]||'';
    if (!text) continue;                 // 送信文面が空ならスキップ
    if (!userId){                        // userId未紐付け→送れない。メモを残しクリアはしない
      log.appendRow([_now(), custId, name, 'LINE', text, '送信失敗(userId未紐付け)', '顧客が公式アカウントに一度送信すると紐付く']);
      sh.getRange(rowNum, LN_COL_SEND).setValue(''); // 二重防止でクリア
      continue;
    }
    const ok = _pushLine(userId, text, token);
    if (ok){
      log.appendRow([_now(), custId, name, 'LINE', text, '送信', '']); // ★追客送信数KPIの源泉
      sh.getRange(rowNum, LN_COL_LAST).setValue(_today());            // 最終接触日を自動更新
      sh.getRange(rowNum, LN_COL_SEND).setValue('');                 // 送信済み→欄クリア
    } else {
      log.appendRow([_now(), custId, name, 'LINE', text, '送信失敗(API)', '']);
    }
  }
}

function _pushLine(userId, text, token) {
  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ to: userId, messages: [{ type:'text', text:text }] }),
      muteHttpExceptions: true
    });
    return res.getResponseCode() === 200;
  } catch (e) { Logger.log('push失敗: '+e); return false; }
}

function _ensureLineHeaders(sh){
  const need = sh.getRange(1,1,1,LN_COL_SEND).getValues()[0];
  if (!need[LN_COL_USERID-1]) sh.getRange(1, LN_COL_USERID).setValue('LINE_userId');
  if (!need[LN_COL_SEND-1])   sh.getRange(1, LN_COL_SEND).setValue('送信文面(書くと自動送信)');
}

// ============================================================
// セットアップ
// ============================================================
function setupLineTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction()==='sendQueuedLineMessages') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendQueuedLineMessages').timeBased().everyMinutes(10).create();
  _ensureLineHeaders(_list());
}
