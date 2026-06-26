/**
 * 物件メール本文から健美家/楽待のlisting URLを抽出し、
 * 物件マッチング一覧スプシの O列「物件URL」へ書き出す。
 * → ローカルの auto_pipeline.py --sheet が O列を読んで本査定→🟢通知。
 *
 * 使い方:
 *  A) 既存 gas_property_matching.gs の不動産行 appendRow に URL を15列目で足す
 *     （下の extractPropertyUrl_ を呼ぶ）
 *  B) もしくは backfillPropertyUrls を単体実行/トリガーで既存行に後埋め
 *
 * 2026-06-02 KHD AI査定エンジン無人化の最終配線
 */

var MATCH_SS_ID = '1a0w6K-fi_BpTGGAVmB1lHqAJYPjnM4M8fw8Rs25ghnc';
var MATCH_TAB   = '物件マッチング一覧（医療テナント）';

/** メール本文から最初の健美家/楽待 物件URLを返す（無ければ ''） */
function extractPropertyUrl_(body) {
  if (!body) return '';
  // 健美家: /pp2/s/.../re_xxxx/  楽待: /syuuekibukken/.../show.html
  var re = /(https?:\/\/(?:www\.)?(?:kenbiya\.com\/pp2\/s\/[^\s"'<>?]+|rakumachi\.jp\/syuuekibukken\/[^\s"'<>?]+show\.html))/i;
  var m = body.match(re);
  if (!m) return '';
  return m[1].split('?')[0]; // utm等のクエリを除去
}

/**
 * 既存の不動産行（O列が空）に物件URLを後埋め。
 * 送信元 kenbiya/楽待系の直近メールを走査し、件名一致でO列へ書く。
 */
function backfillPropertyUrls() {
  var ss = SpreadsheetApp.openById(MATCH_SS_ID);
  var sh = ss.getSheetByName(MATCH_TAB);
  var last = sh.getLastRow();
  var data = sh.getRange(1, 1, last, 15).getValues(); // A..O

  // 直近の物件メールを取得（健美家/楽待/不動産業者）
  var threads = GmailApp.search('(from:kenbiya.com OR from:rakumachi.jp OR label:不動産) newer_than:7d', 0, 80);
  var subjToUrl = {};
  threads.forEach(function (t) {
    t.getMessages().forEach(function (msg) {
      var url = extractPropertyUrl_(msg.getPlainBody() || msg.getBody());
      if (url) {
        var s = (msg.getSubject() || '').replace(/\s/g, '').slice(0, 18);
        if (s) subjToUrl[s] = url;
      }
    });
  });

  var wrote = 0;
  for (var r = 1; r < data.length; r++) {           // 0=ヘッダ
    if ((data[r][1] || '') !== '不動産') continue;   // B列=種別
    if (data[r][14]) continue;                       // O列(15)既にURL有
    var name = (data[r][3] || '').replace(/\s/g, '').slice(0, 18); // D列=物件名(件名)
    var url = subjToUrl[name];
    if (url) {
      sh.getRange(r + 1, 15).setValue(url);          // O列へ
      wrote++;
    }
  }
  Logger.log('物件URL後埋め: ' + wrote + '件');
  try { ss.toast(wrote + '件にlisting URLを書き出しました', '物件URL後埋め 完了', 6); } catch (e) {}  // トリガー実行時はUI通知不可→無視
}

/** O列ヘッダを用意（初回のみ手動実行） */
function setupUrlColumn() {
  var sh = SpreadsheetApp.openById(MATCH_SS_ID).getSheetByName(MATCH_TAB);
  if (!sh.getRange(1, 15).getValue()) sh.getRange(1, 15).setValue('物件URL');
}
