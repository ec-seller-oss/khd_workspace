/**
 * 個人⇔法人 資金移動 按分棚卸し ボタン
 * ----------------------------------------------------
 * 銀行へ試算表/決算書を出す前に押す。
 * 「📋按分棚卸し_テンプレ」を複製し、日付スタンプ付きの
 * 「按分_YYMMDD」タブを生成する。
 *
 * 【設置方法】
 *  1. スプシ（2026_KHD PJ一覧_v2_260601）を開く
 *  2. 拡張機能 > Apps Script
 *  3. このコードを全部貼り付けて保存（フロッピーアイコン）
 *  4. スプシを再読み込み → メニューに「💼帳簿クリーンアップ」が出る
 *  5. （任意）図形ボタンにしたい時は：挿入>図形>図形描画 で四角を置き、
 *     右上「⋮」> スクリプトを割り当て に  createBunpunOroshi  と入力
 */

var TEMPLATE_NAME = '📋按分棚卸し_テンプレ';

/** スプシを開いた時にカスタムメニューを追加 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💼帳簿クリーンアップ')
    .addItem('🔄 按分棚卸しを新規作成', 'createBunpunOroshi')
    .addToParent();
}

/** テンプレを複製して日付つき棚卸しタブを生成 */
function createBunpunOroshi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var tpl = ss.getSheetByName(TEMPLATE_NAME);
  if (!tpl) {
    ui.alert('テンプレ「' + TEMPLATE_NAME + '」が見つかりません。タブ名を確認してください。');
    return;
  }

  // 用途を聞く（例：千葉興銀 試算表提出前）
  var resp = ui.prompt(
    '按分棚卸しを作成',
    '用途を入力（例：千葉興銀 試算表提出前）。空欄でもOK。',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var purpose = resp.getResponseText() || '銀行提出前 棚卸し';

  // 日付スタンプ
  var tz = ss.getSpreadsheetTimeZone() || 'Asia/Tokyo';
  var stamp = Utilities.formatDate(new Date(), tz, 'yyMMdd');
  var dateStr = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd');

  // 一意なタブ名（同日複数回でも衝突しない）
  var base = '按分_' + stamp;
  var name = base;
  var n = 2;
  while (ss.getSheetByName(name)) { name = base + '_' + (n++); }

  // 複製→リネーム→先頭付近へ
  var sheet = tpl.copyTo(ss);
  sheet.setName(name);
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);

  // 日付・用途を書き込み（テンプレと同じ C3 / C4）
  sheet.getRange('C3').setValue(dateStr);
  sheet.getRange('C4').setValue(purpose);

  SpreadsheetApp.flush();
  ss.toast('棚卸し「' + name + '」を作成しました。区分(E列)を全部埋めて、(b)役員貸付を消してから提出。', '✅ 作成完了', 8);
}
