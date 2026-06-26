/** ③BSの旧・予定PLブロック(④資金繰り/①司令塔へ移行済)の#REFを撤去 2026-06-03 */
function cleanBsForecast(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('③ 資産負債（BS）');
  if(!sh){ throw new Error('③ 資産負債（BS）が見つかりません'); }
  // 予定PL行(売上/原価/経費/利益)の予測列 AR4:BJ7 のみ撤去。実績(BI50〜)・ラベルは無傷。
  sh.getRange('AR4:BJ7').clearContent();
  Logger.log('③BS AR4:BJ7（旧予定PLブロック・#REF）を撤去');
  ss.toast('BS上部の壊れたPL行(#REF×5)を撤去。PLは①司令塔へ集約予定。','完了',6);
}
