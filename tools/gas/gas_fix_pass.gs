/** パス：開始日→A列、期限→出口(F)と統合、末尾L/M列を撤去（既存列を活用）2026-06-03 */
function fixPassLayout(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('パス'); if(!sh) throw new Error('パス無し');
  var lr=sh.getLastRow();
  // 見出し：A=開始日 / F=出口・期限
  sh.getRange('A1').setValue('開始日').setFontWeight('bold');
  sh.getRange('F1').setValue('出口・期限').setFontWeight('bold');
  for(var r=2;r<=lr;r++){
    // L(開始日・前回追加) → A（A空欄のみ）
    var Lv=sh.getRange('L'+r).getValue();
    if(Lv && !sh.getRange('A'+r).getValue()) sh.getRange('A'+r).setValue(Lv);
    // M(解約期限・前回追加) → F(出口)に統合
    var Mv=sh.getRange('M'+r).getValue();
    if(Mv){ var Fv=sh.getRange('F'+r).getValue(); sh.getRange('F'+r).setValue(Fv?(Fv+' ／期限:'+Mv):('期限:'+Mv)); }
  }
  // 末尾L/M列を撤去
  if(lr>=1) sh.getRange('L1:M'+lr).clearContent();
  sh.setColumnWidth(1,90);   // A開始日
  sh.setColumnWidth(6,150);  // F出口・期限
  Logger.log('パス：A=開始日／F=出口・期限へ統合、L/M撤去');
  ss.toast('パス整形：開始日→A列／期限→出口(F)統合／L,M撤去','パス列修正',6);
}
