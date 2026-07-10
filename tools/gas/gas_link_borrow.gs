/** ④資金繰りの借入返済を⑤借入(証票・月次実額)へ連動＝借入SSoT 2026-06-03 */
function linkBorrowToTab(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('④ 資金繰り'); if(!sh) throw new Error('④資金繰り無し');
  var T="'⑤ 借入'";
  // ⑤借入 月返済額列: 城北F/朝日M/大東京T(法人)・TB創業AA/TBセゾンAH/公庫AV(個人)・住宅AO
  // 月行: 6月=r9, 7月=r10 … 1月=r16（④の列 B..I に対応）
  for(var i=0;i<8;i++){
    var col=String.fromCharCode(66+i); // B..I
    var br=9+i;                        // ⑤借入の月行
    // ④r22 法人借入返済 = 城北+朝日+大東京
    sh.getRange(col+'22').setFormula('='+T+'!F'+br+'+'+T+'!M'+br+'+'+T+'!T'+br).setNumberFormat('#,##0"円"');
    // ④r23 個人事業借入返済 = TB創業+TBセゾン+公庫
    sh.getRange(col+'23').setFormula('='+T+'!AA'+br+'+'+T+'!AH'+br+'+'+T+'!AV'+br).setNumberFormat('#,##0"円"');
    // ④r18 住宅ローン = 住宅MCJ
    sh.getRange(col+'18').setFormula('='+T+'!AO'+br).setNumberFormat('#,##0"円"');
  }
  Logger.log('④借入返済(r18住宅/r22法人/r23個人)を⑤借入の月次実額へ連動');
  ss.toast('借入返済を⑤借入(証票)連動に。月次実額・自動逓減。','借入SSoT連動',6);
}
