/** 全社売上見込みパイプラインを見やすく整える（B/G/K=④連動は壊さない）2026-06-03 */
function tidyPipeline(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('全社売上見込みパイプライン'); if(!sh) throw new Error('パイプライン無し');
  var YEN='#,##0"円"';
  // ① 使い方（空いてる3行目に明記）
  sh.getRange('A3:M3').merge().setValue(
    '【使い方】B2で「弱気/現実/強気」を選ぶ → 各案件の採用額(K列)が自動で切替 → ④資金繰りの入金に連動。'+
    'E満額×F確度＝期待値(M列)で確度込みの見込みを確認。下部に事業別の採用額合計（=④に流れる額）。')
    .setWrap(true).setBackground('#FBF3D6').setFontWeight('bold');
  // ② ヘッダ整形（4行目）＋期待値列M
  sh.getRange('A4:M4').setFontWeight('bold').setBackground('#AA2E26').setFontColor('#FFFFFF')
    .setWrap(true).setHorizontalAlignment('center');
  sh.getRange('M4').setValue('期待値(満額×確度)');
  for(var r=5;r<=14;r++){ sh.getRange('M'+r).setFormula('=IF(E'+r+'="","",E'+r+'*F'+r+')').setNumberFormat(YEN); }
  // ③ トグルB2・採用額K列を強調
  sh.getRange('B2').setBackground('#FFF4D6').setFontWeight('bold');
  sh.getRange('K4:K14').setBackground('#DDF3DD');
  sh.getRange('F5:F14').setNumberFormat('0%');  // 確度を%表示
  // ④ 下部サマリー（事業別 採用額合計＝④資金繰りに連動する額）
  var sr=16;
  sh.getRange(sr,1,1,3).merge().setValue('■ 事業別 採用額合計（このシナリオで④資金繰りに流れる額）').setFontWeight('bold').setBackground('#F0E2DF');
  var items=[['栄町売却','栄町（物件売却）'],['医療テナントコンサル','医療テナントコンサル'],['テレアポ買取再販','テレアポ買取再販']];
  items.forEach(function(it,i){
    sh.getRange(sr+1+i,1).setValue(it[1]);
    sh.getRange(sr+1+i,2).setFormula('=SUMIF($B$5:$B$14,"'+it[0]+'",$K$5:$K$14)').setNumberFormat(YEN);
  });
  sh.getRange(sr+4,1).setValue('採用額 合計').setFontWeight('bold');
  sh.getRange(sr+4,2).setFormula('=SUM(K5:K14)').setNumberFormat(YEN).setFontWeight('bold').setBackground('#DDF3DD');
  sh.getRange(sr+5,1).setValue('期待値 合計（満額×確度）');
  sh.getRange(sr+5,2).setFormula('=SUM(M5:M14)').setNumberFormat(YEN);
  // ⑤ 整形
  sh.setFrozenRows(4);
  sh.setColumnWidth(1,80); sh.setColumnWidth(2,150); sh.setColumnWidth(3,200); sh.setColumnWidth(4,120);
  sh.setColumnWidth(7,110); sh.setColumnWidth(12,200); sh.setColumnWidth(13,120);
  Logger.log('パイプライン整形完了：使い方/期待値列/トグル強調/事業別サマリー');
  ss.toast('パイプラインを整形（使い方・期待値・事業別合計）。B/G/K連動は維持。','パイプライン整形',6);
}
