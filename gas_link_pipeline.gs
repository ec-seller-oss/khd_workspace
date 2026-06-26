/** ④資金繰りの入金(栄町/医療/テレアポ)をパイプラインにSUMIFS連動＝売上SSoT復元 2026-06-03 */
function linkShikinguriToPipeline(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('④ 資金繰り'); if(!sh) throw new Error('④資金繰り無し');
  var P="'全社売上見込みパイプライン'";
  // ④の入金行 → パイプラインの科目(B列)
  var MAP={5:'栄町売却', 6:'医療テナントコンサル', 7:'テレアポ買取再販'};
  // 月列 B..I = 2026/6 .. 2027/1
  var YM=[[2026,6],[2026,7],[2026,8],[2026,9],[2026,10],[2026,11],[2026,12],[2027,1]];
  var n=0;
  for(var r in MAP){
    for(var i=0;i<YM.length;i++){
      var col=String.fromCharCode(66+i); // B..I
      var fml='=SUMIFS('+P+'!$K:$K,'+P+'!$B:$B,"'+MAP[r]+'",'+P+'!$G:$G,DATE('+YM[i][0]+','+YM[i][1]+',1))';
      sh.getRange(col+r).setFormula(fml).setNumberFormat('#,##0"円"');
      n++;
    }
  }
  Logger.log('④入金'+n+'セルをパイプラインSUMIFSへ連動（栄町/医療/テレアポ・採用額K×着金月G）');
  ss.toast('④資金繰りの入金をパイプライン連動に復元。シナリオ切替で自動更新。','売上SSoT連動',6);
}
