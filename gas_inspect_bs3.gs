/** 確認用（読み取りのみ）：BS 131行目以降の口座行を調べる 2026-06-03 */
function inspectBS3(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var bs=ss.getSheetByName('③ 資産負債（BS）'); if(!bs) ss.getSheets().forEach(function(s){if(!bs&&s.getName().indexOf('BS')>=0)bs=s;});
  if(!bs) throw new Error('BSタブ無し');
  var lr=bs.getLastRow();
  var rng=bs.getRange(131,1,lr-130,62).getValues();
  var log=['=== BS 131〜'+lr+'行 ==='];
  for(var i=0;i<rng.length;i++){
    var rr=131+i;
    var name=[rng[i][1],rng[i][2],rng[i][3],rng[i][4]].filter(function(x){return x!==''&&x!==null;}).join('／');
    var bh=rng[i][59], bi=rng[i][60];
    if(name!=='' || typeof bi==='number' || typeof bh==='number') log.push('行'+rr+' ['+name+'] BH='+bh+' BI='+bi);
  }
  Logger.log(log.join('\n'));
}
