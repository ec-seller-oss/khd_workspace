/** 確認用（読取のみ）：BS集計行が「列ごとの数式」か／月ヘッダ／④資金繰りの月初現金参照を調べる 2026-06-03
 * 日次スナップショット列を安全に挿入するための事前調査。 */
function inspectBSFormula(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var bs=ss.getSheetByName('③ 資産負債（BS）'); if(!bs) ss.getSheets().forEach(function(s){if(!bs&&s.getName().indexOf('BS')>=0)bs=s;});
  if(!bs) throw new Error('BSタブ無し');
  function L(i){ var s='',n=i; while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s; }
  var log=['=== BS 列ごと数式調査 lastCol='+bs.getLastColumn()+'('+L(bs.getLastColumn())+') ==='];
  // 行2 月ヘッダ BF〜BL
  var h=bs.getRange(2,58,1,7).getValues()[0];
  var hs=[]; for(var c=0;c<7;c++){ var v=h[c]; hs.push(L(58+c)+'='+((v instanceof Date)?(v.getFullYear()+'/'+(v.getMonth()+1)+'/'+v.getDate()):v)); }
  log.push('月ヘッダ行2: '+hs.join(' '));
  // 集計行(11,12,36,37,40,44,45)のBH/BIの数式（列ごと数式か値か）
  [11,12,36,37,40,44,45].forEach(function(r){
    var fBH=bs.getRange(r,60).getFormula(), fBI=bs.getRange(r,61).getFormula();
    var vBH=bs.getRange(r,60).getValue(), vBI=bs.getRange(r,61).getValue();
    var lab=bs.getRange(r,4).getValue()||bs.getRange(r,1).getValue();
    log.push('行'+r+' ['+lab+'] BH:'+(fBH||('値'+vBH))+'  ||  BI:'+(fBI||('値'+vBI)));
  });
  // 明細セル(BI50)の中身（値か数式か）
  log.push('明細 BI50 formula:['+bs.getRange(50,61).getFormula()+'] value:'+bs.getRange(50,61).getValue());
  // 資金繰り/収支 が BS の月列を参照しているか
  var cf=null; ss.getSheets().forEach(function(s){if(!cf&&s.getName().indexOf('資金繰り')>=0)cf=s;});
  if(cf){ log.push('=== '+cf.getName()+' 月初現金まわりの数式(B列1-12行) ===');
    var fb=cf.getRange(1,2,12,1).getFormulas();
    for(var i=0;i<12;i++){ if(fb[i][0]) log.push(' B'+(i+1)+': '+fb[i][0]); } }
  Logger.log(log.join('\n'));
}
