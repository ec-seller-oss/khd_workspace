/** 確認用（読み取りのみ）：BSの口座残高入力箇所を調べる 2026-06-03
 * 現在値列=BI(61)想定。各口座ラベル(A列等)とBI値、近接の日付/前回列を洗い出し、残高クイック入力の設計に使う。 */
function inspectBS(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var bs=null;
  ss.getSheets().forEach(function(s){ var n=s.getName(); if(!bs && (n==='BS'||n.indexOf('BS')>=0||n.indexOf('資産')>=0||n.indexOf('収支')>=0)) bs=s; });
  if(!bs) throw new Error('BS/収支タブが見つからない。タブ名一覧='+ss.getSheets().map(function(s){return s.getName();}).join(','));
  var lr=bs.getLastRow(), lc=bs.getLastColumn();
  function L(i){ var s='',n=i; while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s; }
  var log=['=== BSタブ='+bs.getName()+' lastRow='+lr+' lastCol='+lc+'('+L(lc)+') ==='];
  // ヘッダ3行（列見出し把握）
  var head=bs.getRange(1,1,Math.min(3,lr),lc).getValues();
  for(var h=0;h<head.length;h++){ var hs=[]; for(var c=0;c<lc;c++){ var v=head[h][c]; if(v!==''&&v!==null){ var sv=(v instanceof Date)?('D:'+(v.getMonth()+1)+'月'):(''+v); hs.push(L(c+1)+'='+sv);} } if(hs.length) log.push('[ヘッダ行'+(h+1)+'] '+hs.slice(0,30).join(' | ')); }
  // 残高らしき行：A〜D列にラベルがあり、BI(61)付近に数値がある行
  var data=bs.getRange(1,1,lr,Math.min(lc,62)).getValues();
  log.push('--- 口座/残高行（ラベル＋現在値列付近）---');
  var cnt=0;
  for(var r=0;r<lr;r++){
    var label=''; for(var lcl=0;lcl<6&&lcl<lc;lcl++){ if(data[r][lcl]!==''&&data[r][lcl]!==null){ label=L(lcl+1)+(r+1)+':'+data[r][lcl]; break; } }
    var bi=(lc>=61)?data[r][60]:'';   // BI
    var bh=(lc>=60)?data[r][59]:'';   // BH(前回候補)
    if(label && (typeof bi==='number' || typeof bh==='number')){
      log.push(label+'  | BH'+(r+1)+'='+bh+' | BI'+(r+1)+'='+bi);
      cnt++; if(cnt>=60) { log.push('…(以下略)'); break; }
    }
  }
  Logger.log(log.join('\n'));
}
