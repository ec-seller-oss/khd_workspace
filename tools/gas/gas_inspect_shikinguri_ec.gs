/** 確認用（読み取りのみ・書込なし）：④資金繰りのEC粗利行の配置を調べる 2026-06-03
 * クーパン1/2 の行・値が入っている列・月ヘッダ行を特定して、STEP2の配線方法を決める。 */
function inspectShikinguriEC(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var cf=null;
  ss.getSheets().forEach(function(s){ if(!cf && s.getName().indexOf('資金繰り')>=0) cf=s; });
  if(!cf) throw new Error('資金繰りタブが見つからない');
  var lr=cf.getLastRow(), lc=cf.getLastColumn();
  var vals=cf.getRange(1,1,lr,lc).getValues();
  function L(i){ var s='',n=i+1; while(n>0){ var m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26);} return s; }
  var log=['=== ④'+cf.getName()+' lastRow='+lr+' lastCol='+lc+' ('+L(lc-1)+') ==='];
  // 1) クーパン行を探して、その行の非空セルを列付きで出す
  for(var r=0;r<lr;r++){
    var joined=vals[r].join('');
    if(joined.indexOf('クーパン')>=0){
      var cells=[];
      for(var c=0;c<lc;c++){ var v=vals[r][c]; if(v!==''&&v!==null) cells.push(L(c)+(r+1)+'='+v); }
      log.push('[行'+(r+1)+'] '+cells.join(' | '));
    }
  }
  // 2) 月ヘッダ候補行（日付/「月」を多く含む行）を探す
  for(var r2=0;r2<Math.min(lr,8);r2++){
    var cnt=0, sample=[];
    for(var c2=0;c2<lc;c2++){ var v2=vals[r2][c2]; var sv=(v2 instanceof Date)?('DATE:'+v2.getFullYear()+'/'+(v2.getMonth()+1)):(''+v2); if(/月|^20\d\d|\/\d/.test(sv)&&v2!==''){ cnt++; if(sample.length<14) sample.push(L(c2)+(r2+1)+'='+sv);} }
    if(cnt>=3) log.push('[月ヘッダ候補 行'+(r2+1)+' 該当'+cnt+'] '+sample.join(' | '));
  }
  Logger.log(log.join('\n'));
}
