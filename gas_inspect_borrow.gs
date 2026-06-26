/** 確認用（読取のみ）：⑤借入/返済予定表の構造とBS負債(147-167)のマッピング材料を出す 2026-06-03 */
function inspectBorrow(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var log=['=== 全タブ ==='];
  ss.getSheets().forEach(function(s){ log.push(' ・'+s.getName()+' (r'+s.getLastRow()+' c'+s.getLastColumn()+')'); });
  function L(i){ var s='',n=i; while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s; }
  // ⑤借入タブ（名称に「借入」を含み「返済予定」でないもの優先）
  var b=ss.getSheetByName('⑤ 借入'); if(!b) ss.getSheets().forEach(function(s){ var n=s.getName(); if(!b && n.indexOf('借入')>=0 && n.indexOf('予定')<0) b=s; });
  if(b){
    log.push('=== ⑤借入タブ='+b.getName()+' 先頭18行 ===');
    var lc=Math.min(b.getLastColumn(),30);
    var v=b.getRange(1,1,Math.min(18,b.getLastRow()),lc).getValues();
    for(var r=0;r<v.length;r++){ var cells=[]; for(var c=0;c<lc;c++){ var x=v[r][c]; if(x!==''&&x!==null){ var sv=(x instanceof Date)?('D'+(x.getMonth()+1)):(''+x); if((''+sv).length>14)sv=(''+sv).slice(0,14); cells.push(L(c+1)+(r+1)+'='+sv);} } if(cells.length) log.push(cells.join(' ')); }
  } else log.push('⑤借入タブ見つからず');
  // 0返済予定表（証票）の各借入見出し行を探す
  var z=ss.getSheetByName('0返済予定表'); if(!z) ss.getSheets().forEach(function(s){ if(!z && s.getName().indexOf('返済予定')>=0) z=s; });
  if(z){
    log.push('=== 返済予定表='+z.getName()+' 見出し/列ヘッダ候補 ===');
    var zl=z.getLastRow(), zc=Math.min(z.getLastColumn(),12);
    var zv=z.getRange(1,1,zl,zc).getValues();
    // 借入名らしき見出し行（A〜C列に銀行名）と、ヘッダ(残高/返済/利息)を含む行を抽出
    for(var r2=0;r2<zl;r2++){
      var row=zv[r2], j=row.join('|');
      if(/城北|朝日|大東京|公庫|セゾン|浦安|創業|住宅|MCJ|日本住宅/.test(j) && /借入|返済|予定|銀行|信金|信組|金庫|残高|当初/.test(j)){
        var cs=[]; for(var c2=0;c2<zc;c2++){ if(row[c2]!==''&&row[c2]!==null){ var s2=(row[c2] instanceof Date)?('D'+row[c2].getFullYear()+'/'+(row[c2].getMonth()+1)):(''+row[c2]); cs.push(L(c2+1)+(r2+1)+'='+(''+s2).slice(0,16)); } }
        log.push(cs.join(' '));
      }
    }
  } else log.push('返済予定表見つからず');
  Logger.log(log.join('\n'));
}
