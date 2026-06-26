/**
 * 現役02_作業DB(gid=1226094457)の A日付をM/D書式に統一＋Q予実差分・R達成率に関数を全行再挿入
 * 2026-06-25。使い方: script.new → 貼付 → fixDb02DateQR 実行（冪等）
 */
function fixDb02DateQR(){
  var ss=SpreadsheetApp.openById('1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc');
  var sh=ss.getSheets().filter(function(s){return s.getSheetId()===1226094457;})[0];
  if(!sh) throw '現役02_作業DB(gid1226094457)が見つからない';
  var hRow=1;
  var H=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(x){return String(x).trim();});
  function col(n){for(var i=0;i<H.length;i++)if(H[i]===n)return i+1;return -1;}
  function colL(num){var s='';while(num>0){var m=(num-1)%26;s=String.fromCharCode(65+m)+s;num=Math.floor((num-1)/26);}return s;}
  var last=sh.getLastRow(), n=last-hRow;
  if(n<=0) return;
  var dCol=col('日付'), oCol=col('予定所要(分)'), pCol=col('実所要(分)'), qCol=col('予実差分(分)'), rCol=col('達成率%');

  // 1) A日付：テキスト "YYYY-MM-DD" を日付型に変換し、全体を m/d 表示書式へ
  var dR=sh.getRange(hRow+1, dCol, n, 1);
  var dV=dR.getValues();
  for(var i=0;i<n;i++){
    var v=dV[i][0];
    if(typeof v==='string'){
      var m=v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if(m) dV[i][0]=new Date(+m[1], +m[2]-1, +m[3]);
    }
  }
  dR.setValues(dV);
  dR.setNumberFormat('m"/"d');

  // 2) Q予実差分・R達成率 を全行に関数で（冪等・バッチ）
  if(oCol>0 && pCol>0 && qCol>0 && rCol>0){
    var oL=colL(oCol), pL=colL(pCol), qF=[], rF=[];
    for(var rr=hRow+1; rr<=last; rr++){
      qF.push(['=IF(OR('+oL+rr+'="",'+pL+rr+'=""),"",'+oL+rr+'-'+pL+rr+')']);
      rF.push(['=IF(OR('+pL+rr+'="",'+pL+rr+'=0,'+oL+rr+'=""),"",ROUND('+pL+rr+'/'+oL+rr+'*100,1))']);
    }
    sh.getRange(hRow+1, qCol, n, 1).setFormulas(qF);
    sh.getRange(hRow+1, rCol, n, 1).setFormulas(rF);
  }
  SpreadsheetApp.flush();
  Logger.log('fixDb02DateQR 完了 / 行数='+n+' / 日付列='+colL(dCol)+' QR='+colL(qCol)+colL(rCol));
}
