// 【使い捨て・確定修正②】B30に0.5が入らない原因を診断しつつ強制投入。
// 貼り→ fix2 を実行。結果はDrive _khd_dump03.txt をClaudeが読む。
function fix2(){
  var ss = SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var u=shk("売上見込み");
  var c=u.getRange("B30"), log=[];

  // --- 診断（投入前の状態） ---
  log.push("[投入前] B30 value="+JSON.stringify(c.getValue())+" / merged="+c.isPartOfMerge()+
           " / format="+c.getNumberFormat()+" / validation="+(c.getDataValidation()?"あり":"なし"));

  // --- 強制クリア＆投入 ---
  // 1) 結合があれば解除
  if(c.isPartOfMerge()){ c.breakApart(); log.push("→ 結合を解除"); }
  // 2) 入力規則を除去
  c.clearDataValidations();
  // 3) 書式を%にして値を投入
  c.setNumberFormat("0%");
  c.setValue(0.5);
  SpreadsheetApp.flush();

  // --- 診断（投入後） ---
  log.push("[投入後] B30 value="+JSON.stringify(c.getValue())+" / display="+c.getDisplayValue());
  var d=u.getRange(29,2,8,1).getDisplayValues();
  log.push("B29="+d[0][0]+" / B30="+d[1][0]+" / B31="+d[2][0]+" / B32="+d[3][0]+" / B33="+d[4][0]+" / B34="+d[5][0]+" / B35="+d[6][0]+" / B36="+d[7][0]);

  var txt=log.join("\n"), name="_khd_dump03.txt", it=DriveApp.getFilesByName(name), file;
  if(it.hasNext()){ file=it.next(); file.setContent(txt); } else { file=DriveApp.createFile(name,txt); }
  return file.getUrl();
}
