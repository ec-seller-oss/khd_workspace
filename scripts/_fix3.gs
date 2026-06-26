// 【使い捨て・確定修正③v2】B30周りの結合を範囲ごと強制解除→0.5投入。getMergedRanges不使用。
// 貼り→ fix3 を実行。結果はDrive _khd_dump03.txt。
function fix3(){
  var ss = SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var u=shk("売上見込み"), log=[];

  // 結合を範囲ごと解除（横結合→広め→縦も含む の順でtry）
  var tries=["A30:F30","A28:F32","A22:F37"];
  for(var i=0;i<tries.length;i++){ try{ u.getRange(tries[i]).breakApart(); log.push("breakApart "+tries[i]+" OK"); }catch(e){ log.push("breakApart "+tries[i]+" 失敗:"+e.message); } }

  // 入力規則除去＋書式%＋0.5投入
  var c=u.getRange("B30");
  c.clearDataValidations(); c.setNumberFormat("0%"); c.setValue(0.5);
  SpreadsheetApp.flush();

  // 検証
  log.push("[投入後] B30 value="+JSON.stringify(c.getValue())+" / display="+c.getDisplayValue());
  var d=u.getRange(29,2,8,1).getDisplayValues();
  log.push("B29="+d[0][0]+" B30="+d[1][0]+" B31="+d[2][0]+" B32="+d[3][0]+" B33="+d[4][0]+" B34="+d[5][0]+" B35="+d[6][0]+" B36="+d[7][0]);

  var txt=log.join("\n"), name="_khd_dump03.txt", it=DriveApp.getFilesByName(name), file;
  if(it.hasNext()){ file=it.next(); file.setContent(txt); } else { file=DriveApp.createFile(name,txt); }
  return file.getUrl();
}
