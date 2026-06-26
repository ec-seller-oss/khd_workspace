// 【使い捨て・確定修正①】03の2セルを直す＋01司令塔を吐いて次の修正用に読ませる。
// 貼り→ fix1 を実行するだけ。結果はDriveの _khd_dump03.txt をClaudeが読む。
function fix1(){
  var ss = SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var u=shk("売上見込み"), s1=shk("司令");
  var log=[];

  // ① 03!B30(追客→アポ転換率)＝空欄 → 0.5。これでB31/B34/B36の#DIV/0が解消。
  u.getRange("B30").setValue(0.5);
  // ② 03!B45 格付けが固定費額B43を見て"D倒産"誤表示 → 比率D43を見るよう修正。
  u.getRange("B45").setFormula('=IF(D43=0,"-",IF(D43<0.6,"SS 超優良",IF(D43<0.8,"S 優良",IF(D43<0.9,"A 健全",IF(D43<=1,"B 分岐点",IF(D43<=2,"C 赤字","D 倒産"))))))');
  SpreadsheetApp.flush();
  log.push("=== 03修正後の検証(22-47) ===");
  var d=u.getRange(22,1,26,4).getDisplayValues(), f=u.getRange(22,1,26,4).getFormulas();
  for(var r=0;r<26;r++){ log.push("R"+(22+r)+" | A:"+d[r][0]+" | B:"+d[r][1]+(f[r][1]?" {"+f[r][1]+"}":"")+" | C:"+d[r][2]+" | D:"+d[r][3]); }

  // 01司令塔をダンプ(次の修正用)。A〜F列・1〜45行・数式付き。
  log.push("\n=== 01_司令塔 ("+(s1?s1.getName():"なし")+") ===");
  if(s1){
    var n=Math.min(45,s1.getLastRow());
    var dd=s1.getRange(1,1,n,6).getDisplayValues(), ff=s1.getRange(1,1,n,6).getFormulas();
    for(var r=0;r<n;r++){
      log.push("R"+(r+1)+" | A:"+dd[r][0]+" | B:"+dd[r][1]+(ff[r][1]?" {"+ff[r][1]+"}":"")+
               " | C:"+dd[r][2]+(ff[r][2]?" {"+ff[r][2]+"}":"")+" | D:"+dd[r][3]+" | E:"+dd[r][4]+" | F:"+dd[r][5]);
    }
  }

  var txt=log.join("\n"), name="_khd_dump03.txt", it=DriveApp.getFilesByName(name), file;
  if(it.hasNext()){ file=it.next(); file.setContent(txt); } else { file=DriveApp.createFile(name,txt); }
  return file.getUrl();
}
