// 【使い捨て・番地確定用】03のA〜E列(1〜60行)をDriveのテキストに書き出す。
// 貼り→ dump03 を実行するだけ。あとはClaudeがDriveから直接読むのでコピペ不要。
function dump03(){
  var ss = SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  var sh = null, a = ss.getSheets();
  for (var i=0;i<a.length;i++){ if(a[i].getName().indexOf("売上見込み")>=0){ sh=a[i]; break; } }
  if(!sh){ return "03が見つからない"; }
  var disp = sh.getRange(1,1,60,5).getDisplayValues();
  var form = sh.getRange(1,1,60,5).getFormulas();
  var out = [];
  for(var r=0;r<60;r++){
    out.push("R"+(r+1)+" | A:"+disp[r][0]+" | B:"+disp[r][1]+(form[r][1]?" {"+form[r][1]+"}":"")+
             " | C:"+disp[r][2]+" | D:"+disp[r][3]+(form[r][3]?" {"+form[r][3]+"}":"")+" | E:"+disp[r][4]);
  }
  var txt = out.join("\n");
  // 固定ファイル名で上書き(無ければ新規)
  var name = "_khd_dump03.txt", file=null, it=DriveApp.getFilesByName(name);
  if(it.hasNext()){ file=it.next(); file.setContent(txt); } else { file=DriveApp.createFile(name, txt); }
  return file.getUrl();
}
