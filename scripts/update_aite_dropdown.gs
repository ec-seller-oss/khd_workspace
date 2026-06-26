/**
 * 02_作業DB「報告相手」プルダウンを9択に更新（バイセル/そうけん/ゆーし を追加）
 * 2026-06-17。使い方: script.new → 貼付 → updateAiteDropdown 実行（冪等）
 */
function updateAiteDropdown(){
  var ID='1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc';
  var ss=SpreadsheetApp.openById(ID);
  var sh=findDb02_(ss); if(!sh) throw '02_作業DB が見つからない';
  var hRow=findHRow_(sh);
  function col(name){ var h=sh.getRange(hRow,1,1,sh.getLastColumn()).getValues()[0].map(function(x){return String(x).trim();}); for(var i=0;i<h.length;i++) if(h[i]===name) return i+1; return -1; }
  var c=col('報告相手'); if(c<0) throw '報告相手 列が無い';

  var need=hRow+60;
  if(sh.getMaxRows()<need) sh.insertRowsAfter(sh.getMaxRows(), need-sh.getMaxRows());

  var rule=SpreadsheetApp.newDataValidation()
    .requireValueInList(['福井','羽鳥','宮崎','銀行','不動産屋','業者','バイセル','そうけん','ゆーし'], true)
    .setAllowInvalid(true).build();
  sh.getRange(hRow+1, c, need-hRow, 1).setDataValidation(rule);
  sh.getRange(hRow,c).setNote('🟡プルダウン9択。誰向けの報告か。福井/羽鳥/宮崎/銀行/不動産屋/業者/バイセル/そうけん/ゆーし\n※羽鳥(クラウドミル)とバイセルは別人格');

  SpreadsheetApp.flush();
  Logger.log('報告相手 プルダウン9択に更新 完了 col='+c);
  return {ok:true};
}
function findDb02_(ss){var s=ss.getSheets();for(var i=0;i<s.length;i++){var r=Math.min(6,s[i].getLastRow());if(r<1)continue;var v=s[i].getRange(1,1,r,Math.min(40,s[i].getLastColumn())).getValues();for(var x=0;x<v.length;x++){var j=v[x].join('|');if(j.indexOf('案件・相手')>=0&&j.indexOf('日付')>=0)return s[i];}}return null;}
function findHRow_(sh){var r=Math.min(6,sh.getLastRow());var v=sh.getRange(1,1,r,Math.min(40,sh.getLastColumn())).getValues();for(var x=0;x<v.length;x++)if(v[x].join('|').indexOf('案件・相手')>=0)return x+1;return 1;}
