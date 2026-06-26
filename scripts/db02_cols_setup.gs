/**
 * 02_作業DB 列セットアップ（1回で2つ）：
 *  1) 「報告相手」プルダウンを9択に更新（福井/羽鳥/宮崎/銀行/不動産屋/業者/バイセル/そうけん/ゆーし）
 *  2) 「関連資料リンク」列(AB)を新設（この後Claudeがリンクを流し込む）
 * 2026-06-17。使い方: script.new → 貼付 → setupDb02Cols 実行（冪等）
 */
function setupDb02Cols(){
  var ID='1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc';
  var ss=SpreadsheetApp.openById(ID);
  var sh=findDb02_(ss); if(!sh) throw '02_作業DB が見つからない';
  var hRow=findHRow_(sh);

  var stamp=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMdd_HHmm');
  var bk='02_作業DB_backup_'+stamp;
  if(!ss.getSheetByName(bk)){ var b=sh.copyTo(ss); b.setName(bk); b.hideSheet(); }

  function col(name){ var h=sh.getRange(hRow,1,1,sh.getLastColumn()).getValues()[0].map(function(x){return String(x).trim();}); for(var i=0;i<h.length;i++) if(h[i]===name) return i+1; return -1; }

  // 入力枠確保
  var need=hRow+60;
  if(sh.getMaxRows()<need) sh.insertRowsAfter(sh.getMaxRows(), need-sh.getMaxRows());

  // 1) 報告相手 9択プルダウン
  var c=col('報告相手'); if(c<0) throw '報告相手 列が無い';
  var rule=SpreadsheetApp.newDataValidation()
    .requireValueInList(['福井','羽鳥','宮崎','銀行','不動産屋','業者','バイセル','そうけん','ゆーし'], true)
    .setAllowInvalid(true).build();
  sh.getRange(hRow+1, c, need-hRow, 1).setDataValidation(rule);
  sh.getRange(hRow,c).setNote('🟡プルダウン9択。誰向けの報告か。\n福井/羽鳥/宮崎/銀行/不動産屋/業者/バイセル/そうけん/ゆーし\n※羽鳥(クラウドミル)とバイセルは別人格');

  // 2) 関連資料リンク列(AB)
  var l=col('関連資料リンク');
  if(l<0){ sh.insertColumnAfter(sh.getLastColumn()); l=sh.getLastColumn(); sh.getRange(hRow,l).setValue('関連資料リンク'); }
  sh.getRange(hRow,l).setBackground('#D9D9D9').setFontWeight('bold')
    .setNote('各作業の関連資料URL（Notion作業ログ/Doc/notes）。Claudeが過去分を突合して充填');

  SpreadsheetApp.flush();
  Logger.log('setupDb02Cols 完了 / 報告相手9択 col='+c+' / 関連資料リンク col='+l+' / backup='+bk);
  return {ok:true, aite_col:c, link_col:l, backup:bk};
}
function findDb02_(ss){var s=ss.getSheets();for(var i=0;i<s.length;i++){var r=Math.min(6,s[i].getLastRow());if(r<1)continue;var v=s[i].getRange(1,1,r,Math.min(40,s[i].getLastColumn())).getValues();for(var x=0;x<v.length;x++){var j=v[x].join('|');if(j.indexOf('案件・相手')>=0&&j.indexOf('日付')>=0)return s[i];}}return null;}
function findHRow_(sh){var r=Math.min(6,sh.getLastRow());var v=sh.getRange(1,1,r,Math.min(40,sh.getLastColumn())).getValues();for(var x=0;x<v.length;x++)if(v[x].join('|').indexOf('案件・相手')>=0)return x+1;return 1;}
