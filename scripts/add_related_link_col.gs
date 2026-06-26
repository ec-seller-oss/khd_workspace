/**
 * 02_作業DB に「関連資料リンク」列(AB)を追加（右端）。データはこの後 update_db02 で流し込む。
 * 2026-06-17。使い方: script.new → 貼付 → addRelatedLinkColumn 実行（冪等）
 */
function addRelatedLinkColumn(){
  var ID='1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc';
  var ss=SpreadsheetApp.openById(ID);
  var sh=findDb02_(ss); if(!sh) throw '02_作業DB が見つからない';
  var hRow=findHRow_(sh);

  var stamp=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMdd_HHmm');
  var bk='02_作業DB_backup_'+stamp;
  if(!ss.getSheetByName(bk)){ var b=sh.copyTo(ss); b.setName(bk); b.hideSheet(); }

  function H(){ return sh.getRange(hRow,1,1,sh.getLastColumn()).getValues()[0].map(function(x){return String(x).trim();}); }
  function col(name){ var h=H(); for(var i=0;i<h.length;i++) if(h[i]===name) return i+1; return -1; }

  var c=col('関連資料リンク');
  if(c<0){ sh.insertColumnAfter(sh.getLastColumn()); c=sh.getLastColumn(); sh.getRange(hRow,c).setValue('関連資料リンク'); }
  sh.getRange(hRow,c).setBackground('#D9D9D9').setFontWeight('bold')
    .setNote('各作業の関連資料URL（Notion作業ログ/Doc/notes）。=HYPERLINK("url","ラベル") 形式可。Claudeが過去分を突合して充填');

  SpreadsheetApp.flush();
  Logger.log('関連資料リンク 列追加 完了 col='+columnLetter_(c)+' backup='+bk);
  return {ok:true, col:columnLetter_(c), backup:bk};
}

function findDb02_(ss){var s=ss.getSheets();for(var i=0;i<s.length;i++){var r=Math.min(6,s[i].getLastRow());if(r<1)continue;var v=s[i].getRange(1,1,r,Math.min(40,s[i].getLastColumn())).getValues();for(var x=0;x<v.length;x++){var j=v[x].join('|');if(j.indexOf('案件・相手')>=0&&j.indexOf('日付')>=0)return s[i];}}return null;}
function findHRow_(sh){var r=Math.min(6,sh.getLastRow());var v=sh.getRange(1,1,r,Math.min(40,sh.getLastColumn())).getValues();for(var x=0;x<v.length;x++)if(v[x].join('|').indexOf('案件・相手')>=0)return x+1;return 1;}
function columnLetter_(num){var s='';while(num>0){var m=(num-1)%26;s=String.fromCharCode(65+m)+s;num=Math.floor((num-1)/26);}return s;}
