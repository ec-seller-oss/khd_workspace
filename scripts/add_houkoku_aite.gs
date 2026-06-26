/**
 * 02_作業DB に「報告相手」列を追加し、プルダウン（福井/羽鳥/宮崎/銀行/不動産屋/業者）を設定
 * 2026-06-17。報告値Uの直後に挿入（報告項目|報告値|報告相手 のグループ）。
 * 使い方: script.new → 全消し → 全貼り → addHoukokuAite 実行（冪等＝再実行で列は増えない）
 */
function addHoukokuAite(){
  var ID='1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc';
  var ss=SpreadsheetApp.openById(ID);
  var sh=findDb02_(ss); if(!sh) throw '02_作業DB が見つからない';
  var hRow=findHRow_(sh);

  // 0) バックアップ
  var stamp=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMdd_HHmm');
  var bk='02_作業DB_backup_'+stamp;
  if(!ss.getSheetByName(bk)){ var b=sh.copyTo(ss); b.setName(bk); b.hideSheet(); }

  function H(){ return sh.getRange(hRow,1,1,sh.getLastColumn()).getValues()[0].map(function(x){return String(x).trim();}); }
  function col(name){ var h=H(); for(var i=0;i<h.length;i++) if(h[i]===name) return i+1; return -1; }

  // 1) 「報告相手」列（無ければ報告値の直後に挿入。冪等）
  var c=col('報告相手');
  if(c<0){
    var hv=col('報告値');
    var after = hv>0 ? hv : sh.getLastColumn();
    sh.insertColumnAfter(after);
    c=after+1;
    sh.getRange(hRow,c).setValue('報告相手');
  }
  sh.getRange(hRow,c).setBackground('#FFF2CC').setFontWeight('bold')
    .setNote('🟡プルダウン。誰向けの報告か＝報告生成のフィルタキー。福井/羽鳥/宮崎/銀行/不動産屋/業者');

  // 2) 入力枠を確保（不足なら行を足す）
  var need=hRow+50;
  if(sh.getMaxRows()<need) sh.insertRowsAfter(sh.getMaxRows(), need-sh.getMaxRows());

  // 3) プルダウン（データ検証）を設定
  var rule=SpreadsheetApp.newDataValidation()
    .requireValueInList(['福井','羽鳥','宮崎','銀行','不動産屋','業者'], true)
    .setAllowInvalid(true)   // 6択以外の手入力も許容（将来 家族/税理士 等の拡張余地）
    .build();
  sh.getRange(hRow+1, c, need-hRow, 1).setDataValidation(rule);

  SpreadsheetApp.flush();
  Logger.log('報告相手 追加＆プルダウン設定 完了 / col='+columnLetter_(c)+' / backup='+bk);
  return {ok:true, col:columnLetter_(c), backup:bk};
}

// ── helpers ──
function findDb02_(ss){var s=ss.getSheets();for(var i=0;i<s.length;i++){var r=Math.min(6,s[i].getLastRow());if(r<1)continue;var v=s[i].getRange(1,1,r,Math.min(40,s[i].getLastColumn())).getValues();for(var x=0;x<v.length;x++){var j=v[x].join('|');if(j.indexOf('案件・相手')>=0&&j.indexOf('日付')>=0)return s[i];}}return null;}
function findHRow_(sh){var r=Math.min(6,sh.getLastRow());var v=sh.getRange(1,1,r,Math.min(40,sh.getLastColumn())).getValues();for(var x=0;x<v.length;x++)if(v[x].join('|').indexOf('案件・相手')>=0)return x+1;return 1;}
function columnLetter_(num){var s='';while(num>0){var m=(num-1)%26;s=String.fromCharCode(65+m)+s;num=Math.floor((num-1)/26);}return s;}
