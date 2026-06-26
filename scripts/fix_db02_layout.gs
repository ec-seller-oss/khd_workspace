/**
 * 02_作業DB 仕上げ整形：（未使用）列を削除して詰める＋顧客番号/案件番号を案件・相手(D)の右へ移動
 * 2026-06-16。cleanDb02Safe の後に実行する。
 * 使い方: script.new → 全消し → 全貼り → fixDb02Layout 実行（冪等）
 *
 * - 列の移動・削除はGASでもSheetsが参照を自動追従するため、他タブの連動は壊れない。
 * - 実行前にバックアップ（隠しシート）を必ず複製。
 */
function fixDb02Layout(){
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

  // 1) 顧客番号・案件番号 を 案件・相手(D)の直後へ移動
  //    案件番号→D後、顧客番号→D後 の順で入れる → D | 顧客番号 | 案件番号 の並びになる
  var d=col('案件・相手'); if(d<0) throw '案件・相手 が無い';
  var ck=col('案件番号'); if(ck>0) sh.moveColumns(sh.getRange(1,ck,1,1), d+1);
  var cc=col('顧客番号'); if(cc>0) sh.moveColumns(sh.getRange(1,cc,1,1), d+1);

  // 2) 「未使用」を含むヘッダの列を削除（右から左へ＝インデックスずれ防止）
  var h=H(); var del=[];
  for(var i=0;i<h.length;i++) if(h[i].indexOf('未使用')>=0) del.push(i+1);
  del.sort(function(a,b){return b-a;});
  del.forEach(function(c){ sh.deleteColumn(c); });

  SpreadsheetApp.flush();
  Logger.log('fixDb02Layout 完了 / backup='+bk+' / 削除列数='+del.length+' / 管理番号をD列の右へ移動');
  return {ok:true, backup:bk, deleted:del.length};
}

// ── helpers ──
function findDb02_(ss){var s=ss.getSheets();for(var i=0;i<s.length;i++){var r=Math.min(6,s[i].getLastRow());if(r<1)continue;var v=s[i].getRange(1,1,r,Math.min(40,s[i].getLastColumn())).getValues();for(var x=0;x<v.length;x++){var j=v[x].join('|');if(j.indexOf('案件・相手')>=0&&j.indexOf('日付')>=0)return s[i];}}return null;}
function findHRow_(sh){var r=Math.min(6,sh.getLastRow());var v=sh.getRange(1,1,r,Math.min(40,sh.getLastColumn())).getValues();for(var x=0;x<v.length;x++)if(v[x].join('|').indexOf('案件・相手')>=0)return x+1;return 1;}
