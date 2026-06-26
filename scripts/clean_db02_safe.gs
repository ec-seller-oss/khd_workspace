/**
 * 02_作業DB 安全クリーン（列位置を動かさない＝他タブの '02_作業DB'! 列参照を壊さない）
 * 2026-06-16 作成。Plan A。
 * 使い方: script.new → 全消し → 全貼り → cleanDb02Safe 実行（冪等・何度でも可）
 *
 * やること:
 *  0) バックアップ：'02_作業DB_backup_yyyyMMdd_HHmm' を複製（破壊前に必ず退避）
 *  1) 重複「内容」名寄せ：2列目の非空を1列目へ結合 → 2列目を（未使用）化＆クリア
 *  2) 残骸「報告」名寄せ：非空を「結果」へ結合 → （未使用）化＆クリア
 *  3) 旧「実所要分」を「実所要(分)」自動列に一本化 → ラベル無効化（値はbackupに保全）
 *  4) ゴミ尾削除：最終実データ行を検出し、それ以降の空行を物理削除（'—'/'-'のみは空扱い）
 *  5) 03_売上見込み の確度/着金 配線を、現在の「報告項目」「報告値」列へ修復
 *  ※ 既存列は1本も削除・移動しない。（未使用）列はスロットを残置＝他タブの列文字参照を保護。
 */
function cleanDb02Safe(){
  var ID='1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc';
  var ss=SpreadsheetApp.openById(ID);
  var sh=findDb02_(ss); if(!sh) throw '02_作業DB が見つからない';
  var hRow=findHRow_(sh);

  // 0) バックアップ
  var stamp=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMdd_HHmm');
  var bkName='02_作業DB_backup_'+stamp;
  if(!ss.getSheetByName(bkName)){ var bk=sh.copyTo(ss); bk.setName(bkName); bk.hideSheet(); }

  var lastCol=sh.getLastColumn(), lastRow=sh.getLastRow();
  var H=sh.getRange(hRow,1,1,lastCol).getValues()[0].map(function(x){return String(x).trim();});
  function idxsOf(name){var a=[];for(var i=0;i<H.length;i++)if(H[i]===name)a.push(i);return a;}
  function idx1(name){var a=idxsOf(name);return a.length?a[0]:-1;}
  function colL(i0){return columnLetter_(i0+1);} // 0-based index → 列文字
  var n=Math.max(0,lastRow-hRow);
  var log=[];

  // 1) 重複「内容」名寄せ
  var ic=idxsOf('内容');
  if(ic.length>=2 && n>0){
    var a=sh.getRange(hRow+1,ic[0]+1,n,1).getValues();
    var b=sh.getRange(hRow+1,ic[1]+1,n,1).getValues();
    var merged=0;
    for(var r=0;r<n;r++){
      var x=String(a[r][0]||'').trim(), y=String(b[r][0]||'').trim();
      var v=(x&&y&&x!==y)?(x+' / '+y):(x||y);
      if(v!==x){ a[r][0]=v; if(y) merged++; }
    }
    sh.getRange(hRow+1,ic[0]+1,n,1).setValues(a);
    sh.getRange(hRow+1,ic[1]+1,n,1).clearContent();
    sh.getRange(hRow,ic[1]+1).setValue('（未使用）').setBackground('#D9D9D9')
      .setNote('旧・重複「内容」。1列目へ名寄せ後に空に。列は他タブの参照保護のため残置');
    log.push('内容重複 名寄せ '+merged+'件');
  }

  // 2) 残骸「報告」→「結果」名寄せ
  var ir=idx1('報告'), ik=idx1('結果');
  if(ir>=0 && ik>=0 && n>0){
    var rp=sh.getRange(hRow+1,ir+1,n,1).getValues();
    var kk=sh.getRange(hRow+1,ik+1,n,1).getValues();
    var m2=0;
    for(var r2=0;r2<n;r2++){
      var px=String(rp[r2][0]||'').trim(), ky=String(kk[r2][0]||'').trim();
      if(px){ kk[r2][0]=ky?(ky+' / '+px):px; m2++; }
    }
    sh.getRange(hRow+1,ik+1,n,1).setValues(kk);
    sh.getRange(hRow+1,ir+1,n,1).clearContent();
    sh.getRange(hRow,ir+1).setValue('（未使用）').setBackground('#D9D9D9')
      .setNote('旧「報告」残骸。結果へ名寄せ後に空に。列残置');
    log.push('報告→結果 名寄せ '+m2+'件');
  }

  // 3) 旧「実所要分」を無効化（自動「実所要(分)」に一本化。値はbackup保全）
  var im=idx1('実所要分');
  if(im>=0){
    sh.getRange(hRow,im+1).setValue('（旧)実所要分_未使用').setBackground('#D9D9D9')
      .setNote('自動「実所要(分)」に一本化。手入力値はbackupに保全。列残置');
    log.push('実所要分 無効化');
  }

  // 4) ゴミ尾削除（最終実データ行の検出）
  var lastReal=hRow;
  if(n>0){
    var vals=sh.getRange(hRow+1,1,n,lastCol).getValues();
    for(var r3=0;r3<n;r3++){
      var row=vals[r3], has=false;
      for(var c=0;c<row.length;c++){var s=String(row[c]).trim(); if(s&&s!=='—'&&s!=='-'){has=true;break;}}
      if(has) lastReal=hRow+1+r3;
    }
  }
  var maxR=sh.getMaxRows();
  if(maxR>lastReal){ sh.deleteRows(lastReal+1, maxR-lastReal); log.push('ゴミ尾削除 row'+(lastReal+1)+'〜'+maxR); }

  // 5) 03_売上見込み 確度/着金 配線を現在の列へ修復
  var keyI=idx1('案件・相手'), siI=idx1('報告項目'), svI=idx1('報告値');
  var u=byKw_(ss,'売上見込み');
  if(u && keyI>=0 && siI>=0 && svI>=0){
    var D=colL(keyI), O=colL(siI), P=colL(svI);
    for(var rr=5;rr<=14;rr++){
      var fF="=IFERROR(INDEX(FILTER('02_作業DB'!$"+P+"$2:$"+P+"$2000,('02_作業DB'!$"+D+"$2:$"+D+"$2000=$C"+rr+")*('02_作業DB'!$"+O+"$2:$"+O+"$2000=\"確度\")),COUNTA(FILTER('02_作業DB'!$"+P+"$2:$"+P+"$2000,('02_作業DB'!$"+D+"$2:$"+D+"$2000=$C"+rr+")*('02_作業DB'!$"+O+"$2:$"+O+"$2000=\"確度\")))),R"+rr+")";
      var fG="=IFERROR(INDEX(FILTER('02_作業DB'!$"+P+"$2:$"+P+"$2000,('02_作業DB'!$"+D+"$2:$"+D+"$2000=$C"+rr+")*('02_作業DB'!$"+O+"$2:$"+O+"$2000=\"着金月\")),COUNTA(FILTER('02_作業DB'!$"+P+"$2:$"+P+"$2000,('02_作業DB'!$"+D+"$2:$"+D+"$2000=$C"+rr+")*('02_作業DB'!$"+O+"$2:$"+O+"$2000=\"着金月\")))),S"+rr+")";
      u.getRange("F"+rr).setFormula(fF);
      u.getRange("G"+rr).setFormula(fG);
    }
    log.push('03確度/着金 配線修復 報告項目='+O+'・報告値='+P);
  } else {
    log.push('03配線スキップ（売上見込み or 列未検出）');
  }

  SpreadsheetApp.flush();
  Logger.log('cleanDb02Safe 完了 / backup='+bkName+' / '+log.join(' / '));
  return {ok:true, backup:bkName, log:log};
}

// ── helpers ──────────────────────────────────────────────
function findDb02_(ss){var s=ss.getSheets();for(var i=0;i<s.length;i++){var r=Math.min(6,s[i].getLastRow());if(r<1)continue;var v=s[i].getRange(1,1,r,Math.min(40,s[i].getLastColumn())).getValues();for(var x=0;x<v.length;x++){var j=v[x].join('|');if(j.indexOf('案件・相手')>=0&&j.indexOf('日付')>=0)return s[i];}}return null;}
function findHRow_(sh){var r=Math.min(6,sh.getLastRow());var v=sh.getRange(1,1,r,Math.min(40,sh.getLastColumn())).getValues();for(var x=0;x<v.length;x++)if(v[x].join('|').indexOf('案件・相手')>=0)return x+1;return 1;}
function byKw_(ss,kw){var a=ss.getSheets();for(var i=0;i<a.length;i++)if(a[i].getName().indexOf(kw)>=0)return a[i];return null;}
function columnLetter_(num){var s='';while(num>0){var m=(num-1)%26;s=String.fromCharCode(65+m)+s;num=Math.floor((num-1)/26);}return s;}
