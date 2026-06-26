/**
 * 02作業DB に 案件番号(物件)／顧客番号 列を右端追加し、案件・相手(D)から自動抽出して充填
 * 使い方: script.new → 貼付 → addKanriNumbers 実行（冪等＝再実行しても列は増えない）
 * 位置安全: 右端に足すだけ＝D/S/T・O/P/Q/R を一切動かさない（連動・時間式 無傷）
 */
function addKanriNumbers(){
  var ss=SpreadsheetApp.openById('1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc');
  var f=findK_(ss); if(!f) throw '02が見つからない';
  var sh=f.sh, hRow=f.hRow;
  function H(){ return sh.getRange(hRow,1,1,sh.getLastColumn()).getValues()[0]; }
  function col(n){ var h=H(); for(var i=0;i<h.length;i++) if(String(h[i]).trim()===n) return i+1; return -1; }

  // 既存なら使う・無ければ右端に追加（冪等）
  var cK=col('案件番号'), cC=col('顧客番号');
  if(cK<0){ sh.insertColumnAfter(sh.getLastColumn()); cK=sh.getLastColumn(); sh.getRange(hRow,cK).setValue('案件番号'); }
  if(cC<0){ sh.insertColumnAfter(sh.getLastColumn()); cC=sh.getLastColumn(); sh.getRange(hRow,cC).setValue('顧客番号'); }
  sh.getRange(hRow,cK).setBackground('#FFF2CC').setFontWeight('bold').setNote('物件番号(088/277/055等)。Dから自動抽出＋手入力で補完。物件管理マスターDBの物件番号と一致');
  sh.getRange(hRow,cC).setBackground('#FFF2CC').setFontWeight('bold').setNote('顧客番号(H037/H044等)。顧客マスタータブと一致');

  var cD=col('案件・相手'); if(cD<0) throw '案件・相手が無い';
  var last=sh.getLastRow(); if(last<=hRow) { Logger.log('データ無'); return; }
  var n=last-hRow;
  var D=sh.getRange(hRow+1,cD,n,1).getValues();
  var outK=[], outC=[];
  for(var i=0;i<n;i++){
    var s=String(D[i][0]);
    var mh=s.match(/H\d{3}/);                 // 顧客番号 Hxxx
    var mp=s.match(/(?:^|[^0-9A-Za-z])(\d{3})(?:[^0-9]|$)/); // 物件3桁
    outK.push([ mp? mp[1] : '' ]);
    outC.push([ mh? mh[0] : '' ]);
  }
  sh.getRange(hRow+1,cK,n,1).setValues(outK);
  sh.getRange(hRow+1,cC,n,1).setValues(outC);
  SpreadsheetApp.flush();
  Logger.log('案件番号/顧客番号 追加＆充填 完了 行数='+n);
}
function findK_(ss){var s=ss.getSheets();for(var i=0;i<s.length;i++){var r=Math.min(6,s[i].getLastRow());if(r<1)continue;var v=s[i].getRange(1,1,r,Math.min(40,s[i].getLastColumn())).getValues();for(var x=0;x<v.length;x++){var j=v[x].join('|');if(j.indexOf('案件・相手')>=0&&j.indexOf('日付')>=0)return {sh:s[i],hRow:x+1};}}return null;}
