// 03に一本化：営業ドライバーの右に「予実KPI表」＝逆算(必要)⇔週別実績(02連動)を1表で。重複タブ(ロジックツリー/ダッシュボード)は削除。
function build03Unified(){
  var ss=SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  var u=null,a=ss.getSheets();
  for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf("売上見込み")>=0){ u=a[i]; break; } }
  if(!u) return "03無し";
  var DB="'02_作業DB'!",HEAD="#000000",SUBH="#555555",WHITE="#FFFFFF",GOLD="#C8922A",GREEN="#4E8C5A",BRICK="#AA2E26",IN="#FFF2CC";
  var A=u.getRange(1,1,u.getLastRow(),1).getValues();
  function row(kw){ for(var i=0;i<A.length;i++) if((A[i][0]||"").toString().indexOf(kw)>=0) return i+1; return -1; }
  var dh=row("営業ドライバー"); if(dh<0) dh=row("今日の一手"); if(dh<0) return "ドライバー無し";
  var T='TEXT('+DB+'$A$2:$A$3000,"yyyy-mm-dd")';
  function cnt(kws,ws,we){ var k=kws.map(function(x){return 'ISNUMBER(SEARCH("'+x+'",'+DB+'$H$2:$H$3000))';}).join('+'); return 'SUMPRODUCT(('+T+'>="'+ws+'")*('+T+'<="'+we+'")*(('+k+')>0))'; }
  function tim(ws,we){ return 'SUMPRODUCT(('+T+'>="'+ws+'")*('+T+'<="'+we+'")*('+DB+'$N$2:$N$3000="○")*N('+DB+'$P$2:$P$3000))'; }
  var ix=function(l){ return 'IFERROR(INDEX($B:$B,MATCH("'+l+'",$A:$A,0)),0)'; };
  var WK=[["1週\n6/1-7","2026-06-01","2026-06-07",11],["2週\n8-14","2026-06-08","2026-06-14",12],["3週\n15-21","2026-06-15","2026-06-21",13],["4週\n22-30","2026-06-22","2026-06-30",14]];

  // 既存の右側KPIを一掃(E〜O・ドライバー周辺＋採用額/古田土の右も)
  ["営業ドライバー","勘定科目別 採用額","未来会計図表","採用額 合計"].forEach(function(k){ var r=row(k); if(r>0) u.getRange(r,5,9,11).clearContent().setBackground(null).clearNote(); });

  // 統合 予実KPI表：E指標/F月内必要(逆算)/G月内実績/(HIJ非表示)/K1週/L2週/M3週/N4週/O進捗%
  u.getRange(dh,5,1,11).merge().setValue("📊 予実KPI｜逆算(必要)⇔実績(02から自動)｜流れ：接触(フォロー)→面談(アポ)→成約").setBackground(HEAD).setFontColor(WHITE).setFontWeight("bold").setFontSize(10);
  u.getRange(dh+1,5).setValue("指標"); u.getRange(dh+1,6).setValue("月内\n必要"); u.getRange(dh+1,7).setValue("月内\n実績");
  WK.forEach(function(w){ u.getRange(dh+1,w[3]).setValue(w[0]); });
  u.getRange(dh+1,15).setValue("進捗%");
  u.getRange(dh+1,5,1,11).setBackground(SUBH).setFontColor(WHITE).setFontWeight("bold").setFontSize(9).setWrap(true);

  // 行：フォロー(CV)/面談(アポ)/成約/営業直結(分)
  var rows=[
    ["フォロー(CV)",["フォロー","追客","接触"],"必要な反応数",false],
    ["面談(アポ)",["面談","アポ"],"必要な面談数",false],
    ["成約(受注)",["成約","受注","決済"],"必要 成約数",false],
    ["営業直結(分)",null,"",true]
  ];
  for(var s=0;s<rows.length;s++){ var r=dh+2+s;
    u.getRange(r,5).setValue(rows[s][0]).setFontWeight("bold").setFontSize(9.5);
    if(rows[s][2]) u.getRange(r,6).setFormula('='+ix(rows[s][2])).setBackground("#FFF7E6"); // 月内必要(逆算)
    WK.forEach(function(w){ u.getRange(r,w[3]).setFormula('='+(rows[s][3]?tim(w[1],w[2]):cnt(rows[s][1],w[1],w[2]))); });
    u.getRange(r,7).setFormula('=K'+r+'+L'+r+'+M'+r+'+N'+r); // 月内実績
    if(rows[s][2]) u.getRange(r,15).setFormula('=IFERROR(G'+r+'/F'+r+',"")').setNumberFormat("0%"); // 進捗%
  }
  u.getRange(dh+2,6,4,1).setNumberFormat("#,##0"); u.getRange(dh+2,7,4,1).setNumberFormat("#,##0"); u.getRange(dh+2,11,4,4).setNumberFormat("#,##0");
  u.getRange(dh+1,5,5,11).setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  // 着地見込み行
  var rg=dh+7;
  u.getRange(rg,5).setValue("📈 着地見込み").setFontWeight("bold").setFontColor(BRICK);
  u.getRange(rg,6,1,10).merge().setFormula('="今のペースなら 実成約 "&IFERROR(G'+(dh+4)+',0)&"件。必要 "&IFERROR(F'+(dh+4)+',0)&"件。"&IF(IFERROR(G'+(dh+2)+',0)>=IFERROR(F'+(dh+2)+',0),"✅フォロー量OK→質(率)を磨く","⚠️フォロー量が不足→接触を増やす")').setFontColor(BRICK).setFontWeight("bold").setWrap(true);
  u.getRange(dh,5).setNote("予実KPI(03一本化)。F=逆算の必要数(ドライバー連動)・K〜N=02から週別自動カウント・G=月内実績・O=進捗%。逆算と実績を1表で予実対比。HIJ非表示で1週(K)が実績(G)の隣に見える。営業ロジックツリー/ダッシュボードはここに統合。指標辞書はNotion参照。");

  // 重複タブを削除(03一本化)
  var del=[];
  ["📐営業ロジックツリー","📊ダッシュボード","📈予実トレンド"].forEach(function(n){ var s=ss.getSheetByName(n); if(s){ ss.deleteSheet(s); del.push(n); } });
  return {ok:true, driver行:dh, 削除タブ:del};
}
