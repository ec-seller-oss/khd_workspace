// 📈予実トレンド：ひと目KPI＋セル内ミニグラフ(SPARKLINE)を散りばめ＋数値入り埋め込みグラフ。ぱっと見で分かる版。
function buildTrend(){
  var ss=SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  var nm="📈予実トレンド";
  var sh=ss.getSheetByName(nm); if(sh) ss.deleteSheet(sh); sh=ss.insertSheet(nm);
  var chs=sh.getCharts(); for(var i=0;i<chs.length;i++) sh.removeChart(chs[i]);
  var DB="'02_作業DB'!", CF="'05_資金繰り'!", BS="'06_資産負債BS'!", U3="'03_売上見込み'!";
  var CREAM="#F9F6EF",BRICK="#AA2E26",INK="#2B2B2B",GREEN="#4E8C5A",GOLD="#C8922A",NAVY="#3C5A78",WHITE="#FFFFFF",HEAD="#000000",GRAY="#9AA0A8";
  sh.getRange(1,1,70,12).setBackground(CREAM).setFontColor(INK).setFontFamily("Arial");
  sh.setHiddenGridlines(true);
  function band(r,n,t,ac){ sh.getRange(r,1,1,n).merge().setValue(t).setBackground(ac).setFontColor(WHITE).setFontWeight("bold").setFontSize(11).setVerticalAlignment("middle"); sh.setRowHeight(r,24); }
  function head(r,c,arr){ sh.getRange(r,c,1,arr.length).setValues([arr]).setBackground(HEAD).setFontColor(WHITE).setFontWeight("bold").setFontSize(9.5); }
  function spark(r,c,f){ sh.getRange(r,c).setFormula(f); }

  sh.getRange(1,1,1,12).merge().setValue("📈 予実トレンド ｜ 過去 → 今日 → 未来（数値＋ミニグラフで一目）").setBackground(BRICK).setFontColor(WHITE).setFontWeight("bold").setFontSize(14).setVerticalAlignment("middle"); sh.setRowHeight(1,32);

  // 日次データ(直近14日)を先に作る(行5-18)※後段KPIが参照
  band(4,12,"🗓 日次 予実（直近14日）｜予定 vs 実績・達成率・ミニグラフ",NAVY);
  head(5,1,["日付","予定(分)","実(分)","差分","達成率","予定vs実 ▮","達成率推移","",""]);
  for(var d=0;d<14;d++){ var r=6+d;
    sh.getRange(r,1).setFormula('=TODAY()-'+(13-d)).setNumberFormat("mm/dd(ddd)");
    var dk='TEXT('+DB+'$A$2:$A$3000,"yyyy/mm/dd")=TEXT($A'+r+',"yyyy/mm/dd")';
    sh.getRange(r,2).setFormula('=SUMPRODUCT(('+dk+')*N('+DB+'$O$2:$O$3000))');
    sh.getRange(r,3).setFormula('=SUMPRODUCT(('+dk+')*N('+DB+'$P$2:$P$3000))');
    sh.getRange(r,4).setFormula('=C'+r+'-B'+r);
    sh.getRange(r,5).setFormula('=IFERROR(C'+r+'/B'+r+',"")').setNumberFormat("0%");
    spark(r,6,'=SPARKLINE({B'+r+',C'+r+'},{"charttype","bar";"color1","#9AA0A8";"color2","#4E8C5A";"max",MAX(B'+r+',C'+r+',1)})');
  }
  sh.getRange(6,2,14,3).setNumberFormat("#,##0");
  // 達成率14日 折れ線(セル内・1セルで全期間)
  spark(6,7,'=SPARKLINE(E6:E19,{"charttype","line";"color1","#C8922A";"linewidth",2})'); sh.getRange(6,7,14,1).merge();

  // ── ひと目KPI(SPARKLINE散りばめ) 行20〜 ──
  band(21,12,"📌 ひと目KPI（数値＋ミニグラフ）",GREEN);
  function kpi(r,c,label,val,fmt,sparkF,col){
    sh.getRange(r,c).setValue(label).setFontSize(9).setFontColor("#7A7268").setFontWeight("bold");
    sh.getRange(r+1,c).setFormula(val).setFontSize(20).setFontWeight("bold").setFontColor(col||INK); if(fmt) sh.getRange(r+1,c).setNumberFormat(fmt);
    if(sparkF) spark(r+2,c,sparkF);
  }
  kpi(22,1,"今週 達成率",'=IFERROR(SUM(C13:C19)/SUM(B13:B19),"-")',"0%",'=SPARKLINE(E13:E19,{"charttype","line";"color1","#4E8C5A"})',GREEN);
  kpi(22,3,"14日 達成率",'=IFERROR(SUM(C6:C19)/SUM(B6:B19),"-")',"0%",'=SPARKLINE(E6:E19,{"charttype","column";"color1","#C8922A"})',GOLD);
  kpi(22,5,"通期の谷",'=IFERROR(MIN('+CF+'B35:I35),"-")',"#,##0",'=SPARKLINE('+CF+'B35:I35,{"charttype","line";"color1","#3C5A78"})',NAVY);
  kpi(22,7,"純資産",'=IFERROR(SUM('+BS+'BI50:BI144)-SUM('+BS+'BI147:BI167),"-")',"#,##0",'=SPARKLINE('+CF+'B35:I35,{"charttype","line";"color1","#4E8C5A"})',NAVY);
  kpi(22,9,"今月の穴",'=IFERROR(INDEX('+U3+'B:B,MATCH("*経常ギャップ*",'+U3+'A:A,0)),"-")',"#,##0",'',BRICK);
  kpi(22,11,"必要 成約",'=IFERROR(INDEX('+U3+'B:B,MATCH("*必要 成約数*",'+U3+'A:A,0)),"-")','0"件"','',GOLD);

  // ── 月次 資金トレンド(縦持ち)＋数値入りグラフ ──
  band(27,12,"💰 月次 資金トレンド（月末現金＝未来の谷が見える）",NAVY);
  head(28,1,["月(当期)","月末現金"]);
  for(var m=0;m<8;m++) sh.getRange(29+m,1).setValue((m+1)+"月目");
  sh.getRange(29,2).setFormula('=ArrayFormula(TRANSPOSE('+CF+'B35:I35))'); sh.getRange(29,2,8,1).setNumberFormat("#,##0");

  // ── 数値入り 埋め込みグラフ(データラベルON)・下部にまとめて重なり防止 ──
  band(38,12,"📊 グラフ（数値入り）",GOLD);
  var dc=sh.newChart().asColumnChart().addRange(sh.getRange(5,1,15,3)).setPosition(39,1,5,2)
    .setOption("title","日次 予定 vs 実績（分）").setOption("width",460).setOption("height",260)
    .setOption("colors",[GRAY,GREEN]).setOption("backgroundColor",WHITE)
    .setOption("series",{0:{dataLabel:"value"},1:{dataLabel:"value"}}).build();
  sh.insertChart(dc);
  var ac=sh.newChart().asLineChart().addRange(sh.getRange(5,1,15,1)).addRange(sh.getRange(5,5,15,1)).setPosition(39,7,5,2)
    .setOption("title","達成率トレンド(14日)").setOption("width",460).setOption("height",260)
    .setOption("colors",[GOLD]).setOption("backgroundColor",WHITE).setOption("series",{0:{dataLabel:"value"}}).setOption("legend",{position:"none"}).build();
  sh.insertChart(ac);
  var mc=sh.newChart().asLineChart().addRange(sh.getRange(28,1,9,2)).setPosition(54,1,5,2)
    .setOption("title","月末現金の推移（谷＝最薄月）").setOption("width",560).setOption("height",260)
    .setOption("colors",[NAVY]).setOption("backgroundColor",WHITE).setOption("series",{0:{dataLabel:"value"}}).setOption("legend",{position:"none"}).build();
  sh.insertChart(mc);

  sh.setColumnWidth(1,110); for(var c=2;c<=12;c++) sh.setColumnWidth(c,86);
  sh.setRowHeight(6,40); // 達成率推移sparkセル高
  sh.setFrozenRows(1);
  sh.getRange(1,1).setNote("予実トレンド。数値＋ミニグラフ(SPARKLINE)で一目。埋め込みグラフはデータラベルON。日次=02予定O/実P(TEXT正規化で日付型/文字列両対応)。月次=05月末現金。");
  return {ok:true,tab:nm};
}
