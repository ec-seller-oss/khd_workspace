// 01_統合司令塔をDAZN風(ダーク・タイル)に再デザイン＋ダッシュボード全部入り(月次逆算/温度別/財務/達成率/グラフ2枚)。1枚で完結。
function buildDazn01(){
  var ss=SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  var sh=null,a=ss.getSheets();
  for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf("統合司令塔")>=0||a[i].getName().indexOf("司令")>=0){ sh=a[i]; break; } }
  if(!sh) return "01無し";
  var BS="'06_資産負債BS'!", U3="'03_売上見込み'!", CF="'05_資金繰り'!", PL="'04_損益PL'!", KM="'顧客マスター'!", KPI="'時間KPI集計'!", DB="'02_作業DB'!";
  var DARK="#16181D", CARD="#23262E", INK="#E8EAED", SUB="#9AA0A8", BLUE="#3B82F6", GREEN="#22C55E", RED="#EF4444", ORANGE="#F59E0B", PUR="#A855F7";

  sh.clear(); var chs=sh.getCharts(); for(var i=0;i<chs.length;i++) sh.removeChart(chs[i]);
  sh.getRange(1,1,46,9).setBackground(DARK).setFontColor(INK).setFontFamily("Arial");
  for(var c=1;c<=8;c++) sh.setColumnWidth(c, c%2===1?150:110);
  sh.setColumnWidth(9,20);
  sh.setHiddenGridlines(true);

  function tile(r,c,label,formula,accent,fmt){
    sh.getRange(r,c,1,2).merge().setValue(label).setBackground(accent).setFontColor("#FFFFFF").setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    var v=sh.getRange(r+1,c,2,2).merge().setFormula(formula).setBackground(CARD).setFontColor(INK).setFontSize(18).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    if(fmt) v.setNumberFormat(fmt);
  }
  function band(r,t,ac){ sh.getRange(r,1,1,8).merge().setValue(t).setBackground(ac||"#000000").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(11).setVerticalAlignment("middle"); sh.setRowHeight(r,24); }
  function rowH(r){ sh.setRowHeight(r,15); sh.setRowHeight(r+1,16); sh.setRowHeight(r+2,22); }

  // タイトル
  sh.getRange(1,1,1,8).merge().setValue("①  統合司令塔  —  経営パイロット").setBackground("#000000").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(15).setVerticalAlignment("middle"); sh.setRowHeight(1,38);
  sh.getRange(2,1,1,8).merge().setFormula('="基準日 "&TEXT(TODAY(),"yyyy/mm/dd")&"   ｜   朝=01を見る / 日中=動いたら一言 / 夜=質問に答える"').setBackground(DARK).setFontColor(SUB).setFontSize(9.5).setVerticalAlignment("middle");

  // 💰 現在の体力
  band(4,"💰 現在の体力","#1E3A5F");
  tile(5,1,"純資産",'=IFERROR(SUM('+BS+'BI50:BI144)-SUM('+BS+'BI147:BI167),"-")',BLUE,"#,##0");
  tile(5,3,"現預金",'=IFERROR('+BS+'BI12+'+BS+'BI13+'+BS+'BI14,"-")',BLUE,"#,##0");
  tile(5,5,"ランウェイ",'=IFERROR(('+BS+'BI12+'+BS+'BI13+'+BS+'BI14)/631342,"-")',GREEN,'0.0"ヶ月"');
  tile(5,7,"通期の谷",'=IFERROR(MIN('+CF+'B35:I35),"-")',GREEN,"#,##0"); rowH(5);

  // 🎯 月次逆算
  band(8,"🎯 月次逆算ゲージ（必達−確定＝穴→必要成約）","#5F1E1E");
  tile(9,1,"必達ライン/月",'=IFERROR('+U3+'B23,"-")',SUB,"#,##0");
  tile(9,3,"確定 経常粗利",'=IFERROR('+U3+'B24,"-")',GREEN,"#,##0");
  tile(9,5,"🔴 毎月の穴",'=IFERROR('+U3+'B25,"-")',RED,"#,##0");
  tile(9,7,"必要 成約数",'=IFERROR('+U3+'B27,"-")',ORANGE,'0"件"'); rowH(9);

  // 🌡️ 温度別パイプライン
  band(12,"🌡️ 温度別パイプライン（誰を確度UPするか）","#3A1E5F");
  tile(13,1,"HOT 🔥",'=IFERROR(COUNTIF('+KM+'M2:M100,"HOT"),0)',RED,'0"名"');
  tile(13,3,"WARM",'=IFERROR(COUNTIF('+KM+'M2:M100,"WARM"),0)',ORANGE,'0"名"');
  tile(13,5,"COLD",'=IFERROR(COUNTIF('+KM+'M2:M100,"COLD"),0)',BLUE,'0"名"');
  tile(13,7,"顧客 総数",'=IFERROR(COUNTA('+KM+'C2:C100),0)',PUR,'0"名"'); rowH(13);

  // 📈 達成率/営業
  band(16,"📈 達成率 / 営業比率 / 今日の一手","#1E5F3A");
  tile(17,1,"今日の追客本数",'=IFERROR('+U3+'B36,"-")',ORANGE,'0"件"');
  tile(17,3,"平均 達成率%",'=IFERROR(AVERAGE('+DB+'R2:R1000),"-")',GREEN,"0%");
  tile(17,5,"営業直結 比率",'=IFERROR('+KPI+'B2,"-")',ORANGE,"0%");
  tile(17,7,"経常(全社/月)",'=IFERROR('+PL+'B40,"-")',GREEN,"#,##0"); rowH(17);

  // グラフ用 隠しデータ(J:K)
  sh.getRange("J1").setValue("温度"); sh.getRange("K1").setValue("名");
  sh.getRange("J2").setValue("HOT"); sh.getRange("K2").setFormula('=IFERROR(COUNTIF('+KM+'M2:M100,"HOT"),0)');
  sh.getRange("J3").setValue("WARM"); sh.getRange("K3").setFormula('=IFERROR(COUNTIF('+KM+'M2:M100,"WARM"),0)');
  sh.getRange("J4").setValue("COLD"); sh.getRange("K4").setFormula('=IFERROR(COUNTIF('+KM+'M2:M100,"COLD"),0)');
  sh.getRange("J6").setValue("逆算"); sh.getRange("K6").setValue("円");
  sh.getRange("J7").setValue("必達"); sh.getRange("K7").setFormula('=IFERROR('+U3+'B23,0)');
  sh.getRange("J8").setValue("確定"); sh.getRange("K8").setFormula('=IFERROR('+U3+'B24,0)');
  sh.getRange("J9").setValue("穴"); sh.getRange("K9").setFormula('=IFERROR('+U3+'B25,0)');

  band(20,"📊 グラフ（温度別パイプライン ／ 月次逆算）","#000000");
  var c1=sh.newChart().asColumnChart().addRange(sh.getRange("J2:K4")).setPosition(21,1,5,5)
    .setOption("title","🌡️ 温度別パイプライン").setOption("width",420).setOption("height",230)
    .setOption("backgroundColor",CARD).setOption("legend",{position:"none"})
    .setOption("titleTextStyle",{color:"#FFFFFF"}).setOption("colors",["#A855F7"])
    .setOption("hAxis",{textStyle:{color:"#CCCCCC"}}).setOption("vAxis",{textStyle:{color:"#CCCCCC"}}).build();
  sh.insertChart(c1);
  var c2=sh.newChart().asColumnChart().addRange(sh.getRange("J7:K9")).setPosition(21,5,5,5)
    .setOption("title","🎯 月次逆算 必達/確定/穴").setOption("width",420).setOption("height",230)
    .setOption("backgroundColor",CARD).setOption("legend",{position:"none"})
    .setOption("titleTextStyle",{color:"#FFFFFF"}).setOption("colors",["#EF4444"])
    .setOption("hAxis",{textStyle:{color:"#CCCCCC"}}).setOption("vAxis",{textStyle:{color:"#CCCCCC"}}).build();
  sh.insertChart(c2);

  // 📋 今日の優先順位
  band(34,"📋 今日の優先順位（この順に処理・考えない）","#000000");
  var pri=[["① 🔴 緊急","締切/決済/先方を待たせてる"],["② 💰 生命線案件の前進","栄町出口・医療確定案件"],["③ 🎯 ホット客 追客","HOT客から確度UP(温度順)"],["④ 🌱 仕込み","YouTube/HP/X・関係構築"],["⑤ 📋 内務(最小)","記帳/台帳/集計はAIへ"]];
  for(var i=0;i<pri.length;i++){ var r=35+i;
    sh.getRange(r,1,1,2).merge().setValue(pri[i][0]).setBackground(CARD).setFontColor(INK).setFontWeight("bold").setFontSize(10.5).setVerticalAlignment("middle");
    sh.getRange(r,3,1,6).merge().setValue(pri[i][1]).setBackground(DARK).setFontColor(SUB).setFontSize(10).setVerticalAlignment("middle");
    sh.setRowHeight(r,22);
  }
  sh.hideColumns(10,2); // J:K隠す
  sh.setFrozenRows(2);
  // 重複する📊ダッシュボードタブを削除(01に統合)
  var dash=ss.getSheetByName("📊ダッシュボード"); if(dash) ss.deleteSheet(dash);
  sh.getRange(1,1).setNote("DAZN風・1枚完結司令塔。財務/逆算/温度別/達成率/グラフを06BS・03・04・05・顧客M・時間KPIから自動連動(IFERROR包み)。-が出たらrefズレ=Claudeへ。");
  return {ok:true, note:"01に全部入り(DAZN+ダッシュボード+グラフ2枚)・📊タブ統合削除"};
}
