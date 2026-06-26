// 01_統合司令塔：上=視覚グラフ(動かねば!)→中=冷静な数値→下=今日の一手/優先順位(朝ブリーフへ)。会議も上から話せる。
function buildBoard01(){
  var ss=SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  var sh=null,a=ss.getSheets();
  for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf("統合司令塔")>=0||a[i].getName().indexOf("司令")>=0){ sh=a[i]; break; } }
  if(!sh) return "01無し";
  var BS="'06_資産負債BS'!",U3="'03_売上見込み'!",CF="'05_資金繰り'!",PL="'04_損益PL'!",KM="'顧客マスター'!",KPI="'時間KPI集計'!",DB="'02_作業DB'!";
  var CREAM="#F9F6EF",WHITE="#FFFFFF",BRICK="#AA2E26",INK="#2B2B2B",SUB="#7A7268",LINE="#E3DCCD",GREEN="#4E8C5A",GOLD="#C8922A",NAVY="#3C5A78",PUR="#6B3FA0",GRAY="#9AA0A8";
  var Q='"';
  var IM=function(l){return "INDEX("+U3+"B:B,MATCH(\""+l+"\","+U3+"A:A,0))";};
  var FH=IM("*家族が潰れない*"),FK=IM("*確定：経常粗利*"),FA=IM("*経常ギャップ*"),FN=IM("*必要 成約数*"),FT=IM("*今日の追客本数");
  var CASH="("+BS+"BI12+"+BS+"BI13+"+BS+"BI14)", RUN=CASH+"/631342";

  sh.clear(); var chs=sh.getCharts(); for(var i=0;i<chs.length;i++) sh.removeChart(chs[i]);
  sh.getRange(1,1,90,9).setBackground(CREAM).setFontColor(INK).setFontFamily("Arial");
  for(var c=1;c<=8;c++) sh.setColumnWidth(c, c%2===1?150:110); sh.setColumnWidth(9,18);
  sh.setHiddenGridlines(true);
  function tile(r,c,label,formula,accent,fmt){
    sh.getRange(r,c,1,2).merge().setValue(label).setBackground(accent).setFontColor("#FFFFFF").setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    var v=sh.getRange(r+1,c,2,2).merge().setFormula(formula).setBackground(WHITE).setFontColor(INK).setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    v.setBorder(true,true,true,true,false,false,LINE,SpreadsheetApp.BorderStyle.SOLID); if(fmt) v.setNumberFormat(fmt);
  }
  function band(r,t,ac){ sh.getRange(r,1,1,8).merge().setValue(t).setBackground(ac).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(11).setVerticalAlignment("middle"); sh.setRowHeight(r,24); }
  function rowH(r){ sh.setRowHeight(r,15); sh.setRowHeight(r+1,15); sh.setRowHeight(r+2,22); }

  // 隠しグラフデータ(J:L)
  sh.getRange("J2").setValue("優先度 高"); sh.getRange("K2").setFormula('=IFERROR(COUNTIF('+KM+'M2:M100,"HOT"),0)');
  sh.getRange("J3").setValue("優先度 中"); sh.getRange("K3").setFormula('=IFERROR(COUNTIF('+KM+'M2:M100,"WARM"),0)');
  sh.getRange("J4").setValue("優先度 低"); sh.getRange("K4").setFormula('=IFERROR(COUNTIF('+KM+'M2:M100,"COLD"),0)');
  sh.getRange("J7").setValue("必達"); sh.getRange("K7").setFormula('=IFERROR('+FH+',0)');
  sh.getRange("J8").setValue("確定"); sh.getRange("K8").setFormula('=IFERROR('+FK+',0)');
  sh.getRange("J9").setValue("目標差"); sh.getRange("K9").setFormula('=IFERROR('+FA+',0)');
  for(var m=0;m<8;m++){ sh.getRange(12+m,10).setValue((m+1)+"月目"); }
  sh.getRange(12,12).setFormula('=ArrayFormula(TRANSPOSE('+CF+'B35:I35))'); // L12:L19 月末現金
  // 日次予実 helper(N:Q) 直近7日
  sh.getRange("N1").setValue("日"); sh.getRange("O1").setValue("予定"); sh.getRange("P1").setValue("実"); sh.getRange("Q1").setValue("達成率");
  for(var d=0;d<7;d++){ var hr=2+d;
    sh.getRange(hr,14).setFormula('=TODAY()-'+(6-d)).setNumberFormat("mm/dd");
    var dk='TEXT('+DB+'$A$2:$A$3000,"yyyy/mm/dd")=TEXT($N'+hr+',"yyyy/mm/dd")';
    sh.getRange(hr,15).setFormula('=SUMPRODUCT(('+dk+')*N('+DB+'$O$2:$O$3000))');
    sh.getRange(hr,16).setFormula('=SUMPRODUCT(('+dk+')*N('+DB+'$P$2:$P$3000))');
    sh.getRange(hr,17).setFormula('=IFERROR(P'+hr+'/O'+hr+',0)').setNumberFormat("0%");
  }

  // 1 タイトル
  sh.getRange(1,1,1,8).merge().setValue("①  経営司令塔  —  KIKUCHI HOLDINGS").setBackground(BRICK).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(15).setVerticalAlignment("middle"); sh.setRowHeight(1,32);
  // 2 達成ゲージ(動機の見出し・大)
  sh.getRange(2,1,1,8).merge().setFormula('="🔥 今月 "&IFERROR(ROUND('+FK+'/'+FH+'*100,0),0)&"% 達成   "&REPT("●",MIN(10,IFERROR(ROUND('+FK+'/'+FH+'*10,0),0)))&REPT("○",10-MIN(10,IFERROR(ROUND('+FK+'/'+FH+'*10,0),0)))&"   あと "&TEXT(IFERROR(MAX(0,'+FH+'-'+FK+'),0),"#,##0")&" 円・"&IFERROR('+FN+',0)&" 件で必達！"').setBackground("#FBF3E2").setFontColor(BRICK).setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true,true,true,true,false,false,GOLD,SpreadsheetApp.BorderStyle.SOLID); sh.setRowHeight(2,30);

  // ▼ 上=視覚グラフ(まず動かねば!と思う) rows4-23
  band(4,"📊 まず視覚で掴む — 今、動くべきか（穴・伸びしろ・押す相手・未来の谷）",BRICK);
  var c2=sh.newChart().asColumnChart().addRange(sh.getRange("J7:K9")).setPosition(5,1,5,3).setOption("title","🎯 必達 / 確定 / 目標差(=穴)").setOption("width",400).setOption("height",195).setOption("legend",{position:"none"}).setOption("backgroundColor",WHITE).setOption("colors",[BRICK]).setOption("series",{0:{dataLabel:"value"}}).build(); sh.insertChart(c2);
  var c1=sh.newChart().asColumnChart().addRange(sh.getRange("J2:K4")).setPosition(5,5,5,3).setOption("title","🌡️ ご縁の優先度(押す相手)").setOption("width",400).setOption("height",195).setOption("legend",{position:"none"}).setOption("backgroundColor",WHITE).setOption("colors",[PUR]).setOption("series",{0:{dataLabel:"value"}}).build(); sh.insertChart(c1);
  var c3=sh.newChart().asColumnChart().addRange(sh.getRange("N1").offset(0,0,8,3)).setPosition(15,1,5,3).setOption("title","📈 日次 予定vs実(直近7日)").setOption("width",400).setOption("height",195).setOption("colors",[GRAY,GREEN]).setOption("backgroundColor",WHITE).setOption("series",{0:{dataLabel:"value"},1:{dataLabel:"value"}}).build(); sh.insertChart(c3);
  var c4=sh.newChart().asLineChart().addRange(sh.getRange(12,10,8,1)).addRange(sh.getRange(12,12,8,1)).setPosition(15,5,5,3).setOption("title","💰 月末現金の推移(未来の谷)").setOption("width",400).setOption("height",195).setOption("legend",{position:"none"}).setOption("backgroundColor",WHITE).setOption("colors",[NAVY]).setOption("series",{0:{dataLabel:"value"}}).build(); sh.insertChart(c4);

  // ▼ 指針(感情ドライバー) rows24-
  band(24,"💬 今日の指針",GOLD);
  sh.getRange(25,1).setValue("🎚 状態").setFontWeight("bold").setFontColor(BRICK).setVerticalAlignment("middle");
  var modes=["☀️ 通常","🌙 省エネ(疲労時)","🧹 整える(残務)"];
  sh.getRange(25,2,1,2).merge().setValue(modes[0]).setBackground("#FFF3D9").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setBorder(true,true,true,true,false,false,GOLD,SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(25,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(modes,true).build());
  var M='$B$25';
  sh.getRange(25,4,1,5).merge().setFormula('=IF('+M+'='+Q+'🌙 省エネ(疲労時)'+Q+','+Q+'最低ライン：関心の高い先生お一人へ一言で充分。'+Q+',IF('+M+'='+Q+'🧹 整える(残務)'+Q+','+Q+'最低ライン：滞り事項を一つ片付ける。'+Q+',"最低ライン："&IFERROR('+FT+',0)&"件フォロー＋確度UP"))').setBackground(WHITE).setFontColor(INK).setFontWeight("bold").setFontSize(11).setVerticalAlignment("middle").setBorder(true,true,true,true,false,false,LINE,SpreadsheetApp.BorderStyle.SOLID); sh.setRowHeight(25,24);
  sh.getRange(26,1,1,8).merge().setFormula('=IF('+M+'='+Q+'🌙 省エネ(疲労時)'+Q+','+Q+'🍵 秘書より：今日は無理をなさらず。お一人に一言で前進。休息も投資。'+Q+',IF('+M+'='+Q+'🧹 整える(残務)'+Q+','+Q+'🔧 江藤より：滞りを一つ整えましょう。完璧より動かすこと。'+Q+','+Q+'🎯 李牧より：目標差は'+Q+'&TEXT(IFERROR('+FA+',0),'+Q+'#,##0'+Q+')&'+Q+'円。確度の高い先生から順に、数字で確度を上げる。'+Q+'))').setBackground("#FBEEEC").setFontColor(BRICK).setFontSize(12).setFontWeight("bold").setVerticalAlignment("middle").setWrap(true).setBorder(true,true,true,true,false,false,BRICK,SpreadsheetApp.BorderStyle.SOLID); sh.setRowHeight(26,36);

  // ▼ 中=冷静な数値（逆算→達成シミュ→安全余力）
  band(28,"🧮 冷静な数値（逆算）",  "#8C4A2E");
  tile(29,1,"必達 目標/月",'=IFERROR('+FH+',"-")',SUB,"#,##0");
  tile(29,3,"確定 経常粗利",'=IFERROR('+FK+',"-")',GREEN,"#,##0");
  tile(29,5,"目標差(=穴)",'=IFERROR('+FA+',"-")',BRICK,"#,##0");
  tile(29,7,"必要 成約数",'=IFERROR('+FN+',"-")',GOLD,'0"件"'); rowH(29);
  band(32,"🧘 焦らない判断 ＆ 残り時間",GREEN);
  sh.getRange(33,1,1,8).merge().setFormula('=IF(IFERROR('+FK+',0)>=IFERROR('+FH+',0),'+Q+'🟢 必達は達成済。作業を足さず、空いた時間は調査士(10月)・家族へ。'+Q+',IF(IFERROR('+RUN+',0)>=6,'+Q+'🟢 資金に余裕あり。焦らず確度の高い先生に絞り、見えてないリスクの確認を優先。'+Q+','+Q+'🟡 あと '+Q+'&IFERROR('+FN+',0)&'+Q+' 件で必達。優先度の高い先生からお一人ずつ。'+Q+'))').setBackground("#EAF3EC").setFontColor(GREEN).setFontSize(11).setFontWeight("bold").setVerticalAlignment("middle").setWrap(true).setBorder(true,true,true,true,false,false,LINE,SpreadsheetApp.BorderStyle.SOLID); sh.setRowHeight(33,30);
  tile(34,1,"今月の残日数",'=IFERROR(EOMONTH(TODAY(),0)-TODAY(),"-")',NAVY,'0"日"');
  tile(34,3,"調査士試験まで",'=IFERROR(DATE(2026,10,18)-TODAY(),"-")',PUR,'0"日"');
  tile(34,5,"資金持続",'=IFERROR('+RUN+',"-")',GREEN,'0.0"ヶ月"');
  tile(34,7,"通期の谷",'=IFERROR(MIN('+CF+'B35:I35),"-")',GREEN,"#,##0"); rowH(34);
  band(37,"🛟 財務基盤 ＆ 信頼の積み上げ",NAVY);
  tile(38,1,"現預金",'=IFERROR('+CASH+',"-")',NAVY,"#,##0");
  tile(38,3,"純資産",'=IFERROR(SUM('+BS+'BI50:BI144)-SUM('+BS+'BI147:BI167),"-")',NAVY,"#,##0");
  tile(38,5,"ご縁の関係先",'=IFERROR(COUNTA('+KM+'C2:C100),0)',PUR,'0"件"');
  tile(38,7,"本業集中度",'=IFERROR('+KPI+'B2,"-")',GREEN,"0%"); rowH(38);
  sh.getRange(41,1,1,8).merge().setValue("💛 売り込まず、相手目線でお役に立つ。信頼はやがて対価として返る。一人の先生の課題に応えることが半年後の成果に。").setBackground("#F3EEF7").setFontColor(PUR).setFontSize(10).setFontStyle("italic").setVerticalAlignment("middle").setWrap(true); sh.setRowHeight(41,28);

  // ▼ 下=今日の一手 → 優先順位(朝ブリーフのタスク配置へ)
  band(43,"🎯 今日の一手（ここから朝ブリーフでタスク配置）",BRICK);
  sh.getRange(44,1,2,2).merge().setValue("今日のフォロー").setBackground(BRICK).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.getRange(44,3,2,2).merge().setFormula('=IFERROR('+FT+',"-")').setBackground("#FBEEEC").setFontColor(BRICK).setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat('0"件"').setBorder(true,true,true,true,false,false,BRICK,SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(44,5,1,4).merge().setValue("① 締切・お約束 → ② 生命線(栄町・医療) → ③ 関心の高い先生へ確度UP").setBackground(WHITE).setFontColor(INK).setFontSize(10.5).setVerticalAlignment("middle");
  sh.getRange(45,5,1,4).merge().setValue("その後：種まき（発信）／ 内務は最小・AIへ委ねる").setBackground(WHITE).setFontColor(SUB).setFontSize(10).setVerticalAlignment("middle");
  sh.setRowHeight(44,24); sh.setRowHeight(45,22);
  band(47,"📋 優先順位（この順に・考えすぎない）","#000000");
  var pri=[["締切・お約束","先方をお待たせしている事項"],["生命線の案件","栄町の出口・医療の確定案件"],["関心の高い先生","優先度順に確度を上げる"],["種まき","発信・関係づくり"],["内務(最小)","記帳・集計はAIへ委ねる"]];
  for(var i=0;i<pri.length;i++){ var r=48+i;
    sh.getRange(r,1,1,2).merge().setValue(String.fromCharCode(0x2460+i)+" "+pri[i][0]).setBackground(WHITE).setFontColor(INK).setFontWeight("bold").setFontSize(10.5).setVerticalAlignment("middle");
    sh.getRange(r,3,1,6).merge().setValue(pri[i][1]).setBackground(CREAM).setFontColor(SUB).setFontSize(10).setVerticalAlignment("middle"); sh.setRowHeight(r,22);
  }
  sh.getRange(54,1,1,8).merge().setValue("→ この優先順位に沿って、朝ブリーフで今日の時間にタスクを配置する。").setBackground(CREAM).setFontColor(SUB).setFontSize(9.5).setFontStyle("italic");

  sh.hideColumns(10,8); sh.setFrozenRows(3);
  var dash=ss.getSheetByName("📊ダッシュボード"); if(dash) ss.deleteSheet(dash);
  var tr=ss.getSheetByName("📈予実トレンド"); if(tr) ss.deleteSheet(tr);
  sh.getRange(1,1).setNote("司令塔：上=視覚グラフ(動機)→中=冷静な数値(逆算/安全)→下=今日の一手/優先順位(朝ブリーフへ)。03駆動はラベル探索。会議は上からスクロールで話せる。");
  return {ok:true,note:"01再配置=視覚グラフ上→数値中→タスク下(朝ブリーフ動線)"};
}
