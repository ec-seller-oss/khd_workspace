// 📊経営ダッシュボード：02〜06横断KPI＋温度別パイプライン＋月次逆算＋達成率を1画面＋グラフ
// 全クロスタブrefはIFERROR包み＝#REFを出さない(出れば"確認中"表示→後で直す)
function buildDashboard(){
  var ss=SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  var nm="📊ダッシュボード";
  var sh=ss.getSheetByName(nm); if(sh) ss.deleteSheet(sh);
  sh=ss.insertSheet(nm,0); // 先頭=操縦席フロントページ
  var U3="'03_売上見込み'!", KM="'顧客マスター'!", S01="'01_統合司令塔'!", PL="'04_損益PL'!", DB="'02_作業DB'!", KPI="'時間KPI集計'!";
  function L(r,c,v){ sh.getRange(r,c).setValue(v); }
  function F(r,c,f){ sh.getRange(r,c).setFormula(f); }
  function hd(r,c,t){ sh.getRange(r,c).setValue(t).setBackground("#000000").setFontColor("#FFFFFF").setFontWeight("bold"); }
  function sec(r,c,t,col){ sh.getRange(r,c,1,2).merge().setBackground(col).setFontWeight("bold"); sh.getRange(r,c).setValue(t); }

  // タイトル
  sh.getRange(1,1,1,6).merge().setValue("📊 経営ダッシュボード｜誰を・いつ・どう動かして月次を逆算で埋めるか").setBackground("#000000").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(13);

  // 🎯 月次逆算ゲージ(左)
  sec(3,1,"🎯 月次逆算ゲージ","#FCE5CD");
  L(4,1,"必達ライン(家族)/月"); F(4,2,'=IFERROR('+U3+'B23,"確認中")');
  L(5,1,"確定 経常粗利/月");     F(5,2,'=IFERROR('+U3+'B24,"確認中")');
  L(6,1,"🔴 穴(毎月埋める額)");  F(6,2,'=IFERROR('+U3+'B25,"確認中")');
  L(7,1,"→ 必要 成約数");        F(7,2,'=IFERROR('+U3+'B27,"確認中")');
  L(8,1,"💡 今日の追客本数");    F(8,2,'=IFERROR('+U3+'B36,"確認中")');
  sh.getRange(4,2,5,1).setNumberFormat("#,##0");

  // 🌡️ 温度別パイプライン(右上)
  sec(3,4,"🌡️ 温度別パイプライン(顧客M)","#F4CCCC");
  L(4,4,"HOT");  F(4,5,'=IFERROR(COUNTIF('+KM+'M2:M100,"HOT"),0)');
  L(5,4,"WARM"); F(5,5,'=IFERROR(COUNTIF('+KM+'M2:M100,"WARM"),0)');
  L(6,4,"COLD"); F(6,5,'=IFERROR(COUNTIF('+KM+'M2:M100,"COLD"),0)');
  L(7,4,"顧客総数"); F(7,5,'=IFERROR(COUNTA('+KM+'C2:C100),0)');

  // 💰 財務サマリ(左下)
  sec(10,1,"💰 財務サマリ(01/04/05)","#D9EAD3");
  L(11,1,"純資産");       F(11,2,'=IFERROR('+S01+'B5,"確認中")');
  L(12,1,"現預金");       F(12,2,'=IFERROR('+S01+'B10,"確認中")');
  L(13,1,"ランウェイ");   F(13,2,'=IFERROR('+S01+'B11,"確認中")');
  L(14,1,"通期の谷(最悪月現金)"); F(14,2,'=IFERROR('+S01+'B14,"確認中")');
  L(15,1,"経常(全社/月)"); F(15,2,'=IFERROR('+PL+'B40,"確認中")');

  // 📈 達成率/営業比率(右下)
  sec(10,4,"📈 達成率/営業比率","#CFE2F3");
  L(11,4,"平均 達成率%(02)"); F(11,5,'=IFERROR(AVERAGE('+DB+'R2:R1000),"-")'); sh.getRange(11,5).setNumberFormat("0%");
  L(12,4,"営業直結 時間比率"); F(12,5,'=IFERROR('+KPI+'B2,"確認中")');
  L(13,4,"損益分岐リスク(穴)"); F(13,5,'=IFERROR('+S01+'B15,"確認中")'); sh.getRange(13,5).setNumberFormat("#,##0");

  sh.setColumnWidth(1,180); sh.setColumnWidth(2,140); sh.setColumnWidth(3,30); sh.setColumnWidth(4,170); sh.setColumnWidth(5,120);

  // グラフ① 温度別(列)
  var c1=sh.newChart().asColumnChart().addRange(sh.getRange(4,4,3,2)).setPosition(17,1,0,0)
    .setOption("title","🌡️ 温度別パイプライン").setOption("width",420).setOption("height",240).setOption("legend",{position:"none"}).build();
  sh.insertChart(c1);
  // グラフ② 月次逆算(必達/確定/穴)
  var c2=sh.newChart().asColumnChart().addRange(sh.getRange(4,1,3,2)).setPosition(17,4,0,0)
    .setOption("title","🎯 月次逆算 必達/確定/穴").setOption("width",420).setOption("height",240).setOption("legend",{position:"none"}).build();
  sh.insertChart(c2);

  sh.getRange(1,1).setNote("操縦席フロントページ。各KPIは02-06から自動連動(IFERROR包み)。確認中表示が出たらrefズレ=Claudeに言えば直す。温度別は顧客マスター、逆算は03ドライバー、達成率は02、財務は01/04連動。");
  sh.setFrozenRows(2);
  return {ok:true, tab:nm};
}
