// 📐営業ロジックツリー：売上(遅行)を行動量(先行・日次)まで分解。毎日入力→率→着地見込みが自動算出のシミュレーター。KHD厚利少本適用。
function buildLogicTree(){
  var ss=SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  var nm="📐営業ロジックツリー";
  var sh=ss.getSheetByName(nm); if(sh) ss.deleteSheet(sh); sh=ss.insertSheet(nm);
  var DB="'02_作業DB'!", U3="'03_売上見込み'!";
  var CREAM="#F9F6EF",WHITE="#FFFFFF",BRICK="#AA2E26",INK="#2B2B2B",SUB="#7A7268",GREEN="#4E8C5A",GOLD="#C8922A",NAVY="#3C5A78",IN="#FFF2CC",AUTO="#E8F0FE",LINE="#E3DCCD";
  sh.getRange(1,1,60,8).setBackground(CREAM).setFontColor(INK).setFontFamily("Arial");
  sh.setColumnWidth(1,40); sh.setColumnWidth(2,180); sh.setColumnWidth(3,150); sh.setColumnWidth(4,120); sh.setColumnWidth(5,120); sh.setColumnWidth(6,90); sh.setColumnWidth(7,90); sh.setColumnWidth(8,140);
  sh.setHiddenGridlines(true);
  function band(r,t,ac){ sh.getRange(r,1,1,8).merge().setValue(t).setBackground(ac).setFontColor(WHITE).setFontWeight("bold").setFontSize(11).setVerticalAlignment("middle"); sh.setRowHeight(r,24); }
  function inp(r,c,v,fmt){ var x=sh.getRange(r,c).setValue(v).setBackground(IN).setFontWeight("bold").setBorder(true,true,true,true,false,false,GOLD,SpreadsheetApp.BorderStyle.SOLID); if(fmt)x.setNumberFormat(fmt); }
  function calc(r,c,f,fmt){ var x=sh.getRange(r,c).setFormula(f).setBackground(AUTO); if(fmt)x.setNumberFormat(fmt); return x; }

  sh.getRange(1,1,1,8).merge().setValue("📐 営業ロジックツリー  —  売上(遅行) を 行動量(先行・日次) まで分解 ／ 毎日積むのは一番下だけ").setBackground(BRICK).setFontColor(WHITE).setFontWeight("bold").setFontSize(14).setVerticalAlignment("middle"); sh.setRowHeight(1,34);
  sh.getRange(2,1,1,8).merge().setValue("黄=あなたが入力（前提率＆今日の行動）／青=自動計算。下を積むと上が動く。結果(売上)は遅行＝先行(行動)を積めば後から必ず動く。").setBackground(CREAM).setFontColor(SUB).setFontSize(9.5);

  // 🟡 前提(肌感で1回決める)
  band(4,"🟡 前提（肌感で1回決めて固定・実績が出たら微調整）",GOLD);
  sh.getRange(5,2).setValue("月の売上目標(KGI)"); inp(5,4,1710000,"#,##0"); sh.getRange(5,5).setValue("← 03の売上目標と揃える").setFontColor(SUB).setFontSize(9);
  sh.getRange(6,2).setValue("平均受注単価(P)"); inp(6,4,1100000,"#,##0"); sh.getRange(6,5).setValue("医療110万/物件単発の平均").setFontColor(SUB).setFontSize(9);
  sh.getRange(7,2).setValue("成約率(受注÷商談)"); inp(7,4,0.5,"0%"); sh.getRange(7,5).setValue("商談したら何割決まるか").setFontColor(SUB).setFontSize(9);
  sh.getRange(8,2).setValue("商談化率(商談÷アポ)"); inp(8,4,0.6,"0%"); sh.getRange(8,5).setValue("会ったら何割が商談に進むか").setFontColor(SUB).setFontSize(9);
  sh.getRange(9,2).setValue("アポ獲得率(アポ÷接触)"); inp(9,4,0.2,"0%"); sh.getRange(9,5).setValue("接触したら何割が会えるか").setFontColor(SUB).setFontSize(9);
  sh.getRange(10,2).setValue("月の稼働日"); inp(10,4,20,"0"); sh.getRange(10,5).setValue("月に何日動くか").setFontColor(SUB).setFontSize(9);

  // 🌳 ロジックツリー(逆算：目標→今日の行動)
  band(12,"🌳 ロジックツリー｜逆算（目標 → 今日 何件動けばいいか）",NAVY);
  sh.getRange(13,2,1,5).setValues([["階層 / 指標","計算式","必要数(月)","今日あたり","遅/先行"]]).setBackground("#555555").setFontColor(WHITE).setFontWeight("bold").setFontSize(9);
  // L0 売上 = 受注×単価(目標を表示)
  sh.getRange(14,2).setValue("0｜売上高(KGI)").setFontWeight("bold"); sh.getRange(14,3).setValue("=受注×単価"); calc(14,4,"=D5","#,##0"); sh.getRange(14,6).setValue("遅行").setFontColor(BRICK);
  // L1 受注 = 目標÷単価
  sh.getRange(15,2).setValue("1｜必要 受注件数"); sh.getRange(15,3).setValue("=目標÷単価"); calc(15,4,"=ROUNDUP(D5/D6,0)",'0"件"'); calc(15,5,"=ROUNDUP(D15/D10,2)",'0.00"件"'); sh.getRange(15,6).setValue("遅行");
  // L2 商談 = 受注÷成約率
  sh.getRange(16,2).setValue("2｜必要 商談数"); sh.getRange(16,3).setValue("=受注÷成約率"); calc(16,4,"=ROUNDUP(D15/D7,0)",'0"件"'); calc(16,5,"=ROUNDUP(D16/D10,2)",'0.00"件"'); sh.getRange(16,6).setValue("中間");
  // L3 アポ = 商談÷商談化率
  sh.getRange(17,2).setValue("3｜必要 アポ数"); sh.getRange(17,3).setValue("=商談÷商談化率"); calc(17,4,"=ROUNDUP(D16/D8,0)",'0"件"'); calc(17,5,"=ROUNDUP(D17/D10,2)",'0.00"件"'); sh.getRange(17,6).setValue("中間");
  // L4 アプローチ = アポ÷アポ獲得率 ★先行
  sh.getRange(18,2).setValue("4｜必要 アプローチ数 ★").setFontWeight("bold").setFontColor(BRICK); sh.getRange(18,3).setValue("=アポ÷アポ獲得率"); calc(18,4,"=ROUNDUP(D17/D9,0)",'0"件"'); calc(18,5,"=ROUNDUP(D18/D10,1)",'0.0"件"').setFontSize(13).setFontWeight("bold"); sh.getRange(18,6).setValue("先行").setFontColor(GREEN).setFontWeight("bold");
  sh.getRange(19,2,1,7).merge().setFormula('="⬆ ★毎日積むのはココ。今日のアプローチ目標 = "&IFERROR(ROUNDUP(D18/D10,0),0)&" 件（接触＝紹介依頼・業者への物件情報引き出し等）"').setFontColor(BRICK).setFontWeight("bold").setFontSize(11).setBackground("#FBEEEC");
  sh.getRange(14,4,5,2).setBorder(true,true,true,true,true,true,LINE,SpreadsheetApp.BorderStyle.SOLID);

  // ✊ 毎日積む行動量(先行・日次)＝02から自動カウント(当月)
  band(21,"✊ 当月の実績（02作業DBから自動）→ 実際の率 → 着地見込み",GREEN);
  var mS='"'+ "2026-06-01" +'"', mE='"'+ "2026-06-30" +'"';
  var T='TEXT('+DB+'$A$2:$A$3000,"yyyy-mm-dd")';
  function cnt(kws){ var k=kws.map(function(x){return 'ISNUMBER(SEARCH("'+x+'",'+DB+'$H$2:$H$3000))';}).join('+'); return '=SUMPRODUCT(('+T+'>='+mS+')*('+T+'<='+mE+')*(('+k+')>0))'; }
  sh.getRange(22,2,1,4).setValues([["指標","当月 実績","実際の率","判定"]]).setBackground("#555555").setFontColor(WHITE).setFontWeight("bold").setFontSize(9);
  sh.getRange(23,2).setValue("実 アプローチ(接触)"); calc(23,3,cnt(["架電","メール","訪問","紹介","フォロー","追客","接触"]),'0"件"');
  sh.getRange(24,2).setValue("実 アポ(面談)"); calc(24,3,cnt(["面談","アポ"]),'0"件"'); calc(24,4,"=IFERROR(D24/D23,\"-\")","0%"); sh.getRange(24,5).setValue("実アポ獲得率").setFontColor(SUB).setFontSize(9);
  sh.getRange(25,2).setValue("実 商談"); calc(25,3,cnt(["商談","提案","見積"]),'0"件"'); calc(25,4,"=IFERROR(D25/D24,\"-\")","0%"); sh.getRange(25,5).setValue("実商談化率").setFontColor(SUB).setFontSize(9);
  sh.getRange(26,2).setValue("実 受注(成約)"); calc(26,3,cnt(["成約","受注","決済"]),'0"件"'); calc(26,4,"=IFERROR(D26/D25,\"-\")","0%"); sh.getRange(26,5).setValue("実成約率").setFontColor(SUB).setFontSize(9);
  // 着地見込み
  sh.getRange(28,2).setValue("📊 着地売上見込み(当月)").setFontWeight("bold"); calc(28,3,"=D26*D6","#,##0").setFontSize(13).setFontWeight("bold"); sh.getRange(28,4).setValue("=実受注×単価").setFontColor(SUB).setFontSize(9);
  sh.getRange(29,2).setValue("達成率(着地÷目標)").setFontWeight("bold"); calc(29,3,"=IFERROR(D28/D5,0)","0%").setFontWeight("bold");
  sh.getRange(30,2,1,5).merge().setFormula('="判定：実アプローチ "&IFERROR(D23,0)&"件 / 必要 "&IFERROR(D18,0)&"件。"&IF(IFERROR(D23,0)>=IFERROR(D18,0),"✅ ペースOK・量は足りてる→率(質)を磨く","⚠️ 量が不足→アプローチを増やす")').setFontColor(NAVY).setFontWeight("bold").setBackground("#EAF3EC").setWrap(true); sh.setRowHeight(30,30);

  // 🏠 KHD適用
  band(32,"🏠 KHD適用（厚利少本・紹介で年1.2件）",BRICK);
  sh.getRange(33,1,4,8).merge().setValue(
    "あなたの本業は『件数少・単価大』。一般的なテレアポ件数でなく、日次の核は ①紹介依頼数 ②業者(バイセル/サンエー等)への「物件情報の引き出し」接触数 ③既存の先生・関係先への追客接触。\n"+
    "・毎日積むのは『接触の質×回数』＝上の L4 アプローチ。\n"+
    "・率(アポ獲得/成約)は週次で振り返り、低ければトーク・提案の中身を改善（量はあるのに率が低い＝中身の問題）。\n"+
    "・売上(結果)は月次。先行(接触)を積めば後から動く＝『今日サボると1ヶ月後に響く』が数字で見える。"
  ).setBackground("#FBEEEC").setFontColor(INK).setFontSize(10).setVerticalAlignment("top").setWrap(true);
  sh.setFrozenRows(2);
  sh.getRange(1,1).setNote("営業ロジックツリー(KHD適用)。前提(黄)を決める→逆算で今日のアプローチ数→毎日02に接触を記録→実績/率/着地見込みが自動。指標辞書の全体版はNotion参照。");
  return {ok:true,tab:nm};
}
