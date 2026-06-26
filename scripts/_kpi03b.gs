// 採用額(金額)と未来会計図表(売上/粗利)にも週別トラッキング表(E指標/F月内計/G1週/HIJ非表示/KLM2-4週)を付ける。
// 実績=02の成約カウント×平均単価で概算連動。古田土は売上(PQ)→粗利(MQ)。
function build03KPI2(){
  var ss=SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  var u=null,a=ss.getSheets();
  for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf("売上見込み")>=0){ u=a[i]; break; } }
  if(!u) return "03無し";
  var DB="'02_作業DB'!",HEAD="#000000",SUBH="#555555",WHITE="#FFFFFF",GREEN="#4E8C5A",GOLD="#C8922A";
  var A=u.getRange(1,1,u.getLastRow(),1).getValues();
  function row(kw){ for(var i=0;i<A.length;i++) if((A[i][0]||"").toString().indexOf(kw)>=0) return i+1; return -1; }
  var T='TEXT('+DB+'$A$2:$A$3000,"yyyy-mm-dd")';
  function cnt(kws,ws,we){ var k=kws.map(function(x){return 'ISNUMBER(SEARCH("'+x+'",'+DB+'$H$2:$H$3000))';}).join('+'); return '(SUMPRODUCT(('+T+'>="'+ws+'")*('+T+'<="'+we+'")*(('+k+')>0)))'; }
  var WK=[["1週\n6/1-7","2026-06-01","2026-06-07",7],["2週\n8-14","2026-06-08","2026-06-14",11],["3週\n15-21","2026-06-15","2026-06-21",12],["4週\n22-30","2026-06-22","2026-06-30",13]];
  var TANKA='IFERROR(INDEX($B:$B,MATCH("平均成約単価",$A:$A,0)),1100000)';
  var RITSU='IFERROR(INDEX($D:$D,MATCH("粗利益額（MQ",$A:$A,0)),0.65)';

  function gridHeader(r,title){
    u.getRange(r,5,1,9).merge().setValue(title).setBackground(HEAD).setFontColor(WHITE).setFontWeight("bold").setFontSize(10);
    u.getRange(r+1,5).setValue("指標"); u.getRange(r+1,6).setValue("月内計");
    u.getRange(r+1,7).setValue(WK[0][0]); u.getRange(r+1,11).setValue(WK[1][0]); u.getRange(r+1,12).setValue(WK[2][0]); u.getRange(r+1,13).setValue(WK[3][0]);
    u.getRange(r+1,5,1,9).setBackground(SUBH).setFontColor(WHITE).setFontWeight("bold").setFontSize(9).setWrap(true);
  }

  // ① 勘定科目別 採用額 の右：確定売上トラッキング(週別・概算)
  var rk=row("勘定科目別 採用額"); if(rk<0) rk=row("採用額 合計");
  if(rk>0){ u.getRange(rk,5,8,9).clearContent().setBackground(null);
    gridHeader(rk,"💴 6月 確定売上トラッキング（週別・概算＝成約件×平均単価）");
    // 成約(件)
    var r2=rk+2; u.getRange(r2,5).setValue("成約(件)").setFontWeight("bold").setFontSize(9.5);
    WK.forEach(function(w){ u.getRange(r2,w[3]).setFormula('='+cnt(["成約","受注","決済"],w[1],w[2])); });
    u.getRange(r2,6).setFormula('=G'+r2+'+K'+r2+'+L'+r2+'+M'+r2);
    // 確定売上(概算)=成約件×単価
    var r3=rk+3; u.getRange(r3,5).setValue("確定売上(概算)").setFontWeight("bold").setFontSize(9.5);
    WK.forEach(function(w){ u.getRange(r3,w[3]).setFormula('='+cnt(["成約","受注","決済"],w[1],w[2])+'*'+TANKA); });
    u.getRange(r3,6).setFormula('=G'+r3+'+K'+r3+'+L'+r3+'+M'+r3);
    u.getRange(r3,6,1,1).setNumberFormat("#,##0"); u.getRange(r3,7).setNumberFormat("#,##0"); u.getRange(r3,11,1,3).setNumberFormat("#,##0");
    u.getRange(rk,5,5,9).setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    u.getRange(rk+4,5,1,9).merge().setValue("※確定見込(確度加重)は左のB20。ここは実際に成約した件×単価の概算。着金は④資金繰りへ。").setFontSize(8.5).setFontColor("#777777").setWrap(true);
  }

  // ② 未来会計図表 の右：売上(PQ)→粗利(MQ) トラッキング(週別)
  var rg=row("未来会計図表");
  if(rg>0){ u.getRange(rg,5,8,9).clearContent().setBackground(null);
    gridHeader(rg,"📈 6月 売上・粗利トラッキング（週別・実績見合い）");
    var r2=rg+2; u.getRange(r2,5).setValue("売上PQ(実)").setFontWeight("bold").setFontSize(9.5);
    WK.forEach(function(w){ u.getRange(r2,w[3]).setFormula('='+cnt(["成約","受注","決済"],w[1],w[2])+'*'+TANKA); });
    u.getRange(r2,6).setFormula('=G'+r2+'+K'+r2+'+L'+r2+'+M'+r2);
    var r3=rg+3; u.getRange(r3,5).setValue("粗利MQ(実)").setFontWeight("bold").setFontSize(9.5);
    WK.forEach(function(w){ u.getRange(r3,w[3]).setFormula('=G'+r2+'*0+'+cnt(["成約","受注","決済"],w[1],w[2])+'*'+TANKA+'*'+RITSU); });
    u.getRange(r3,6).setFormula('=G'+r3+'+K'+r3+'+L'+r3+'+M'+r3);
    u.getRange(r2,6,2,1).setNumberFormat("#,##0"); u.getRange(r2,7,2,1).setNumberFormat("#,##0"); u.getRange(r2,11,2,3).setNumberFormat("#,##0");
    u.getRange(rg,5,5,9).setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    u.getRange(rg+4,5,1,9).merge().setValue("※売上PQ=成約件×単価の実績。粗利MQ=売上×粗利率(古田土D列)。月次の確定構造は左の図、ここは週別の積み上げ。").setFontSize(8.5).setFontColor("#777777").setWrap(true);
  }
  return {ok:true, 採用額行:rk, 古田土行:rg};
}
