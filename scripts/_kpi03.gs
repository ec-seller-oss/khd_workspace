// 03に「6月KPI実績表」を埋込：02作業DBの動きを週別に自動カウント。E指標/F月内計/G1週/(HIJ非表示)/KLM2-4週。
// ファネル：IMP→CTR→流入→CVR→CV(フォロー)→アポ(面談)→成約 を連動意識できる構成。
function build03KPI(){
  var ss=SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  var u=null,a=ss.getSheets();
  for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf("売上見込み")>=0){ u=a[i]; break; } }
  if(!u) return "03無し";
  var DB="'02_作業DB'!", HEAD="#000000", SUBH="#555555", WHITE="#FFFFFF", GOLD="#C8922A", NAVY="#3C5A78", IN="#FFF2CC";
  var A=u.getRange(1,1,u.getLastRow(),1).getValues();
  function row(kw){ for(var i=0;i<A.length;i++) if((A[i][0]||"").toString().indexOf(kw)>=0) return i+1; return -1; }
  var dh=row("営業ドライバー"); if(dh<0) dh=row("今日の一手"); if(dh<0) return "ドライバー見出し無し";

  // 週範囲(6月・全日)
  var WK=[["6/1-7","2026-06-01","2026-06-07",7],["6/8-14","2026-06-08","2026-06-14",11],["6/15-21","2026-06-15","2026-06-21",12],["6/22-30","2026-06-22","2026-06-30",13]];
  var T='TEXT('+DB+'$A$2:$A$3000,"yyyy-mm-dd")';
  function cnt(kws,ws,we){ var k=kws.map(function(x){return 'ISNUMBER(SEARCH("'+x+'",'+DB+'$H$2:$H$3000))';}).join('+'); return '=SUMPRODUCT(('+T+'>="'+ws+'")*('+T+'<="'+we+'")*(('+k+')>0))'; }
  function tim(ws,we){ return '=SUMPRODUCT(('+T+'>="'+ws+'")*('+T+'<="'+we+'")*('+DB+'$N$2:$N$3000="○")*N('+DB+'$P$2:$P$3000))'; }

  // 既存の右側(E:M)をクリア
  u.getRange(dh,5,9,9).clearContent().setBackground(null);

  // 見出し(ファネル連動を明記)
  u.getRange(dh,5,1,9).merge().setValue("📅 6月KPI 実績カウント（02作業DBから自動）｜流れ：IMP→CTR→流入→CVR→CV(フォロー)→アポ(面談)→成約").setBackground(HEAD).setFontColor(WHITE).setFontWeight("bold").setFontSize(10);
  // ヘッダ行：E指標 F月内計 G1週 (HIJ非表示) K2週 L3週 M4週
  u.getRange(dh+1,5).setValue("指標(略)"); u.getRange(dh+1,6).setValue("月内計");
  u.getRange(dh+1,7).setValue("1週\n6/1-7"); u.getRange(dh+1,11).setValue("2週\n6/8-14"); u.getRange(dh+1,12).setValue("3週\n15-21"); u.getRange(dh+1,13).setValue("4週\n22-30");
  u.getRange(dh+1,5,1,9).setBackground(SUBH).setFontColor(WHITE).setFontWeight("bold").setFontSize(9).setWrap(true);

  // 指標行：CV(フォロー)/アポ(面談)/成約/営業直結時間
  var rows=[
    ["CV(フォロー)",["フォロー","追客"],false],
    ["アポ(面談)",["面談","アポ"],false],
    ["成約(受注)",["成約","受注","決済"],false],
    ["営業直結(分)",null,true]
  ];
  for(var s=0;s<rows.length;s++){ var r=dh+2+s;
    u.getRange(r,5).setValue(rows[s][0]).setFontWeight("bold").setFontSize(9.5);
    WK.forEach(function(w){ u.getRange(r,w[3]).setFormula( rows[s][2]?tim(w[1],w[2]):cnt(rows[s][1],w[1],w[2]) ); });
    u.getRange(r,6).setFormula('=G'+r+'+K'+r+'+L'+r+'+M'+r); // 月内計=各週SUM
  }
  // 将来メディア行(IMP/CTR/流入/CVR)＝手動入力枠(黄)
  var med=[["IMP(表示)"],["CTR(クリック率)"],["流入(セッション)"],["CVR(反応率)"]];
  for(var s=0;s<med.length;s++){ var r=dh+6+s;
    u.getRange(r,5).setValue(med[s][0]).setFontWeight("bold").setFontSize(9.5).setFontColor("#888888");
    u.getRange(r,7,1,1).setBackground(IN); u.getRange(r,11,1,3).setBackground(IN); // 将来手入力(黄)
    u.getRange(r,6).setFormula('=G'+r+'+K'+r+'+L'+r+'+M'+r);
  }
  // 体裁
  u.getRange(dh+1,5,9,1).setBackground("#EFEFEF");
  u.getRange(dh+2,6,4,1).setNumberFormat("#,##0"); u.getRange(dh+2,7,4,1).setNumberFormat("#,##0"); u.getRange(dh+2,11,4,3).setNumberFormat("#,##0");
  u.getRange(dh,5,12,9).setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  u.getRange(dh+10,5,1,9).merge().setValue("※実績は02作業DBの『内容』からフォロー/追客・面談/アポ・成約を週別自動カウント。左の逆算(必要数)と対比してペース管理。IMP〜CVRは媒体開始後に手入力(黄)。").setFontSize(8.5).setFontColor("#777777").setWrap(true);
  u.getRange(dh,5).setNote("6月KPI実績表。02の動きを週別に自動カウント。ファネル全体(IMP→…→成約)を1枚で。HIJ列は非表示(弱気/現実/強気)なので飛ばしてG=1週,K/L/M=2-4週。");
  return {ok:true, driver行:dh};
}
