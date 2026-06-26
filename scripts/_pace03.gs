// 03に「逆算ペース表(目標まで今日/今週/月内で何件)」＋金額ペース＋入力黄/自動青/将来メディア分離/注記 を実装。
function build03Pace(){
  var ss=SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  var u=null,a=ss.getSheets();
  for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf("売上見込み")>=0){ u=a[i]; break; } }
  if(!u) return "03無し";
  var IN="#FFF2CC",AUTO="#CFE2F3",MEDIA="#ECECEC",HEAD="#000000",WHITE="#FFFFFF",GOLD="#C8922A",NAVY="#3C5A78",INK="#2B2B2B";
  var A=u.getRange(1,1,u.getLastRow(),1).getValues();
  function row(kw){ for(var i=0;i<A.length;i++) if((A[i][0]||"").toString().indexOf(kw)>=0) return i+1; return -1; }
  function paint(kw,bg){ var r=row(kw); if(r>0) u.getRange(r,2).setBackground(bg); return r; }
  function note(kw,c,txt){ var r=row(kw); if(r>0) u.getRange(r,c).setNote(txt); }

  // ── A/B：入力=黄 / 自動=青 ──
  ["平均成約単価","面談→成約 転換率","フォロー→面談 転換率","反応率（CVR","クリック率（CTR","月の稼働日"].forEach(function(k){ paint(k,IN); });
  ["家族が潰れない","確定：経常粗利","経常ギャップ","必要 成約数","必要な面談数","必要な反応数","必要な認知数","今日の追客本数","本業集中度"].forEach(function(k){ paint(k,AUTO); });
  // 入力セルに注記
  note("平均成約単価",1,"【入力・黄】1件あたり平均いくらか(医療110万等)。肌感で1回決めて固定。実績出たら微調整。");
  note("面談→成約 転換率",1,"【入力・黄】面談したら何割成約するか(肌感・例20%)。");
  note("フォロー→面談 転換率",1,"【入力・黄】連絡したら何割会えるか(肌感・例50%)。");
  note("月の稼働日",1,"【入力・黄】月に何日動くか(例20)。");

  // ── 今/将来 分離：IMP/CTR/CVRは将来メディア用＝グレー＋注記 ──
  ["反応率（CVR","クリック率（CTR","必要な認知数"].forEach(function(k){ var r=row(k); if(r>0) u.getRange(r,1,1,4).setBackground(MEDIA); });
  note("必要な認知数",1,"▼ここから上(IMP/CTR/CVR)は【将来メディア用】YouTube/HP/Xで「新しいご縁」を作る時の指標。今は使わなくてよい。");
  note("今日の追客本数",3,"◀ 今はココを見て動く。既存のご縁(顧客マスター)へ今日何件フォローするか。");
  // ドライバー見出しに使い方注記
  var dh=row("営業ドライバー"); if(dh<0) dh=row("今日の一手");
  if(dh>0) u.getRange(dh,1).setNote("【使い方】黄=入力(肌感の前提・固定)/青=自動。必達−確定=穴→単価で割り→成約→面談→フォロー→『今日のフォロー本数』。IMP/CTR段(グレー)は将来メディア用。右の『逆算ペース表』で今日/今週/月内の件数が分かる。");

  // ── 逆算ペース表(件数)：ドライバーの右 E〜H ──
  var remD="(EOMONTH(TODAY(),0)-TODAY()+1)"; var remW="ROUNDUP("+remD+"/7,0)";
  var ix=function(l){ return 'IFERROR(INDEX($B:$B,MATCH("'+l+'",$A:$A,0)),0)'; };
  if(dh>0){ var R=dh;
    u.getRange(R,5,1,4).merge().setValue("📅 逆算ペース｜目標まで 今日/今週/月内 で何件").setBackground(HEAD).setFontColor(WHITE).setFontWeight("bold").setFontSize(10);
    u.getRange(R+1,5,1,4).setValues([["段階","月内 必要","今週","今日"]]).setBackground("#555555").setFontColor(WHITE).setFontWeight("bold").setFontSize(9);
    var stg=[["成約(受注)","必要 成約数"],["面談(アポ)","必要な面談数"],["フォロー(CV)","必要な反応数"]];
    for(var s=0;s<stg.length;s++){ var rr=R+2+s;
      u.getRange(rr,5).setValue(stg[s][0]).setFontWeight("bold");
      u.getRange(rr,6).setFormula('='+ix(stg[s][1]));
      u.getRange(rr,7).setFormula('=ROUNDUP(F'+rr+'/'+remW+',0)');
      u.getRange(rr,8).setFormula('=ROUNDUP(F'+rr+'/'+remD+',0)');
    }
    u.getRange(R+5,5).setValue("残り").setFontWeight("bold").setFontColor(NAVY);
    u.getRange(R+5,6,1,3).merge().setFormula('='+remD+'&"日 / "&'+remW+'&"週(全日ベース)"').setFontColor(NAVY);
    u.getRange(R+1,5,5,4).setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    u.getRange(R+2,6,3,3).setBackground(WHITE);
  }

  // ── 金額ペース：勘定科目別 採用額の右 ──
  var rk=row("勘定科目別 採用額"); if(rk<0) rk=row("採用額 合計");
  if(rk>0){ var K=rk;
    u.getRange(K,5,1,4).merge().setValue("💴 金額ペース｜あといくら(今日/今週/月内)").setBackground(HEAD).setFontColor(WHITE).setFontWeight("bold").setFontSize(10);
    u.getRange(K+1,5,1,4).setValues([["","月内","今週","今日"]]).setBackground("#555555").setFontColor(WHITE).setFontWeight("bold").setFontSize(9);
    u.getRange(K+2,5).setValue("あと積む額").setFontWeight("bold");
    u.getRange(K+2,6).setFormula('='+ix("経常ギャップ"));
    u.getRange(K+2,7).setFormula('=ROUNDUP(F'+(K+2)+'/'+remW+',0)');
    u.getRange(K+2,8).setFormula('=ROUNDUP(F'+(K+2)+'/'+remD+',0)');
    u.getRange(K+2,6,1,3).setNumberFormat("#,##0").setBackground(WHITE);
    u.getRange(K+1,5,2,4).setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  }
  return {ok:true, driver行:dh, 採用額行:rk};
}
