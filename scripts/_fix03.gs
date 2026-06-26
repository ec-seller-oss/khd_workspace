// 03修正：EC(行12)・賃料(行13)の確度が空→固定毎月収入なので100%に戻す＋着金=毎月。採用額/期待値が復活。
function fix03Recurring(){
  var ss=SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  var u=null,a=ss.getSheets();
  for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf("売上見込み")>=0){ u=a[i]; break; } }
  if(!u) return "03無し";
  // F=確度% / G=着金見込月 / H弱気 I現実 J強気 K採用額 M期待値
  [12,13].forEach(function(r){
    u.getRange("F"+r).setValue(1);          // 確度100%
    if(!u.getRange("G"+r).getValue()) u.getRange("G"+r).setValue("毎月");
  });
  SpreadsheetApp.flush();
  // 採用額/期待値が値で固定の行なら確度連動しないので、満額を弱気/現実/強気/採用額/期待値へ直接補完(EC61000・賃料81000)
  var amt={12:61000,13:81000};
  [12,13].forEach(function(r){
    var man=u.getRange("E"+r).getValue(); var v=Number(man)||amt[r];
    // 採用額(K)・期待値(M)が数式でなく0なら満額で補完
    if(!u.getRange("K"+r).getFormula()){ u.getRange("H"+r).setValue(v); u.getRange("I"+r).setValue(v); u.getRange("J"+r).setValue(v); u.getRange("K"+r).setValue(v); }
    if(!u.getRange("M"+r).getFormula()) u.getRange("M"+r).setValue(v);
  });
  SpreadsheetApp.flush();
  // 検証読み戻し
  var out=[];
  [3,12,13].forEach(function(r){ out.push("R"+r+" F="+u.getRange("F"+r).getDisplayValue()+" K(採用額)="+u.getRange("K"+r).getDisplayValue()+" M(期待値)="+u.getRange("M"+r).getDisplayValue()); });
  return {ok:true, 確認:out};
}
