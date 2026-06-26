// 案件名統一連動：03案件名(C5:C14)を「顧客マスター名＋現03案件」のドロップダウンに(表記ゆれ撲滅)
function buildAnkenDropdown(){
  var ss=SpreadsheetApp.openById("1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc");
  function shk(kw){ var a=ss.getSheets(); for(var i=0;i<a.length;i++){ if(a[i].getName().indexOf(kw)>=0) return a[i]; } return null; }
  var u=shk("売上見込み"); if(!u) return {error:"03無し"};
  var km=shk("顧客");
  var list=[], seen={};
  function add(v){ v=(v==null?"":v).toString().trim(); if(v && !seen[v]){ seen[v]=1; list.push([v]); } }
  // ① 03の現案件(C5:C14)
  var c=u.getRange("C5:C14").getValues(); for(var i=0;i<c.length;i++) add(c[i][0]);
  // ② 顧客マスターの名前列(ヘッダから探す)
  if(km){
    var lc=km.getLastColumn(), hd=km.getRange(1,1,Math.min(5,km.getLastRow()),lc).getValues(), hRow=-1, nmCol=-1;
    for(var r=0;r<hd.length;r++){ for(var cc=0;cc<lc;cc++){ var h=(hd[r][cc]||"").toString(); if(h.indexOf("名前")>=0||h.indexOf("顧客名")>=0){ hRow=r+1; nmCol=cc+1; break; } } if(hRow>0) break; }
    if(nmCol>0 && km.getLastRow()>hRow){
      var nm=km.getRange(hRow+1,nmCol,km.getLastRow()-hRow,1).getValues();
      for(var r=0;r<nm.length;r++) add(nm[r][0]);
    }
  }
  // ③ 隠しリストタブ _案件名M に書く
  var hm=ss.getSheetByName("_案件名M"); if(hm) ss.deleteSheet(hm); hm=ss.insertSheet("_案件名M");
  hm.getRange(1,1).setValue("案件名マスター(自動生成・03ドロップダウン用／顧客M＋現03案件)");
  if(list.length>0) hm.getRange(2,1,list.length,1).setValues(list);
  // ④ 03!C5:C14 にドロップダウン(allowInvalid=true＝既存自由入力も可)
  var rule=SpreadsheetApp.newDataValidation().requireValueInRange(hm.getRange(2,1,Math.max(list.length,1),1),true).setAllowInvalid(true).build();
  u.getRange("C5:C14").setDataValidation(rule);
  hm.hideSheet();
  return {ok:true, 候補数:list.length, note:"03案件名がドロップダウン化。02確度報告時は同じ名前を選ぶ＝連動が切れない"};
}
