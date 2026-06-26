/**
 * 02作業DB 報告タブ化リフォーム（位置を動かさず＝O/P/Q/R・03連動を壊さない）
 * 使い方: script.new → 貼付 → redesignDb02 実行
 * 変更点:
 *  ① 温度→確度 に改名＋値変換(HOT→A等)＋プルダウン(A/B/C/D/ブレイク)
 *  ② 右端に新規3列: 報告相手(プルダウン)/相談したいこと/③相談後に決めること(穴埋め)
 *  ③ 不要列を非表示: 実所要分(旧重複)・報告項目・報告値(03連動の内部)
 *  ※報告項目/報告値は03連動が使うので削除せず非表示のみ。
 */
function redesignDb02(){
  var ss=SpreadsheetApp.openById('1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc');
  var sh=findDb02_(ss); if(!sh){throw '02が見つからない';}
  var hRow=findHeaderRow_(sh); var lastCol=sh.getLastColumn();
  var H=sh.getRange(hRow,1,1,lastCol).getValues()[0];
  function col(name){ for(var i=0;i<H.length;i++) if(String(H[i]).trim()===name) return i+1; return -1; }

  // ① 温度→確度ランク（※S/T列の「確度%(03連動)」と名前衝突させない）
  var g=col('温度');
  if(g>0){
    sh.getRange(hRow,g).setValue('確度ランク(A/B/C/D)');
    var last=sh.getLastRow();
    if(last>hRow){
      var rng=sh.getRange(hRow+1,g,last-hRow,1); var v=rng.getValues();
      var m={'HOT':'A','WARM':'B','COLD':'C','🔥':'A','😐':'C','😣':'D','⚡':'C','🛟':'D'};
      for(var i=0;i<v.length;i++){ var x=String(v[i][0]).trim();
        if(['A','B','C','D','ブレイク'].indexOf(x)>=0) continue;
        v[i][0]= m[x]!==undefined ? m[x] : '';
      }
      rng.setValues(v);
      var dv=SpreadsheetApp.newDataValidation().requireValueInList(['A','B','C','D','ブレイク'],true).build();
      rng.setDataValidation(dv);
    }
  }

  // ② 右端に新規3列
  var start=sh.getLastColumn();
  sh.insertColumnsAfter(start,3);
  var c1=start+1,c2=start+2,c3=start+3;
  sh.getRange(hRow,c1,1,3).setValues([['報告相手','相談したいこと','③相談後に決めること(穴埋め)']])
    .setFontWeight('bold').setBackground('#AA2E26').setFontColor('#FFFFFF').setWrap(true);
  var lastRow=sh.getLastRow();
  if(lastRow>hRow){
    var dv2=SpreadsheetApp.newDataValidation().requireValueInList(['福井','羽鳥','宮崎','バイセル','チーム','家族','—'],true).build();
    sh.getRange(hRow+1,c1,lastRow-hRow,1).setDataValidation(dv2).setBackground('#FFF2CC');
    sh.getRange(hRow+1,c2,lastRow-hRow,2).setBackground('#FFF2CC');
  }
  sh.setColumnWidth(c1,90); sh.setColumnWidth(c2,260); sh.setColumnWidth(c3,260);

  // ③ 不要列を非表示 ※報告項目/報告値(S/T)は03連動の心臓なので絶対に触らない（非表示にもしない）
  ['実所要分'].forEach(function(n){ var c=col(n); if(c>0) sh.hideColumns(c); });

  SpreadsheetApp.flush();
  Logger.log('redesign完了: 確度化＋報告3列追加＋不要列非表示 '+ss.getUrl());
}

function findDb02_(ss){
  var shs=ss.getSheets();
  for(var i=0;i<shs.length;i++){
    var r=Math.min(5,shs[i].getLastRow()); if(r<1) continue;
    var vals=shs[i].getRange(1,1,r,Math.min(30,shs[i].getLastColumn())).getValues();
    for(var x=0;x<vals.length;x++){ var row=vals[x].join('|');
      if(row.indexOf('案件・相手')>=0 && row.indexOf('日付')>=0) return shs[i];
    }
  }
  return null;
}
function findHeaderRow_(sh){
  var r=Math.min(5,sh.getLastRow());
  var vals=sh.getRange(1,1,r,Math.min(30,sh.getLastColumn())).getValues();
  for(var x=0;x<vals.length;x++) if(vals[x].join('|').indexOf('案件・相手')>=0) return x+1;
  return 1;
}
