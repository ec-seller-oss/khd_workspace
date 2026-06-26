/** 残高クイック入力 v2（推測で書かない安全版）2026-06-03
 * 反省：v1は「6月が空の口座は5月値を勝手に繰越プレフィル」→pushが推測値を書込む事故(行61 0→1674)。
 * v2方針：①入力E列は空 ②pushは“実際に打った数字だけ”書込み、空欄はBS放置（勝手に繰越しない）
 *         ③前月・当月BS現在値は「表示のみ」。
 * revertRB2June() … 事故の1セル(行61 6月)を0へ戻す（一度だけ） */
var BS_NAME='③ 資産負債（BS）';
var BANK_ROWS=[50,51,52,53,54,55,57,58,59,60,61,62,63];
var SRC_DEFAULT={
  50:'城北信金（窓口/電話・2ヶ月毎）',51:'法人SBI（ネット/MF）',52:'朝日信金（portal）',53:'朝日信金（portal/窓口）',
  54:'大東京信金（窓口/電話・2ヶ月毎）',55:'法人TB 東京ベイ（ネット）',57:'SBI 1（ネット/MF）',58:'みずほ（ネット/MF）',
  59:'ゆうちょ（ダイレクト/通帳）',60:'楽天銀行RB1（ネット/MF）',61:'楽天銀行RB2（ネット/MF）',62:'東京ベイTB 個人（ネット）',63:'現金（財布）'
};
function _bs(){ var ss=SpreadsheetApp.getActiveSpreadsheet(); var bs=ss.getSheetByName(BS_NAME); if(!bs) ss.getSheets().forEach(function(s){if(!bs&&s.getName().indexOf('BS')>=0)bs=s;}); if(!bs) throw new Error('BSタブ無し'); return bs; }
function _monthCol(bs){ var hdr=bs.getRange(2,1,1,bs.getLastColumn()).getValues()[0]; var now=new Date(); for(var c=0;c<hdr.length;c++){ var v=hdr[c]; if(v instanceof Date && v.getFullYear()==now.getFullYear() && v.getMonth()==now.getMonth()) return c+1; } throw new Error('当月('+(now.getMonth()+1)+'月)の列がBS行2に無い'); }
function _colL(i){ var s='',n=i; while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s; }

/** 事故の修正：行61(楽天RB2)の6月を0へ戻す（一度だけ実行） */
function revertRB2June(){
  var bs=_bs(), mcol=_monthCol(bs);
  var old=bs.getRange(61,mcol).getValue();
  bs.getRange(61,mcol).setValue(0);
  Logger.log('行61 6月('+_colL(mcol)+') を '+old+' → 0 に戻した');
  SpreadsheetApp.getActiveSpreadsheet().toast('行61(RB2) 6月を '+old+'→0 に戻しました','事故修正',6);
}

function buildBalanceQuickInput(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), bs=_bs();
  var mcol=_monthCol(bs), pcol=mcol-1;
  var data=bs.getRange(1,1,bs.getLastRow(),mcol).getValues();
  var t=ss.getSheetByName('🏦残高クイック入力') || ss.insertSheet('🏦残高クイック入力',0);
  t.clear();
  t.getRange(1,1,1,8).merge();
  t.getRange('A1').setValue('🏦 残高クイック入力（'+_colL(mcol)+'＝今月）— 見て確認した口座だけ「★今月入力」に数字を打つ → pushBalancesToBS。空欄はBS据置（勝手に繰越しない）')
    .setFontWeight('bold').setBackground('#FCE8B2');
  t.getRange(2,1,1,8).setValues([['BS行','口座','取得元（どこを見る）','前月残高','当月BS現在値','★今月入力(空=据置)','差分(入力-現在)','備考']])
    .setFontWeight('bold').setBackground('#DDDDDD');
  var out=[];
  BANK_ROWS.forEach(function(r){
    var nm=(data[r-1][1]||'')+'／'+(data[r-1][3]||'');
    var prev=data[r-1][pcol-1];   // 前月（表示のみ）
    var cur =data[r-1][mcol-1];   // 当月BS現在値（表示のみ）
    out.push([r, nm, SRC_DEFAULT[r]||'', prev, cur, '', '', '']); // ★入力(F)は空
  });
  t.getRange(3,1,out.length,8).setValues(out);
  var last=2+out.length;
  for(var i=0;i<out.length;i++){ var rr=3+i;
    t.getRange(rr,4).setNumberFormat('#,##0'); t.getRange(rr,5).setNumberFormat('#,##0'); t.getRange(rr,6).setNumberFormat('#,##0');
    t.getRange(rr,7).setFormula('=IF(F'+rr+'="","",F'+rr+'-E'+rr+')').setNumberFormat('+#,##0;▲#,##0;0');
  }
  t.getRange(last+1,2).setValue('銀行・現金 計');
  t.getRange(last+1,5).setFormula('=SUM(E3:E'+last+')').setNumberFormat('#,##0').setFontWeight('bold');
  t.getRange(last+1,6).setFormula('=SUM(F3:F'+last+')').setNumberFormat('#,##0').setFontWeight('bold');
  // 入力欄(F列)を薄黄で強調
  t.getRange(3,6,out.length,1).setBackground('#FFF7DB');
  t.getRange(last+3,1,1,8).merge();
  t.getRange(last+3,1).setValue('運用：①MF/通帳/portalで残高確認 ②確認した口座だけF列に実数を入力（見てない口座は空のまま＝BS据置・推測で埋めない）③pushBalancesToBS→BS当月列('+_colL(mcol)+')へ反映。信金(城北/朝日/大東京)は2ヶ月毎でOK。')
    .setWrap(true).setFontColor('#666666');
  t.setColumnWidth(1,40);t.setColumnWidth(2,170);t.setColumnWidth(3,230);t.setColumnWidth(4,105);t.setColumnWidth(5,115);t.setColumnWidth(6,130);t.setColumnWidth(7,110);t.setColumnWidth(8,160);
  t.setFrozenRows(2);
  Logger.log('残高クイック入力v2 作成: '+out.length+'口座 / 当月列='+_colL(mcol)+' / ★入力(F)は空・空欄は据置');
  ss.toast('残高クイック入力v2（推測なし）作成','残高半自動',6);
}

function pushBalancesToBS(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), bs=_bs();
  var t=ss.getSheetByName('🏦残高クイック入力'); if(!t) throw new Error('入力タブ無し。先にbuildBalanceQuickInput');
  var mcol=_monthCol(bs), n=BANK_ROWS.length;
  var vals=t.getRange(3,1,n,6).getValues(); // A:BS行 .. F:★今月入力
  var log=['=== pushBalancesToBS 当月列='+_colL(mcol)+'（実際に打った数字だけ書込・空欄は据置）==='], changed=0, skip=0;
  vals.forEach(function(row){
    var bsr=row[0], v=row[5]; // F列
    if(!bsr) return;
    if(v===''||v===null){ skip++; return; }      // 空欄＝据置（書かない）
    if(typeof v!=='number'){ log.push('行'+bsr+' 数値でないのでスキップ:'+v); return; }
    var old=bs.getRange(bsr,mcol).getValue();
    if(old!==v){ bs.getRange(bsr,mcol).setValue(v); log.push('行'+bsr+' '+old+' → '+v); changed++; }
  });
  log.push('書込'+changed+'件 / 空欄据置'+skip+'件');
  Logger.log(log.join('\n'));
  ss.toast('書込'+changed+'件・据置'+skip+'件（当月'+_colL(mcol)+'）','残高確定',6);
}
