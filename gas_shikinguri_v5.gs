/** ‘¹‰vE—\‘ªŒn‚ÌŽdã‚°F…CŽ‘‹àŒJ‚è‚ðv5‰»{ƒeƒ“ƒvƒŒ3–‡‘Þ–ð 2026-06-02 */
var YEN5='#,##0"‰~"';
function runV5All(){ buildShikinguriV5(); retirePlanTemplates();
  SpreadsheetApp.getActiveSpreadsheet().toast('…CŽ‘‹àŒJ‚è‚ðv5‰»{ƒeƒ“ƒvƒŒ3–‡‚ð‘Þ–ð‚µ‚Ü‚µ‚½','Š®—¹',6); }

function buildShikinguriV5(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('…C Ž‘‹àŒJ‚è'); if(!sh){ sh=ss.insertSheet('…C Ž‘‹àŒJ‚è'); } else { sh.clear(); }
  var months=['2026/06','2026/07','2026/08','2026/09','2026/10','2026/11','2026/12','2027/01'];
  var N=months.length; var cols=[]; for(var i=0;i<N;i++) cols.push(String.fromCharCode(66+i)); // B..I
  function setRow(r, label, vals){ // vals: ”z—ñ(”’l) or null
    sh.getRange(r,1).setValue(label);
    if(vals){ for(var i=0;i<N;i++){ sh.getRange(r,2+i).setValue(vals[i]).setNumberFormat(YEN5); } }
  }
  function setF(r, fns){ for(var i=0;i<N;i++){ sh.getRange(r,2+i).setFormula(fns[i]).setNumberFormat(YEN5); } }
  // ƒwƒbƒ_
  sh.getRange(1,1).setValue('Ž‘‹àŒJ‚èv5iBS˜A“®E¶Šˆ”ï/ˆç‹x”½‰fEÜ—^0j\ ŒŽ‰Œ»‹à‚Í…@Ži—ß“ƒ˜A“®').setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#AA2E26');
  sh.getRange(1,2,1,N).setBackground('#AA2E26');
  sh.getRange(2,1).setValue('€–Ú').setFontWeight('bold').setBackground('#F0E2DF');
  for(var i=0;i<N;i++) sh.getRange(2,2+i).setValue(months[i]).setFontWeight('bold').setBackground('#F0E2DF').setHorizontalAlignment('center');
  // r3 ŒŽ‰Œ»—a‹àF6ŒŽ=…@Ži—ß“ƒ!B10(BS˜A“®)AˆÈ~=‘OŒŽ––(r29)
  sh.getRange(3,1).setValue('ŒŽ‰Œ»—a‹àŽc‚iBS˜A“®j').setFontWeight('bold');
  sh.getRange(3,2).setFormula("='…@“‡Ži—ß“ƒ'!B10").setNumberFormat(YEN5);
  for(var i=1;i<N;i++) sh.getRange(3,2+i).setFormula('='+cols[i-1]+'29').setNumberFormat(YEN5);
  // “ü‹à
  sh.getRange(4,1).setValue('¡“ü‹à').setFontWeight('bold').setBackground('#FCEFE7');
  setRow(5,'•¨Œ”„‹pi‰h’¬6/20EŠm“x90%j',[1900000,0,0,0,0,0,0,0]);
  setRow(6,'ˆã—ÃƒRƒ“ƒTƒ‹iŽè“ü—Í/ƒpƒCƒvƒ‰ƒCƒ“j',[0,0,0,0,0,0,0,0]);
  setRow(7,'”ƒŽæÄ”ÌE‰¡•lNPO',[0,0,0,0,0,0,0,0]);
  setRow(8,'EC‘e—˜ ƒN[ƒpƒ“1',[100000,100000,100000,100000,100000,100000,100000,100000]);
  setRow(9,'EC‘e—˜ ƒN[ƒpƒ“2',[50000,50000,50000,50000,50000,50000,50000,50000]);
  setRow(10,'š–ƒ—œ“Þ ˆç‹x‹‹•tiŒŽŠ„j',[156659,156659,156659,156659,156659,156659,156659,156659]);
  setRow(11,'‚»‚Ì‘¼“ü‹à',[0,0,0,0,0,0,0,0]);
  setF(12, cols.map(function(c){return '=SUM('+c+'5:'+c+'11)';})); sh.getRange(12,1).setValue('“ü‹à‡Œv').setFontWeight('bold');
  // o‹à
  sh.getRange(13,1).setValue('¡o‹à').setFontWeight('bold').setBackground('#FCEFE7');
  setRow(14,'–@lFÅ—ŽmŒÚ–âiZH‹´–{j',rep(110000,N));
  setRow(15,'–@lFŠO’i]“¡j',rep(99000,N));
  setRow(16,'–@lFŽÐ‰ï•ÛŒ¯',rep(22167,N));
  setRow(17,'ŒÂlŽ–‹ÆFÅ—ŽmiZHj',rep(22000,N));
  setRow(18,'ƒvƒ‰ƒCƒx[ƒgFZ‘îƒ[ƒ“',rep(130668,N));
  setRow(19,'ƒvƒ‰ƒCƒx[ƒgFSMBCŠÇ—”ï',rep(22130,N));
  setRow(20,'ƒvƒ‰ƒCƒx[ƒgFPayPay',rep(10000,N));
  setRow(21,'š¢‘Ñ¶Šˆ”ïiŠy“V18+–ƒ—œ“Þ12.1j',rep(301000,N));
  setRow(22,'š–@lFŽØ“ü•ÔÏié–k+’©“ú+‘å“Œ‹žj',rep(76035,N));
  setRow(23,'šŒÂlŽ–‹ÆFŽØ“ü•ÔÏi“Œ‹žƒxƒC2+ŒöŒÉj',rep(145001,N));
  setRow(24,'“ŠMÏ—§i”CˆÓEŽ~‚ß‰Âj',rep(100000,N));
  setRow(25,'–ðˆõÜ—^iŒ©‘—‚èŠm’è=0j',rep(0,N));
  setRow(26,'‚»‚Ì‘¼o‹à',rep(0,N));
  setF(27, cols.map(function(c){return '=SUM('+c+'14:'+c+'26)';})); sh.getRange(27,1).setValue('o‹à‡Œv').setFontWeight('bold');
  // ƒ‘Œ¸EŒŽ––
  setF(28, cols.map(function(c){return '='+c+'12-'+c+'27';})); sh.getRange(28,1).setValue('“–ŒŽƒ‘Œ¸i“ü‹à-o‹àj').setFontWeight('bold');
  setF(29, cols.map(function(c){return '='+c+'3+'+c+'28';})); sh.getRange(29,1).setValue('ŒŽ––Œ»—a‹àŽc‚i‘Ì—ÍŒvj').setFontWeight('bold').setBackground('#DDF3DD');
  for(var i=0;i<N;i++) sh.getRange(29,2+i).setBackground('#DDF3DD');
  // ƒŒŽŽŸ”RÄiÏ—§Ž~‚ßƒx[ƒXjo‹à | EC | ˆç‹x | Ï—§
  setF(31, cols.map(function(c){return '='+c+'27-'+c+'8-'+c+'9-'+c+'10-'+c+'24';})); sh.getRange(31,1).setValue('ƒŒŽŽŸ”RÄiÏ—§Ž~‚ßEŒoíj');
  // ƒ‰ƒ“ƒEƒFƒCi6ŒŽŒŽ‰€”RÄj{M†
  sh.getRange(32,1).setValue('ƒ‰ƒ“ƒEƒFƒCiŒŽ‰Œ»‹à€”RÄj').setFontWeight('bold');
  sh.getRange(32,2).setFormula('=B3/B31').setNumberFormat('0.0"ƒ–ŒŽ"').setFontWeight('bold');
  sh.getRange(32,3).setFormula('=IF(B32<3,"??Žç‚è:“ŠŽ‘“€Œ‹",IF(B32<6,"??’ˆÓ","??U‚ßOK"))').setFontWeight('bold').setBackground('#DDF3DD');
  sh.getRange(34,1).setValue('yŽg‚¢•ûzŒŽ‰Œ»‹à=…@Ži—ß“ƒ(BS)˜A“®B‰h’¬/ˆã—Ã‚ÍŽè“ü—ÍorƒpƒCƒvƒ‰ƒCƒ“B¶Šˆ”ï/ˆç‹x/Ï—§‚ðŽÀŠz”½‰fÏiv4‚Ìu12ŒŽƒVƒ‡[ƒgv‚ÍŒŽ‰Œ»‹à‰ß¬‚ÌŒë•ñ¨“P‰ñjB').setWrap(true);
  sh.getRange(34,1,1,9).merge();
  sh.setColumnWidth(1,30*9); for(var i=0;i<N;i++) sh.setColumnWidth(2+i,90);
}
function rep(v,n){ var a=[]; for(var i=0;i<n;i++)a.push(v); return a; }

function retirePlanTemplates(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var T=['‘¹‰vŒv‰æ','—\ŽZ','—\ŽÀ']; var del=[];
  T.forEach(function(n){ var s=ss.getSheetByName(n); if(s){ ss.deleteSheet(s); del.push(n);} });
  Logger.log('‘Þ–ð: '+(del.join(' / ')||'‚È‚µ'));
}
