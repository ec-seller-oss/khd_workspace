/** MF実数→BS一括反映（連携が正常で1:1対応する口座のみ）2026-06-03読取
 * 値はClaudeがMF画面から読取った実数。月次は「MF再読込→このMF配列を差し替え→再実行」。
 * 当月列＝今日の月で自動判定。変わった行だけ書込み＋ログ。B（集約不能/連携不調/MF外/手動）は含めない＝触らない。 */
var BS_NAME_MF='③ 資産負債（BS）';
// [BS行, 口座名, MF実数, MF取得日] ※2026-06-03時点のMF ME読取
var MF_VALUES=[
  [50,'法人/城北信金',281012,'06/01'],
  [51,'法人/法人SBI(住信)',139320,'06/01'],
  [141,'麻梨奈/SBI銀行',1295429,'06/01'],
  [142,'麻梨奈/SBI証券',2898367,'06/01'],
  [57,'研太/SBI1',1224610,'06/01'],
  [59,'研太/ゆうちょ',137132,'06/01'],
  [60,'研太/RB1',92400,'06/01'],
  [61,'研太/RB2',0,'06/03'],
  [62,'研太/東京ベイTB',404939,'06/02'],
  [75,'研太/SBI証券',0,'06/01'],
  [91,'マイル/ANA1',59431,'06/01'],
  [92,'マイル/ANA2',52182,'06/01'],
  [93,'マイル/JAL',181045,'06/01'],
  [102,'P/楽天市場',63535,'06/03'],
  [103,'P/楽天市場2',19,'06/01'],
  [109,'P/モッピー',37034,'06/01'],
  [114,'P/メルペイ',8725,'06/03']
];
function _bsMf(){ var ss=SpreadsheetApp.getActiveSpreadsheet(); var bs=ss.getSheetByName(BS_NAME_MF); if(!bs) ss.getSheets().forEach(function(s){if(!bs&&s.getName().indexOf('BS')>=0)bs=s;}); if(!bs) throw new Error('BSタブ無し'); return bs; }
function _monthColMf(bs){ var hdr=bs.getRange(2,1,1,bs.getLastColumn()).getValues()[0]; var now=new Date(); for(var c=0;c<hdr.length;c++){ var v=hdr[c]; if(v instanceof Date && v.getFullYear()==now.getFullYear() && v.getMonth()==now.getMonth()) return c+1; } throw new Error('当月の列がBS行2に無い'); }
function _colLMf(i){ var s='',n=i; while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s; }

function applyMfBalances(){
  var bs=_bsMf(), mcol=_monthColMf(bs);
  var log=['=== applyMfBalances 当月='+_colLMf(mcol)+'（MF実数2026-06-03読取・正常連携'+MF_VALUES.length+'口座）==='], ch=0, same=0;
  MF_VALUES.forEach(function(m){
    var old=bs.getRange(m[0],mcol).getValue();
    if(old!==m[2]){ bs.getRange(m[0],mcol).setValue(m[2]); log.push('✏️行'+m[0]+' '+m[1]+'：'+old+' → '+m[2]+'（MF'+m[3]+'）'); ch++; }
    else same++;
  });
  log.push('— 更新'+ch+'件 / 既に一致'+same+'件（計'+MF_VALUES.length+'口座）');
  log.push('※連携不調(みずほ/V/PayPay/UA)・集約不能(RS/野村/Amazon)・MF外(朝日/大東京/法人TB)・手動(現金/d払い)は対象外＝据置');
  Logger.log(log.join('\n'));
  SpreadsheetApp.getActiveSpreadsheet().toast('MF実数を反映：更新'+ch+'件/一致'+same+'件','MF自動反映',7);
}
