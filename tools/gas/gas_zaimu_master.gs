/** 💰財務マスター（定型化）2026-06-03 — これを一度だけ貼れば、以後はシート上部メニュー「💰財務」から全操作。
 * 毎回のコピペ→保存→関数選択→実行 を撲滅。MF残高は📥MF残高タブを私がChromeで更新→あなたはメニュー1クリック。 */

function onOpen(){
  var ui=SpreadsheetApp.getUi();
  ui.createMenu('💰財務')
    .addItem('📸 今日のスナップショット記録', 'snapshotToday')
    .addSeparator()
    .addItem('📥 MF残高 → BSへ反映', 'applyMfBalances')
    .addItem('🔧 MF残高タブを作成/初期化', 'setupMfTab')
    .addSeparator()
    .addItem('🏦 手動残高 入力タブ作成', 'buildBalanceQuickInput')
    .addItem('✅ 手動入力 → BSへ反映', 'pushBalancesToBS')
    .addSeparator()
    .addItem('🏷 連携区分タグ 再付与', 'tagBalanceSource')
    .addItem('🔗 EC粗利アンカー更新(EC側)', 'buildEcAnchor')
    .addItem('🧾 PLをEC実績に連動(一度だけ)', 'linkPLtoEC')
    .addItem('👔 役員報酬5万をPLに反映(未払)', 'setYakuinHoshu')
    .addItem('🧮 損益PL 再整形', 'redesignPL')
    .addSeparator()
    .addItem('🔗 BS負債を⑤借入へ連動(一度だけ)', 'linkBSdebtToBorrow')
    .addSeparator()
    .addItem('🔎 検算: PL/資金繰りログ', 'verifyPL')
    .addItem('🔍 確認: 借入/タブ構造ログ', 'inspectBorrow')
    .addToUi();
  // tougou_dashboard.gs の関数も同じonOpenから（向こうのonOpenは無効化する）
  ui.createMenu('📊統合ダッシュボード')
    .addItem('② 本部マトリクス(実態時間版)を再構築', 'buildMatrixV2')
    .addItem('② 実績hをカレンダーから更新', 'updateActualHoursV2')
    .addSeparator()
    .addItem('（旧）全構築①②⑥', 'buildAll')
    .addToUi();
}

/* ===== 📸日次スナップショット（押した日で集計＋日付列を右端に追記） =====
 * 既存の月列は無傷。右端に「今日の日付」列を追記＝列挿入しないので数式は壊れない。
 * 集計行(総資産/自己資本/月末残高)は列ごと数式の計算結果を凍結値でコピー＝その日の集計が固定で残る。
 * 月列(BI)は最新を維持→①司令塔・④資金繰りは触らず自動で今日を表示。 */
function snapshotToday(){ applyMfBalances(); appendDailySnapshot(); }
function appendDailySnapshot(){
  var bs=_bs(), mcol=_mcol(bs), lr=bs.getLastRow();
  var newCol=bs.getLastColumn()+1;
  var vals=bs.getRange(1,mcol,lr,1).getValues();   // 当月列の計算結果を凍結値で取得
  bs.getRange(1,newCol,lr,1).setValues(vals);
  var today=new Date();
  bs.getRange(2,newCol).setValue(today).setNumberFormat('m/d').setFontWeight('bold').setBackground('#D9EAD3');
  bs.setColumnWidth(newCol,80);
  var sumAsset=bs.getRange(36,newCol).getValue(), nw=bs.getRange(44,newCol).getValue();
  Logger.log('📸日次スナップショット列 '+_colL(newCol)+' = '+Utilities.formatDate(today,'JST','yyyy/MM/dd')+' / 総資産'+sumAsset+' 純資産'+nw);
  _toast('日次スナップショット列 '+_colL(newCol)+' 追加（総資産'+Math.round(sumAsset/10000)+'万）','📸スナップ');
}

/* ===== 共通ヘルパー ===== */
function _bs(){ var ss=SpreadsheetApp.getActiveSpreadsheet(); var bs=ss.getSheetByName('③ 資産負債（BS）'); if(!bs) ss.getSheets().forEach(function(s){if(!bs&&s.getName().indexOf('BS')>=0)bs=s;}); if(!bs) throw new Error('BSタブ無し'); return bs; }
function _mcol(bs){ var hdr=bs.getRange(2,1,1,bs.getLastColumn()).getValues()[0]; var now=new Date(); for(var c=0;c<hdr.length;c++){ var v=hdr[c]; if(v instanceof Date && v.getFullYear()==now.getFullYear() && v.getMonth()==now.getMonth()) return c+1; } throw new Error('当月の列がBS行2に無い'); }
function _colL(i){ var s='',n=i; while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s; }
function _find(kw,not){ var ss=SpreadsheetApp.getActiveSpreadsheet(),r=null; ss.getSheets().forEach(function(s){var n=s.getName(); if(!r&&n.indexOf(kw)>=0&&(!not||n.indexOf(not)<0))r=s;}); return r; }
function _toast(m,t){ SpreadsheetApp.getActiveSpreadsheet().toast(m,t||'💰財務',6); }

/* ===== MF残高（📥MF残高タブ経由・月次はChromeで値だけ更新） ===== */
var MF_SEED=[ // [BS行,口座,MF残高,取得日] ※初期=2026-06-03読取
 [50,'法人/城北信金',281012,'2026-06-03'],[51,'法人/法人SBI',139320,'2026-06-03'],
 [141,'麻梨奈/SBI銀行',1295429,'2026-06-03'],[142,'麻梨奈/SBI証券',2898367,'2026-06-03'],
 [57,'研太/SBI1',1224610,'2026-06-03'],[59,'研太/ゆうちょ',137132,'2026-06-03'],
 [60,'研太/RB1',92400,'2026-06-03'],[61,'研太/RB2',0,'2026-06-03'],[62,'研太/東京ベイTB',404939,'2026-06-03'],
 [75,'研太/SBI証券',0,'2026-06-03'],[91,'マイル/ANA1',59431,'2026-06-03'],[92,'マイル/ANA2',52182,'2026-06-03'],
 [93,'マイル/JAL',181045,'2026-06-03'],[102,'P/楽天市場',63535,'2026-06-03'],[103,'P/楽天市場2',19,'2026-06-03'],
 [109,'P/モッピー',37034,'2026-06-03'],[114,'P/メルペイ',8725,'2026-06-03']
];
function setupMfTab(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var t=ss.getSheetByName('📥MF残高')||ss.insertSheet('📥MF残高',0);
  t.clear();
  t.getRange(1,1,1,4).setValues([['BS行','口座','MF残高(Chrome更新)','取得日']]).setFontWeight('bold').setBackground('#DDDDDD');
  t.getRange(2,1,MF_SEED.length,4).setValues(MF_SEED);
  t.getRange(2,3,MF_SEED.length,1).setNumberFormat('#,##0').setBackground('#FFF7DB');
  t.getRange(MF_SEED.length+3,1).setValue('使い方：月初に私(Claude)がMFを読んでC列(MF残高)とD列(取得日)を更新→あなたはメニュー「📥MF残高→BSへ反映」を1クリック。C列が空の行は据置(書かない)。')
    .setWrap(true).setFontColor('#666666');
  t.getRange(MF_SEED.length+3,1,1,4).merge();
  t.setColumnWidth(1,50);t.setColumnWidth(2,170);t.setColumnWidth(3,150);t.setColumnWidth(4,110); t.setFrozenRows(1);
  _toast('📥MF残高タブ作成('+MF_SEED.length+'口座)','MF定型化');
}
function applyMfBalances(){
  var bs=_bs(), mcol=_mcol(bs);
  var t=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📥MF残高');
  if(!t) throw new Error('📥MF残高タブ無し。先にメニュー「MF残高タブを作成/初期化」');
  var rows=t.getRange(2,1,t.getLastRow()-1,3).getValues();
  var log=['=== MF→BS '+_colL(mcol)+'列（C列の実数だけ書込・空は据置）==='], ch=0, sk=0;
  rows.forEach(function(r){ var bsr=r[0], v=r[2]; if(!bsr) return; if(v===''||v===null){sk++;return;} if(typeof v!=='number'){log.push('行'+bsr+' 非数値skip:'+v);return;}
    var old=bs.getRange(bsr,mcol).getValue(); if(old!==v){ bs.getRange(bsr,mcol).setValue(v); log.push('✏️行'+bsr+' '+old+'→'+v); ch++; } });
  log.push('更新'+ch+'件 / 据置'+sk+'件');
  Logger.log(log.join('\n')); _toast('MF反映：更新'+ch+'件','MF→BS');
}

/* ===== 手動残高（MF外口座・銀行13＋必要分） ===== */
var BANK_ROWS=[50,51,52,53,54,55,57,58,59,60,61,62,63];
var SRC={50:'城北信金(MF自動)',51:'法人SBI(MF自動)',52:'朝日信金(portal・2ヶ月毎)',53:'朝日信金(portal/窓口)',54:'大東京(窓口/電話・2ヶ月毎)',55:'法人TB東京ベイ(MF外)',57:'SBI1(MF自動)',58:'みずほ(MF連携エラー)',59:'ゆうちょ(MF自動)',60:'RB1(MF自動)',61:'RB2(MF自動)',62:'東京ベイTB(MF自動)',63:'現金(財布実数・MF無視)'};
function buildBalanceQuickInput(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), bs=_bs(), mcol=_mcol(bs), pcol=mcol-1;
  var data=bs.getRange(1,1,bs.getLastRow(),mcol).getValues();
  var t=ss.getSheetByName('🏦残高クイック入力')||ss.insertSheet('🏦残高クイック入力',0);
  t.clear();
  t.getRange(1,1,1,7).merge(); t.getRange('A1').setValue('🏦 手動残高入力('+_colL(mcol)+'＝今月)— 見た口座だけF列に実数→メニュー「手動入力→BSへ反映」。空=据置').setFontWeight('bold').setBackground('#FCE8B2');
  t.getRange(2,1,1,7).setValues([['BS行','口座','取得元','前月','当月BS現在','★今月入力(空=据置)','差分']]).setFontWeight('bold').setBackground('#DDDDDD');
  var out=[]; BANK_ROWS.forEach(function(r){ out.push([r,(data[r-1][1]||'')+'／'+(data[r-1][3]||''),SRC[r]||'',data[r-1][pcol-1],data[r-1][mcol-1],'','']); });
  t.getRange(3,1,out.length,7).setValues(out); var last=2+out.length;
  for(var i=0;i<out.length;i++){var rr=3+i; t.getRange(rr,4,1,2).setNumberFormat('#,##0'); t.getRange(rr,6).setNumberFormat('#,##0'); t.getRange(rr,7).setFormula('=IF(F'+rr+'="","",F'+rr+'-E'+rr+')').setNumberFormat('+#,##0;▲#,##0;0');}
  t.getRange(3,6,out.length,1).setBackground('#FFF7DB'); t.setFrozenRows(2);
  t.setColumnWidth(1,40);t.setColumnWidth(2,170);t.setColumnWidth(3,210);t.setColumnWidth(4,100);t.setColumnWidth(5,110);t.setColumnWidth(6,130);t.setColumnWidth(7,100);
  _toast('🏦手動残高入力タブ作成','手動残高');
}
function pushBalancesToBS(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), bs=_bs(), mcol=_mcol(bs);
  var t=ss.getSheetByName('🏦残高クイック入力'); if(!t) throw new Error('入力タブ無し');
  var vals=t.getRange(3,1,BANK_ROWS.length,6).getValues();
  var log=['=== 手動→BS '+_colL(mcol)+'列 ==='], ch=0, sk=0;
  vals.forEach(function(r){ var bsr=r[0],v=r[5]; if(!bsr)return; if(v===''||v===null){sk++;return;} if(typeof v!=='number'){log.push('行'+bsr+' 非数値skip');return;}
    var old=bs.getRange(bsr,mcol).getValue(); if(old!==v){bs.getRange(bsr,mcol).setValue(v);log.push('✏️行'+bsr+' '+old+'→'+v);ch++;} });
  log.push('更新'+ch+'件/据置'+sk+'件'); Logger.log(log.join('\n')); _toast('手動反映：更新'+ch+'件','手動→BS');
}

/* ===== 連携区分タグ（BL=連携/BM=理由・50〜lastRow） ===== */
function tagBalanceSource(){
  var bs=_bs(), BL=64, BM=65, lr=bs.getLastRow();
  var M={50:['🟢自動','MF連携(城北信金)'],51:['🟢自動','MF連携(住信SBI)'],52:['🟡手動','MF外・朝日portal・2ヶ月毎'],53:['🟡手動','MF外・朝日portal/窓口'],54:['🟡手動','MF外・大東京 窓口/電話・2ヶ月毎'],55:['🟡手動','MF外・法人TB ネット'],56:['🟡手動','保険・低頻度'],57:['🟢自動','MF連携(SBI銀行)'],58:['🟡手動','MF連携エラー(要再設定)'],59:['🟢自動','MF連携(ゆうちょ)'],60:['🟢自動','MF連携(RB1)'],61:['🟢自動','MF連携(RB2)'],62:['🟢自動','MF連携(東京ベイTB)'],63:['🟡手動','財布実数・MF無視(BSが正)'],
    69:['🟡手動','MFはRS1集約・BS銘柄別'],70:['🟡手動','MFはRS1集約・BS銘柄別'],71:['🟡手動','MFはRS1集約・BS銘柄別'],72:['🟡手動','MFはRS1集約・BS銘柄別'],73:['🟡手動','MFはRS1集約・BS銘柄別'],74:['🟡手動','MFはRS1集約・BS銘柄別'],75:['🟢自動','MF連携(SBI証券)'],76:['🟡手動','MF集約(RS2)'],79:['🟡手動','MFと差異(野村)・要確認'],80:['🟡手動','DC・低頻度'],
    82:['🟡手動','アマギフ・MF外/税金積立'],83:['🟡手動','アマギフ・MF外'],84:['🟡手動','Apple・MF外'],85:['🟡手動','Apple・MF外'],86:['🟡手動','PayPay・MF要OTP'],87:['🟡手動','PayPay・MF要OTP'],88:['🟡手動','JAL Pay・MF外'],89:['🟡手動','WAON・MF外'],90:['🟡手動','楽天キャッシュ・要楽天ログイン'],
    91:['🟢自動','MF連携(ANAマイ1)'],92:['🟢自動','MF連携(ANAマイ2)'],93:['🟢自動','MF連携(JALマイル)'],94:['🟡手動','MF一時停止(UA)'],95:['🟡手動','AGP・MF外'],96:['🟡手動','AGP2・MF外'],97:['🟡手動','えきねっと/JRE・MF取得不可'],98:['🟡手動','MB・MF外'],99:['🟡手動','MB無料宿泊・MF外'],
    100:['🟡手動','Amazonポイント・MF集約と不一致'],101:['🟡手動','Amazonポイント2・MF集約'],102:['🟢自動','MF連携(楽天市場)'],103:['🟢自動','MF連携(楽天市場2)'],104:['🟡手動','B1・MF外'],107:['🟡手動','ANAソラチカ・MF外'],108:['🟡手動','JCB Okidoki・MF外'],109:['🟢自動','MF連携(モッピー)'],110:['🟡手動','モッピー麻梨奈・要妻確認'],111:['🟡手動','Vポイント・MF要OTP'],112:['🟡手動','d払い・解約予定6月末'],114:['🟢自動','MF連携(メルペイ)'],115:['🟡手動','メルカリ2・MF外'],116:['🟡手動','MF外'],117:['🟡手動','MF外'],
    124:['🟡手動','現物メルカリ出品中'],125:['🟡手動','寝かせ商品ポケカ等'],126:['🟡手動','FBA代行'],127:['🟡手動','Amazon倉庫'],128:['🟡手動','保留金'],129:['🟡手動','売掛金ペイオニア1'],130:['🟡手動','売掛金ペイオニア2'],
    131:['🟡手動','売掛金スマスト・EC手動'],132:['🟡手動','売掛金11番街・EC手動'],133:['🟡手動','保留金ペイオニア1'],134:['🟡手動','保留金ペイオニア2'],135:['🟡手動','保留金スマスト'],136:['🟡手動','保留金11番街'],137:['🟡手動','還付金・手動'],138:['🟡手動','不動産自宅・評価/低頻度'],139:['🟡手動','不動産K北千住・低頻度'],140:['🟡手動','譲渡マリパパ・低頻度'],141:['🟢自動','MF連携(麻梨奈SBI銀行)'],142:['🟢自動','MF連携(麻梨奈SBI証券)'],143:['🟡手動','麻梨奈楽天NISA・MF未連携'],144:['🟡手動','麻梨奈 金・低頻度'],146:['—','▼負債/借入セクション見出し'],
    147:['🟢連動','城北←⑤借入(証票)自動連動'],148:['🟢連動','朝日←⑤借入 自動連動'],149:['🟢連動','大東京←⑤借入 自動連動'],150:['🟡手動','福井(役員借入)・⑤に無し'],151:['🟡手動','法人カードJAL VIEW・MF利用額/手動'],152:['🟡手動','奨学金(負債)・手動'],153:['🟡手動','カードMB/AGP'],154:['🟡手動','カードR1楽天'],155:['🟡手動','カードR2楽天'],156:['🟡手動','PayPayカード(解約検討)'],157:['🟡手動','AA1(解約)'],158:['🟡手動','AA2(解約)'],159:['🟡手動','AGP2'],160:['🟡手動','三井住友(解約)'],161:['🟡手動','三井住友(解約)'],162:['🟡手動','ソラチカ(解約)'],163:['🟡手動','買掛金OSR'],164:['🟢連動','住宅MCJ←⑤借入 自動連動'],165:['🟢連動','TBセゾン←⑤借入 自動連動'],166:['🟢連動','公庫←⑤借入 自動連動'],167:['🟢連動','浦安(TB創業)←⑤借入 自動連動']};
  bs.getRange(2,BL).setValue('連携'); bs.getRange(2,BM).setValue('自動/手動の理由'); bs.getRange(2,BL,1,2).setFontWeight('bold').setBackground('#DDDDDD');
  var labels=bs.getRange(1,2,lr,3).getValues(), cnt=0;
  for(var r=50;r<=lr;r++){ var hl=(labels[r-1][0]!==''&&labels[r-1][0]!==null)||(labels[r-1][2]!==''&&labels[r-1][2]!==null); if(!hl)continue;
    var v=M[r]||['🟡手動','要確認']; var bg=(v[0].indexOf('自動')>=0||v[0].indexOf('連動')>=0)?'#CDE9D6':(v[0]==='—'?'#EEEEEE':'#FFF1C2');
    bs.getRange(r,BL).setValue(v[0]).setBackground(bg); bs.getRange(r,BM).setValue(v[1]).setFontColor('#555555'); cnt++; }
  bs.setColumnWidth(BL,64); bs.setColumnWidth(BM,260);
  Logger.log('連携区分タグ '+cnt+'行'); _toast('連携区分タグ '+cnt+'行','連携タグ');
}

/* ===== EC粗利アンカー（EC側ダッシュボードに直近3ヶ月平均） ===== */
function buildEcAnchor(){
  var EC='1QjyPPOto7J1HiqA_Zb9-UIOe_FQZyqAGSn321R37Tzo', ec=SpreadsheetApp.openById(EC);
  var m=ec.getSheetByName('月次推移'); if(!m) ec.getSheets().forEach(function(s){if(!m&&(s.getName().indexOf('月次')>=0||s.getName().indexOf('推移')>=0))m=s;});
  if(!m) throw new Error('月次推移タブ無し'); var mn=m.getName();
  var scan=m.getRange(1,1,Math.min(12,m.getLastRow()),Math.min(15,m.getLastColumn())).getValues(), hdr=null;
  for(var r=0;r<scan.length;r++){var j=scan[r].join('|'); if(j.indexOf('年月')>=0&&j.indexOf('粗利')>=0){hdr=scan[r];break;}}
  if(!hdr) throw new Error('月次推移ヘッダ無し');
  function cl(k){for(var i=0;i<hdr.length;i++){if((''+hdr[i]).indexOf(k)>=0)return String.fromCharCode(65+i);}return null;}
  var cYM=cl('年月'),cAcc=cl('アカ'),cG=cl('粗利'),cSales=cl('売上'),cCost=cl('原価'); if(!cYM||!cAcc||!cG)throw new Error('列特定失敗');
  function q3(selCol,acc){ return "=IFERROR(AVERAGE(QUERY('"+mn+"'!A:Z,\"select "+selCol+" where "+cAcc+"='"+acc+"' order by "+cYM+" desc limit 3\",0)),0)"; }
  var a=ec.getSheetByName('④連動アンカー')||ec.insertSheet('④連動アンカー'); a.clear();
  a.getRange('A1').setValue('■ ④資金繰り/⑦損益 連動アンカー（直近3ヶ月平均）').setFontWeight('bold').setBackground('#FCE8B2');
  a.getRange('A2').setValue('クーパン1 直近3M平均粗利'); a.getRange('B2').setFormula(q3(cG,'クーパン1')).setNumberFormat('#,##0"円"').setFontWeight('bold');
  a.getRange('A3').setValue('クーパン2 直近3M平均粗利'); a.getRange('B3').setFormula(q3(cG,'クーパン2')).setNumberFormat('#,##0"円"').setFontWeight('bold');
  a.getRange('A4').setValue('合計粗利＝④へ流す値'); a.getRange('B4').setFormula('=B2+B3').setNumberFormat('#,##0"円"').setFontWeight('bold').setBackground('#CDE9D6');
  // ⑦損益(個人事業EC)用：合計の純売上・原価（直近3M平均）
  a.getRange('A6').setValue('合計 直近3M平均 純売上（→⑦損益r10）');
  a.getRange('B6').setFormula(cSales?q3(cSales,'合計'):'=0').setNumberFormat('#,##0"円"');
  a.getRange('A7').setValue('合計 直近3M平均 原価（→⑦損益r11）');
  a.getRange('B7').setFormula(cCost?q3(cCost,'合計'):'=0').setNumberFormat('#,##0"円"');
  a.getRange('A8').setValue('合計 直近3M平均 粗利（=売上-原価・検算）'); a.getRange('B8').setFormula('=B6-B7').setNumberFormat('#,##0"円"');
  SpreadsheetApp.flush(); Logger.log('EC連動アンカー更新 粗利合計='+a.getRange('B4').getValue()+' 売上='+a.getRange('B6').getValue()+' 原価='+a.getRange('B7').getValue()); _toast('EC粗利アンカー更新(売上/原価も)','EC連動');
}

/* ===== PLリアル化：⑦損益 個人事業(EC) 売上r10/原価r11 を EC粗利DBアンカーへ連動（一度だけ） ===== */
function linkPLtoEC(){
  var pl=_find('損益')||_find('管理会計')||_find('PL'); if(!pl)throw new Error('損益タブ無し');
  var EC='1QjyPPOto7J1HiqA_Zb9-UIOe_FQZyqAGSn321R37Tzo';
  // 行特定：A列に「売上」「原価」を含む最初の個人事業ブロック行（r10/r11想定だが保険で探索）
  var av=pl.getRange(1,1,Math.min(20,pl.getLastRow()),1).getValues();
  var rSales=-1,rCost=-1;
  for(var i=0;i<av.length;i++){ var s=''+av[i][0]; if(rSales<0&&s.indexOf('売上')>=0&&s.indexOf('EC')>=0)rSales=i+1; if(rCost<0&&rSales>0&&i+1>rSales&&s.indexOf('原価')>=0)rCost=i+1; }
  if(rSales<0)rSales=10; if(rCost<0)rCost=11; // 既定（buildKanriPL構造）
  for(var c=2;c<=9;c++){ // B..I（6月..1月）
    if(c===2){
      pl.getRange(rSales,c).setFormula('=IMPORTRANGE("'+EC+'","④連動アンカー!B6")').setNumberFormat('#,##0"円"');
      pl.getRange(rCost,c).setFormula('=IMPORTRANGE("'+EC+'","④連動アンカー!B7")').setNumberFormat('#,##0"円"');
    } else {
      pl.getRange(rSales,c).setFormula('=$B'+rSales).setNumberFormat('#,##0"円"');
      pl.getRange(rCost,c).setFormula('=$B'+rCost).setNumberFormat('#,##0"円"');
    }
  }
  Logger.log('PLリアル化: r'+rSales+'(売上)/r'+rCost+'(原価)をEC粗利DBアンカーB6/B7へ連動。粗利r12=自動。');
  _toast('PL個人事業EC 売上/原価をEC実績に連動','PLリアル化');
}

/* ===== 役員報酬5万/月（定期同額・未払い）をPLに反映＝法人費用に算入（CF流出なし・BSは未払累積） ===== */
function setYakuinHoshu(){
  var pl=_find('損益')||_find('管理会計')||_find('PL'); if(!pl)throw new Error('損益タブ無し');
  // 法人固定費(r7)に役員報酬5万を含める（税理士+外注+社保231,167 + 役員報酬50,000 = 281,167）全月B..I
  var BASE=231167, YAKU=50000, total=BASE+YAKU;
  for(var c=2;c<=9;c++){ pl.getRange(7,c).setValue(total).setNumberFormat('#,##0"円"'); }
  pl.getRange('A7').setValue('  固定費計（税理士+外注+社保231,167＋役員報酬50,000※未払＝281,167）');
  // 3タブ扱いメモ（下部・行ズレなし）
  pl.getRange('A33').setValue('◆ 役員報酬5万/月（定期同額・未払い）の3タブ扱い').setFontWeight('bold').setBackground('#FBF3D6');
  pl.getRange('A33:I33').merge();
  pl.getRange('A34').setValue('  ①PL(ここ)＝法人費用に+5万→法人経常−5万（損金算入） ②④資金繰り＝未払いのため現金流出ゼロ（計上しない） ③BS＝未払役員報酬が法人負債に月5万累積→決算で役員借入120万と相殺').setWrap(true);
  pl.getRange('A34:I34').merge();
  Logger.log('役員報酬5万を⑦損益 法人固定費へ算入（'+BASE+'→'+total+'）＋3タブ扱いメモ。CF=0/BS=未払累積。');
  _toast('役員報酬5万をPLに反映（未払・損金）','役員報酬');
}

/* ===== 損益PL 再整形（事業損益分岐／家族が潰れないライン分離） ===== */
function redesignPL(){
  var pl=_find('損益')||_find('管理会計')||_find('PL'); var cf=_find('資金繰り'); if(!pl)throw new Error('損益タブ無し');
  var cfName=cf?cf.getName():'④ 資金繰り';
  pl.getRange('A24').setValue('  ★事業損益分岐（事業固定費＋目標利益）＝粗利がこれ超で事業は目標達成');
  pl.getRange('A25').setValue('  事業 黒字判定（限界利益 vs 事業損益分岐）');
  pl.getRange('A27').setValue('■ 家族が潰れないライン（生活費・借入込み＝資金繰り連動）').setFontWeight('bold').setBackground('#CDE9D6');
  pl.getRange('A28').setValue('  必要粗利/月（純月次燃焼）'); pl.getRange('B28').setFormula("='"+cfName+"'!B31").setNumberFormat('#,##0"円"').setFontWeight('bold');
  pl.getRange('A29').setValue('  ※事業損益分岐(上)＝事業が黒字／家族が潰れないライン(これ)＝生活費・借入も賄うのに毎月要る粗利。別物。①司令塔のランウェイで確認。').setWrap(true);
  Logger.log('PL再整形 OK'); _toast('損益PL再整形','PL');
}

/* ===== BS負債(147-167)を⑤借入(証票)へ月次自動連動＝二重管理の解消 =====
 * 各借入のローン残高(月末)列を、BS月次列(行2が1日の月)にINDEX/MATCHで当月行自動選択。
 * D列(返済月)=テキスト"yyyy/MM"。福井(役員借入)は⑤に無いので対象外。 */
function linkBSdebtToBorrow(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), bs=_bs(), T="'⑤ 借入'";
  var MAP={147:'I',148:'P',149:'W',164:'AR',165:'AK',166:'AY',167:'AD'}; // BS負債行→⑤残高(月末)列
  // ⑤借入が持つ月(D列)の集合＝範囲外の月は触らない（空白化を防ぐ）
  var b5=ss.getSheetByName('⑤ 借入'); if(!b5) throw new Error('⑤借入無し');
  var dv=b5.getRange(4,4,Math.max(1,b5.getLastRow()-3),1).getValues();
  var dset={}; dv.forEach(function(x){ var s=''+x[0]; if(s)dset[s]=1; });
  var lc=bs.getLastColumn(), hdr=bs.getRange(2,1,1,lc).getValues()[0];
  var cols=[]; for(var c=0;c<lc;c++){ var v=hdr[c]; if(v instanceof Date && v.getDate()===1){ var k=Utilities.formatDate(v,'JST','yyyy/MM'); if(dset[k]) cols.push(c+1); } } // ⑤にある月のみ
  if(!cols.length) throw new Error('⑤と一致するBS月次列が無い');
  var n=0;
  Object.keys(MAP).forEach(function(r){
    var col5=MAP[r];
    cols.forEach(function(c){ var L=_colL(c);
      bs.getRange(Number(r),c).setFormula('=IFERROR(INDEX('+T+'!$'+col5+':$'+col5+',MATCH(TEXT('+L+'$2,"yyyy/MM"),'+T+'!$D:$D,0)),"")'); n++; });
  });
  Logger.log('BS負債→⑤借入連動(⑤にある月のみ): 城北I/朝日P/大東京W/住宅AR/TBセゾンAK/公庫AY/浦安AD ×'+cols.length+'列='+n+'セル。範囲外月は不触。福井除外。');
  _toast('BS負債を⑤借入へ連動：'+n+'セル(範囲内月のみ)','借入↔BS連動');
}

/* ===== 検算：⑦損益(役員報酬/EC実績反映後)＋④資金繰り(ランウェイ)を当月B列でダンプ ===== */
function verifyPL(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var pl=_find('損益')||_find('管理会計')||_find('PL'); var cf=_find('資金繰り');
  var log=[];
  if(pl){ log.push('=== ⑦'+pl.getName()+' 当月(B列) ===');
    var pv=pl.getRange(1,1,Math.min(31,pl.getLastRow()),2).getValues();
    for(var i=0;i<pv.length;i++){ var a=(''+pv[i][0]).trim(), b=pv[i][1]; if(a!==''&&b!==''&&b!==null) log.push('r'+(i+1)+' '+a+' ＝ '+b); } }
  if(cf){ log.push('=== ④'+cf.getName()+' 当月(B列) ===');
    var fv=cf.getRange(1,1,Math.min(34,cf.getLastRow()),2).getValues();
    for(var j=0;j<fv.length;j++){ var a2=(''+fv[j][0]).trim(), b2=fv[j][1]; if(a2!==''&&b2!==''&&b2!==null) log.push('r'+(j+1)+' '+a2+' ＝ '+b2); } }
  Logger.log(log.join('\n'));
}

/* ===== ②本部マトリクス v2（実態時間版）=====
 * 目的＝「来週どこに時間を寄せるか」を3秒で決める。(1)実態正確 (2)円/hで序列 (3)内務を圧縮対象で可視。
 * 実績h=カレンダー自動＋補正h(隙間・手入力/Claude週次推定)。内務5分解。AI査定等は仕込みへ。分類定義は注釈で固定。 */
var MX_BURN=660947; // 純月次燃焼(検算値)。家族◯ヶ月分の分母
// [本部,活動,分類,見込粗利,確度,将来価値メモ,判断,メモ,catKey(classifyV2の返り値)]
var MX_BIZ=[
 ['05 物件','不動産 売却/買取(栄町6/20)','営業',1900000,0.9,'','続ける(6月集中)','6/20決済・一回性。継続源(医療)と両睨み','不動産'],
 ['04 コンサル','医療テナント/承継','営業',660000,0.5,'','増やす','1件66万・継続性◎の本命現金源','医療'],
 ['03/05 協働','買取再販テレアポ(宮崎)','営業',100000,0.3,'','様子見','オーロラ次第・確度低の補助線','テレアポ'],
 ['03 運営','EC 韓国輸出(クーパン)','維持',150000,0.9,'','維持(黒字回転)','月次継続・粗利は実績連動','EC'],
 ['04 士業','土地家屋調査士','仕込',0,0,'将来:2027合格で単価UP','続ける','今期見込0・将来ROIで評価','調査士'],
 ['本命','メディア×AI(YouTube/HP/MyAI)','仕込',0,0,'将来:継続収益の本命','増やす','今は仕込み・将来ROIで評価','メディア'],
 ['05 仕込','AI査定エンジン(物件)','仕込',0,0,'将来:仕入の自動化','続ける','★隙間8h判明・物件営業の仕込みへ再分類','AI査定'],
 ['01 経営','内務-経営(損切り/意思決定)','内務',0,0,'','圧縮','','内務経営'],
 ['02 資金','内務-資金(記帳/帳簿)','内務',0,0,'','圧縮','銀行面談は営業へ別計上','内務資金'],
 ['03 運営','内務-運営(台帳/雑務)','内務',0,0,'','圧縮','','内務運営'],
 ['横断','内務-ツール開発(財務ダッシュ等)','内務',0,0,'','圧縮(沼注意)','営業仕込みでないツールはコスト','内務ツール'],
 ['横断','内務-朝礼終礼/朝ブリーフ','内務',0,0,'','圧縮','','朝礼終礼'],
 ['-','その他(移動/雑)','内務',0,0,'','減らす','','その他'],
 ['00 家族','親子/夫婦(目的・死守)','目的',null,null,'','死守','稼ぐ目的そのもの・ROI対象外','家族']
];
function _freshMX(name){ var ss=SpreadsheetApp.getActiveSpreadsheet(); var sh=ss.getSheetByName(name); if(sh)sh.clear(); else sh=ss.insertSheet(name); sh.setHiddenGridlines(true); return sh; }
function buildMatrixV2(){
  var YEN='#,##0"円"',HRS='0.0"h"',ROI='#,##0"円/h"',PCT='0.0%',MON='0.0"ヶ月"';
  var BRICK='#AA2E26',SUBF='#F0E2DF',FAM='#CDE9D6',HI='#DDF3DD',GREENF='#CDE9D6',GOLDF='#FBF3D6',NF='#FCEFE7';
  var sh=_freshMX('②本部マトリクス');
  sh.getRange('A1:O1').merge().setValue('② 本部マトリクス v2（実態時間×円/h×家族）— 来週どこに時間を寄せるかを3秒で決める表')
    .setFontColor('#FFF').setBackground(BRICK).setFontWeight('bold');
  sh.getRange('A2:O2').merge().setValue('実績h=カレンダー自動／補正h=隙間時間(早朝夜移動)を手入力or週次でClaude推定／合計h=実績+補正。円/h(期待)=期待粗利÷合計h・序列用。円/h(将来)=仕込みの将来ROI(定性)。家族◯ヶ月分=金÷世帯燃焼'+Math.round(MX_BURN/10000)+'万。')
    .setBackground(FAM).setWrap(true);
  var H=['本部','活動','分類','実績h(カレンダー)','補正h(隙間)','合計h','構成比','見込粗利','確度','期待粗利','円/h(期待)','円/h(将来価値)','家族◯ヶ月分','判断','メモ'];
  sh.getRange(4,1,1,H.length).setValues([H]).setFontWeight('bold').setBackground(SUBF).setHorizontalAlignment('center').setWrap(true);
  // 分類定義を注釈で固定（C4ヘッダ）
  sh.getRange(4,3).setNote('分類の定義：\n・営業＝相手を動かす行為(商談/追客/面談/送信)\n・仕込＝将来の継続収益の種(メディア/調査士/AI査定)\n・維持＝黒字で回す既存(EC)\n・内務＝それ以外(台帳/記帳/ツール/朝礼)＝圧縮対象\n・目的＝家族(稼ぐ目的・ROI対象外)');
  sh.getRange(4,5).setNote('補正h＝カレンダーに無い隙間時間(早朝/夜/移動)。手入力、または週次でClaudeがkhd-log・会話ログから推定して埋める。「AI査定8hが実績に出ない」漏れをここで塞ぐ。');
  sh.getRange(4,12).setNote('円/h(将来価値)＝調査士・メディア等の"将来の継続収益"を時間ROIで別評価する欄(定性or概算)。期待粗利0でも将来価値で残す判断材料。');
  var first=5,n=MX_BIZ.length,last=first+n-1,tot=last+2;
  for(var i=0;i<n;i++){
    var r=first+i,b=MX_BIZ[i],rev=(typeof b[3]==='number'&&b[3]>0);
    sh.getRange(r,1).setValue(b[0]); sh.getRange(r,2).setValue(b[1]);
    sh.getRange(r,3).setValue(b[2]).setHorizontalAlignment('center');
    sh.getRange(r,4).setValue(0).setNumberFormat(HRS);                                   // D 実績h(カレンダー)
    sh.getRange(r,5).setValue(0).setNumberFormat(HRS).setBackground('#FFF7DB');          // E 補正h(手入力)
    sh.getRange(r,6).setFormula('=D'+r+'+E'+r).setNumberFormat(HRS);                     // F 合計h
    sh.getRange(r,7).setFormula('=IF($F$'+tot+'=0,0,F'+r+'/$F$'+tot+')').setNumberFormat(PCT); // G 構成比
    if(rev){
      sh.getRange(r,8).setValue(b[3]).setNumberFormat(YEN);
      sh.getRange(r,9).setValue(b[4]).setNumberFormat(PCT).setHorizontalAlignment('center');
      sh.getRange(r,10).setFormula('=H'+r+'*I'+r).setNumberFormat(YEN);                  // J 期待粗利
      sh.getRange(r,11).setFormula('=IF(F'+r+'=0,"-",J'+r+'/F'+r+')').setNumberFormat(ROI); // K 円/h期待
      sh.getRange(r,13).setFormula('=J'+r+'/'+MX_BURN).setNumberFormat(MON);             // M 家族ヶ月
    } else {
      sh.getRange(r,8,1,3).setValues([['-','-','-']]).setHorizontalAlignment('center');
      sh.getRange(r,11).setValue('-').setHorizontalAlignment('center');
      sh.getRange(r,13).setValue('-').setHorizontalAlignment('center');
    }
    sh.getRange(r,12).setValue(b[5]||'').setWrap(true);     // L 円/h将来(定性メモ)
    sh.getRange(r,14).setValue(b[6]).setHorizontalAlignment('center'); // N 判断
    sh.getRange(r,15).setValue(b[7]).setWrap(true);          // O メモ
    // 分類で色分け
    var cls=b[2], bg=(cls==='営業')?HI:(cls==='仕込')?GOLDF:(cls==='目的')?FAM:(cls==='維持')?NF:null;
    if(bg) sh.getRange(r,3).setBackground(bg);
    if(cls==='目的') sh.getRange(r,1,1,15).setBackground(FAM);
  }
  // 合計行
  sh.getRange(tot,2).setValue('合計（総時間）').setFontWeight('bold').setBackground(SUBF);
  sh.getRange(tot,4).setFormula('=SUM(D'+first+':D'+last+')').setNumberFormat(HRS).setFontWeight('bold').setBackground(SUBF);
  sh.getRange(tot,5).setFormula('=SUM(E'+first+':E'+last+')').setNumberFormat(HRS).setFontWeight('bold').setBackground(SUBF);
  sh.getRange(tot,6).setFormula('=SUM(F'+first+':F'+last+')').setNumberFormat(HRS).setFontWeight('bold').setBackground(SUBF);
  sh.getRange(tot,10).setFormula('=SUM(J'+first+':J'+last+')').setNumberFormat(YEN).setFontWeight('bold').setBackground(SUBF);
  sh.getRange(tot,13).setFormula('=SUM(M'+first+':M'+last+')').setNumberFormat(MON).setFontWeight('bold').setBackground(SUBF);
  // 営業直結% と 内務%（圧縮対象の可視化）
  var e1=tot+1;
  sh.getRange(e1,2).setValue('★営業直結比率（分類=営業）／目標60%').setFontWeight('bold');
  sh.getRange(e1,6).setFormula('=IF(F'+tot+'=0,"-",SUMIF(C'+first+':C'+last+',"営業",F'+first+':F'+last+')/F'+tot+')').setNumberFormat(PCT).setFontWeight('bold').setBackground(HI);
  var e2=tot+2;
  sh.getRange(e2,2).setValue('▲内務比率（分類=内務）／圧縮対象').setFontWeight('bold');
  sh.getRange(e2,6).setFormula('=IF(F'+tot+'=0,"-",SUMIF(C'+first+':C'+last+',"内務",F'+first+':F'+last+')/F'+tot+')').setNumberFormat(PCT).setFontWeight('bold').setBackground(GOLDF);
  var e3=tot+3;
  sh.getRange(e3,2).setValue('◇仕込み比率（分類=仕込）／将来の種').setFontWeight('bold');
  sh.getRange(e3,6).setFormula('=IF(F'+tot+'=0,"-",SUMIF(C'+first+':C'+last+',"仕込",F'+first+':F'+last+')/F'+tot+')').setNumberFormat(PCT).setFontWeight('bold').setBackground(GOLDF);
  sh.getRange(e3+2,1,1,15).merge().setValue('【週次(月)】実績hをカレンダー自動更新→隙間は補正h入力→営業%(目標60)・内務%(圧縮)・仕込%を点検し来週の配分1つを決める。【月次】粗利確定で見込みを更新。AI査定等ツールは「営業の仕込みか/ただの内務か」を毎回問う。').setWrap(true);
  sh.setColumnWidth(1,80); sh.setColumnWidth(2,210); sh.setColumnWidth(3,52); sh.setColumnWidth(12,150); sh.setColumnWidth(15,240);
  for(var c=4;c<=11;c++) sh.setColumnWidth(c,76); sh.setColumnWidth(13,82); sh.setColumnWidth(14,80);
  sh.setFrozenRows(4);
  Logger.log('②本部マトリクスv2 構築：'+n+'活動(内務5分解・AI査定→仕込み)。実績hはupdateActualHoursV2で更新。');
  SpreadsheetApp.getActiveSpreadsheet().toast('②マトリクスv2構築。次に「実績hをカレンダーから更新」','②v2',6);
}
function classifyV2(t){
  t=(t||'').trim(); var m=t.match(/^\s*(\d{2})[_\.]/); var hb=m?m[1]:null;
  var addr=/丁目|番地|県|市|区|町|字/.test(t);
  // 仕込み・AI査定を最優先で拾う
  if(/査定|reinfolib|スクリーニング|物件intake|物件収集|AI査定|スクレイ/.test(t))return'AI査定';
  if(/そうけん|My ?AI|メディア|YouTube|HP|デッキ|サロン|発信|X投稿|note/.test(t))return'メディア';
  if(/調査士|土地家屋|自己投資|過去問|書式/.test(t))return'調査士';
  // 朝礼終礼/ツール開発(内務分解)
  if(/朝礼|終礼|朝ブリ|日次|週次報告|振り返り/.test(t))return'朝礼終礼';
  if(/ダッシュ|BS|資金繰り作成|DB構築|GAS|スプシ|スクリプト|パイプライン構築|KPI整備|ツール|自動化/.test(t))return'内務ツール';
  // 本部prefix
  if(hb==='00'){ if(/親子|家族|葵斗|モーニング|夫婦/.test(t))return'家族'; if(addr)return'不動産'; return'家族'; }
  if(hb==='01')return'内務経営';
  if(hb==='02'){ if(/面談|融資打診|銀行訪問|公庫/.test(t))return'テレアポ'; return'内務資金'; } // 銀行営業は営業扱い(テレアポ枠流用)
  if(hb==='03'){ if(/韓国輸出|クーパン|EC|せどり/.test(t))return'EC'; if(/バイセル|物上げ|仕入|追客/.test(t))return'不動産'; return'内務運営'; }
  if(hb==='04'){ if(/オーロラ|テレアポ|インディード|鍼灸|採用/.test(t))return'テレアポ';
                 if(/TAW|歯科|医療|セミナー|福井|クリニック|診療/.test(t))return'医療'; return'医療'; }
  if(hb==='05')return'不動産';
  // prefix無し キーワード
  if(/親子|家族|葵斗/.test(t))return'家族';
  if(/韓国輸出|クーパン/.test(t))return'EC';
  if(/オーロラ|テレアポ|インディード|石原|採用/.test(t))return'テレアポ';
  if(/TAW|歯科|医療|セミナー|福井|クリニック|診療|野口/.test(t))return'医療';
  if(/バイセル|物上げ|仕入|決済|追客/.test(t)||addr)return'不動産';
  if(/記帳|帳簿|税理士|還付/.test(t))return'内務資金';
  if(/台帳|報告|整理|資料/.test(t))return'内務運営';
  return'その他';
}
function updateActualHoursV2(){
  var DAYS=7, end=new Date(), start=new Date(end.getTime()-DAYS*24*3600*1000);
  var evs=CalendarApp.getDefaultCalendar().getEvents(start,end), agg={};
  for(var i=0;i<evs.length;i++){ var e=evs[i]; if(e.isAllDayEvent())continue;
    var h=(e.getEndTime()-e.getStartTime())/3600000; var b=classifyV2(e.getTitle()); agg[b]=(agg[b]||0)+h; }
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName('②本部マトリクス');
  if(!sh){ buildMatrixV2(); sh=ss.getSheetByName('②本部マトリクス'); }
  var first=5;
  for(var j=0;j<MX_BIZ.length;j++){ var key=MX_BIZ[j][8]; sh.getRange(first+j,4).setValue(Math.round((agg[key]||0)*10)/10); }
  var tz=Session.getScriptTimeZone(), f=function(d){return Utilities.formatDate(d,tz,'M/d');};
  sh.getRange(2,1,1,15).merge().setValue('実績h=カレンダー自動【窓 '+f(start)+'〜'+f(end)+'・直近'+DAYS+'日】＋補正h(隙間)。円/h(期待)で序列／円/h(将来)で仕込み評価／家族◯ヶ月分=金÷燃焼'+Math.round(MX_BURN/10000)+'万。隙間時間はE列に手入力 or 週次でClaude推定。').setBackground('#CDE9D6').setWrap(true);
  Logger.log('②v2 実績h更新（直近'+DAYS+'日・分類14種）。補正h(E列)は別途。');
  ss.toast('②v2 実績h更新。隙間時間はE列に補正入力を','②v2',6);
}

/* ===== 確認：借入/タブ構造 ===== */
function inspectBorrow(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(), log=['=== 全タブ ==='];
  ss.getSheets().forEach(function(s){log.push(' ・'+s.getName()+' (r'+s.getLastRow()+' c'+s.getLastColumn()+')');});
  var b=_find('借入','予定');
  if(b){ log.push('=== '+b.getName()+' 先頭18行 ==='); var lc=Math.min(b.getLastColumn(),30); var v=b.getRange(1,1,Math.min(18,b.getLastRow()),lc).getValues();
    for(var r=0;r<v.length;r++){var cs=[];for(var c=0;c<lc;c++){var x=v[r][c];if(x!==''&&x!==null){var sv=(x instanceof Date)?('D'+(x.getMonth()+1)):(''+x);cs.push(_colL(c+1)+(r+1)+'='+(''+sv).slice(0,14));}}if(cs.length)log.push(cs.join(' '));} }
  var z=_find('返済予定');
  if(z){ log.push('=== '+z.getName()+' 借入見出し ==='); var zl=z.getLastRow(),zc=Math.min(z.getLastColumn(),12),zv=z.getRange(1,1,zl,zc).getValues();
    for(var r2=0;r2<zl;r2++){var row=zv[r2],j=row.join('|'); if(/城北|朝日|大東京|公庫|セゾン|浦安|創業|住宅|MCJ|日本住宅/.test(j)&&/借入|返済|予定|銀行|信金|信組|金庫|残高|当初/.test(j)){var cs=[];for(var c2=0;c2<zc;c2++){if(row[c2]!==''&&row[c2]!==null){var s2=(row[c2] instanceof Date)?('D'+row[c2].getFullYear()+'/'+(row[c2].getMonth()+1)):(''+row[c2]);cs.push(_colL(c2+1)+(r2+1)+'='+(''+s2).slice(0,16));}}log.push(cs.join(' '));}} }
  Logger.log(log.join('\n'));
}
