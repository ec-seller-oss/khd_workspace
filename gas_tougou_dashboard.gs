/**
 * KHD 統合ダッシュボード（Google Sheets ネイティブ版）2026-06-02
 * 稼働SSoT v2 (1ofLJOFuW…) に貼り付けて使う Apps Script。
 * Excel(xlsx)をやめ、Sheets+GASに一本化。
 *   ①統合司令塔  : 過去→現在→未来＋行動の4軸。財務は既存「BS」タブを参照。
 *   ②本部マトリクス: 時間×金×家族。実績hはCalendarApp自動集計。妻向けに「家族◯ヶ月分」翻訳。
 *   ⑥日次ループ   : 夜=抽出→タスク化／朝=ブリーフ→再配置 の運用手順。
 * 使い方: 拡張機能▸Apps Script に貼付 → buildAll() を実行(初回はCalendar/Sheets認可) → 完了。
 *        以後はシート上部メニュー「📊統合ダッシュボード」から再実行・実績h更新が可能。
 * ※③④⑤(財務)は作らない＝既存のBS/資金繰りタブを使う(重複排除)。
 */

// ============ CONFIG（ここだけ毎月/環境で調整）============
var CFG = {
  BS_SHEET: 'BS',     // 純資産BS本体タブ名（gid985549912をBSにリネーム済）
  CUR_COL : 'BI',     // 当月列。6/1=BI。月が変わったら翌列(BJ…)に1箇所だけ更新。
  ASSET_ROWS: '50:144',   // 資産明細の行範囲（総資産=SUM）
  LIAB_ROWS : '147:167',  // 負債明細の行範囲（総負債=SUM）
  CASH_ROWS : [12,13,14], // 現預金(流動性ハイ)の要約行：法人/研太/麻梨奈
  BURN: 631342,           // 世帯 純月次燃焼(memory)。ランウェイ・家族◯ヶ月分の分母。
  CAL_DAYS: 7             // 実績h集計の対象日数（直近N日）
};

// 配色（スライド標準）
var C = {HF:'#AA2E26', WHITE:'#FFFFFF', SUBF:'#F0E2DF', FAM:'#CDE9D6', HI:'#DDF3DD',
         P:'#EAE0DA', N:'#FCEFE7', F:'#E7EEF6', A:'#FBF3D6'};
var YEN='#,##0"円"', HRS='0.0"h"', ROI='#,##0"円/h"', PCT='0.0%', MON='0.0"ヶ月"';

function onOpen(){
  SpreadsheetApp.getUi().createMenu('📊統合ダッシュボード')
    .addItem('全構築（①②⑥）', 'buildAll')
    .addItem('実績hをカレンダーから更新', 'updateActualHours')
    .addItem('②本部マトリクスだけ再構築', 'buildMatrix')
    .addToUi();
}

function buildAll(){
  buildMatrix();
  buildDailyLoop();
  buildCockpit();
  updateActualHours();
  SpreadsheetApp.getActiveSpreadsheet().toast('①②⑥を構築し、実績hを更新しました', '統合ダッシュボード', 6);
}

// ---- ユーティリティ ----
function freshSheet(name){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (sh) sh.clear(); else sh = ss.insertSheet(name);
  sh.setHiddenGridlines(true);
  return sh;
}
function head(sh, range, text){
  sh.getRange(range).merge().setValue(text).setFontColor(C.WHITE).setBackground(C.HF).setFontWeight('bold');
}

// ============ ② 本部マトリクス（時間×金×家族）============
function buildMatrix(){
  var sh = freshSheet('②本部マトリクス');
  head(sh,'A1:M1','② 本部マトリクス（時間×金×家族）— パパの1時間がいくらを生み、家族の暮らし何ヶ月分になるか');
  sh.getRange('A2:M2').merge().setValue(
    '👨‍👩‍👦 妻と見る：高い活動に時間を寄せる＝将来の家族時間を“買う”こと。'+
    '実績h=カレンダー自動／見込み粗利×確度＝期待粗利／円per h＝1時間の価値／家族◯ヶ月分＝その金で暮らせる月数(世帯燃焼'+
    Math.round(CFG.BURN/10000)+'万)。').setBackground(C.FAM).setWrap(true);
  var H=['本部','活動','実績h(月)','構成比','見込み粗利','確度','期待粗利','円/h(期待)','家族◯ヶ月分','営業直結','判断','投下予定h','メモ'];
  sh.getRange(4,1,1,H.length).setValues([H]).setFontWeight('bold').setBackground(C.SUBF)
    .setHorizontalAlignment('center').setWrap(true);
  // [本部,活動,見込み,確度,投下予定h,営業直結,判断,メモ]
  var biz=[
   ['05 物件調達','不動産 売却/買取(栄町6/20決済)',1900000,0.9,40,'営業','続ける(6月集中)','6/20決済・確度90%。一回性なので継続源(医療)と両睨み'],
   ['04 コンサル','医療テナント/承継コンサル',660000,0.5,60,'営業','増やす','1件66万・継続性◎の本命現金源'],
   ['03 事業運営','EC 韓国輸出(クーパン)',150000,0.9,40,'','維持(黒字回転)','月次継続・粗利は各DBから月末転記'],
   ['04 調査士','土地家屋調査士(将来ROI)',0,0,80,'仕込','続ける','将来ROI・今期見込0。2027合格で単価UP'],
   ['本命','メディア×AI(YouTube/HP/MyAI)',0,0,50,'仕込','増やす','将来の継続収益源・今は仕込み'],
   ['03/05 協働','買取再販テレアポ・採用(宮崎)',100000,0.3,25,'営業','様子見','オーロラ次第・確度低の補助線'],
   ['01-03 内務','経営/資金/運営(朝礼終礼・台帳等)',0,0,60,'','減らす','★最大の塊=コスト。圧縮し営業へ振替'],
   ['-','その他',0,0,10,'','減らす','映画・雑など。意図的に減らす'],
   ['00 家族','親子/夫婦(目的・死守)',null,null,125,'目的','死守','稼ぐ目的そのもの。ROI対象外・増やすのがゴール']
  ];
  var first=5, n=biz.length, last=first+n-1, tot=last+2;  // 空行を1つ挟む
  for (var i=0;i<n;i++){
    var r=first+i, b=biz[i], rev=(typeof b[2]==='number' && b[2]>0);
    sh.getRange(r,1).setValue(b[0]);
    sh.getRange(r,2).setValue(b[1]);
    sh.getRange(r,3).setValue(0).setNumberFormat(HRS);                                  // C 実績h
    sh.getRange(r,4).setFormula('=IF($C$'+tot+'=0,0,C'+r+'/$C$'+tot+')').setNumberFormat(PCT); // D 構成比
    if (rev){
      sh.getRange(r,5).setValue(b[2]).setNumberFormat(YEN);
      sh.getRange(r,6).setValue(b[3]).setNumberFormat(PCT).setHorizontalAlignment('center');
      sh.getRange(r,7).setFormula('=E'+r+'*F'+r).setNumberFormat(YEN);
      sh.getRange(r,8).setFormula('=IF(L'+r+'=0,"-",G'+r+'/L'+r+')').setNumberFormat(ROI);
      sh.getRange(r,9).setFormula('=G'+r+'/'+CFG.BURN).setNumberFormat(MON);
    } else {
      sh.getRange(r,5,1,5).setValues([['-','-','-','-','-']]).setHorizontalAlignment('center');
    }
    sh.getRange(r,10).setValue(b[5]).setHorizontalAlignment('center');
    sh.getRange(r,11).setValue(b[6]).setHorizontalAlignment('center');
    sh.getRange(r,12).setValue(b[4]).setNumberFormat(HRS);
    sh.getRange(r,13).setValue(b[7]).setWrap(true);
    if (b[5]==='目的') sh.getRange(r,1,1,13).setBackground(C.FAM);
  }
  // 合計行
  sh.getRange(tot,2).setValue('合計（全活動＝総時間）').setFontWeight('bold').setBackground(C.SUBF);
  sh.getRange(tot,3).setFormula('=SUM(C'+first+':C'+last+')').setNumberFormat(HRS).setFontWeight('bold').setBackground(C.SUBF);
  sh.getRange(tot,4).setFormula('=IF(C'+tot+'=0,0,1)').setNumberFormat('0%').setFontWeight('bold').setBackground(C.SUBF);
  sh.getRange(tot,7).setFormula('=SUM(G'+first+':G'+last+')').setNumberFormat(YEN).setFontWeight('bold').setBackground(C.SUBF);
  sh.getRange(tot,9).setFormula('=SUM(I'+first+':I'+last+')').setNumberFormat(MON).setFontWeight('bold').setBackground(C.SUBF);
  sh.getRange(tot,12).setFormula('=SUM(L'+first+':L'+last+')').setNumberFormat(HRS).setFontWeight('bold').setBackground(C.SUBF);
  // 妻サマリー
  var sp=tot+1;
  sh.getRange(sp,2).setValue('👨‍👩‍👦 今月の期待粗利 合計 →').setFontWeight('bold').setBackground(C.FAM);
  sh.getRange(sp,7).setFormula('=G'+tot).setNumberFormat(YEN).setFontWeight('bold').setBackground(C.FAM);
  sh.getRange(sp,8).setValue('＝家族').setFontWeight('bold').setBackground(C.FAM).setHorizontalAlignment('right');
  sh.getRange(sp,9).setFormula('=I'+tot).setNumberFormat(MON).setFontWeight('bold').setBackground(C.FAM).setHorizontalAlignment('center');
  sh.getRange(sp,10,1,4).merge().setValue('分の暮らしを確保。だから今は時間を寄せる。').setBackground(C.FAM).setWrap(true);
  // 営業直結%
  var eig=sp+1;
  sh.getRange(eig,2).setValue('★営業直結比率(不動産+医療+テレアポ)／目標60%').setFontWeight('bold');
  sh.getRange(eig,3).setFormula('=IF(C'+tot+'=0,"-",SUMIF(J'+first+':J'+last+',"営業",C'+first+':C'+last+')/C'+tot+')')
    .setNumberFormat(PCT).setFontWeight('bold').setBackground(C.HI);
  // 注記
  var nt=eig+2;
  sh.getRange(nt,1,1,13).merge().setValue(
    '【日次】予定どおり時間を使えたか(カレンダー実測)。【週次】営業直結%とメディア時間を点検し配分調整。【月次】粗利確定→実績で見込みを更新し続ける/減らす/やめるを決定。').setWrap(true);
  // 幅
  sh.setColumnWidth(1,95); sh.setColumnWidth(2,230); sh.setColumnWidth(13,250);
  for (var c=3;c<=12;c++) sh.setColumnWidth(c,84);
  // メタ（他タブが参照する行番号）を記録
  PropertiesService.getDocumentProperties().setProperties({
    MX_TOT:String(tot), MX_EIG:String(eig), MX_LAST:String(last)});
  return {first:first,last:last,tot:tot,eig:eig,sp:sp};
}

// ============ 実績h をカレンダーから集計して ②C列へ ============
function classify(t){
  t=(t||'').trim();
  var m=t.match(/^\s*(\d{2})[_\.]/); var hb=m?m[1]:null;
  var addr=/丁目|番地|県|市|区|町|字/.test(t);
  if(hb==='00'){ if(/親子|家族|葵斗|モーニング|会議/.test(t))return'家族'; if(addr)return'不動産'; return'家族'; }
  if(hb==='01'||hb==='02')return'内務';
  if(hb==='03'){ if(/韓国輸出|クーパン|EC|せどり/.test(t))return'EC'; if(/バイセル|物上げ|仕入/.test(t))return'不動産'; return'内務'; }
  if(hb==='04'){ if(/調査士|土地家屋|マン菅|賃管|診断士/.test(t))return'調査士';
                 if(/オーロラ|テレアポ|インディード|鍼灸/.test(t))return'テレアポ';
                 if(/YouTube|メディア|HP/.test(t))return'メディア'; return'医療'; }
  if(hb==='05')return'不動産';
  if(/親子|家族|葵斗/.test(t))return'家族';
  if(/調査士|土地家屋|自己投資/.test(t))return'調査士';
  if(/韓国輸出|クーパン/.test(t))return'EC';
  if(/オーロラ|テレアポ|インディード|石原|鍼灸|採用/.test(t))return'テレアポ';
  if(/そうけん|My ?AI|メディア|YouTube|HP|デッキ/.test(t))return'メディア';
  if(/TAW|歯科|医療|セミナー|福井|クリニック|診療|野口|ソニー生命/.test(t))return'医療';
  if(/バイセル|物上げ|仕入|決済/.test(t)||addr)return'不動産';
  if(/台帳|BS|DB|ダッシュ|パイプライン|週次|KPI|報告/.test(t))return'内務';
  return'その他';
}
function updateActualHours(){
  var end=new Date(); var start=new Date(end.getTime()-CFG.CAL_DAYS*24*3600*1000);
  var evs=CalendarApp.getDefaultCalendar().getEvents(start,end);
  var agg={};
  for (var i=0;i<evs.length;i++){
    var e=evs[i]; if(e.isAllDayEvent())continue;
    var h=(e.getEndTime()-e.getStartTime())/3600000;
    var b=classify(e.getTitle()); agg[b]=(agg[b]||0)+h;
  }
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('②本部マトリクス');
  if(!sh){ buildMatrix(); sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('②本部マトリクス'); }
  // 行5-13の対応（buildMatrixのbiz順）
  var rowOf={5:'不動産',6:'医療',7:'EC',8:'調査士',9:'メディア',10:'テレアポ',11:'内務',12:'その他',13:'家族'};
  for (var r in rowOf){ sh.getRange(Number(r),3).setValue(Math.round((agg[rowOf[r]]||0)*10)/10); }
  // 窓を注記
  var tz=Session.getScriptTimeZone();
  var fmt=function(d){return Utilities.formatDate(d,tz,'M/d');};
  sh.getRange(2,1,1,13).merge().setValue(
    '👨‍👩‍👦 妻と見る：高い活動に時間を寄せる＝将来の家族時間を“買う”こと。実績h=カレンダー自動集計【窓 '
    +fmt(start)+'〜'+fmt(end)+'・直近'+CFG.CAL_DAYS+'日】／期待粗利=見込み×確度／円/h=1時間の価値／家族◯ヶ月分=金÷世帯燃焼'
    +Math.round(CFG.BURN/10000)+'万。').setBackground(C.FAM).setWrap(true);
}

// ============ ① 統合司令塔（4軸1枚・財務は既存BS参照）============
function bsRef(rows){ // 例 '50:144' → SUM(BS!BI50:BI144)
  var a=rows.split(':');
  return "SUM('"+CFG.BS_SHEET+"'!"+CFG.CUR_COL+a[0]+":"+CFG.CUR_COL+a[1]+")";
}
function cashRef(){
  var p=[]; for (var i=0;i<CFG.CASH_ROWS.length;i++) p.push("'"+CFG.BS_SHEET+"'!"+CFG.CUR_COL+CFG.CASH_ROWS[i]);
  return p.join('+');
}
function buildCockpit(){
  var sh=freshSheet('①統合司令塔');
  // ②の行番号を取得（buildMatrixが記録）
  var pr=PropertiesService.getDocumentProperties();
  var MT=pr.getProperty('MX_TOT')||'14', ME=pr.getProperty('MX_EIG')||'16';
  var MX="'②本部マトリクス'";
  head(sh,'A1:E1','① 統合司令塔 — 過去 → 現在 → 未来 ＋ 行動');
  sh.getRange('A2').setValue('基準日 '+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy/MM/dd')
    +'（②本部マトリクス／BSタブ '+CFG.BS_SHEET+'!'+CFG.CUR_COL+'列 と連動）');
  var block=function(r,t,bg){ sh.getRange(r,1,1,4).setBackground(bg); sh.getRange(r,1).setValue(t).setFontWeight('bold'); };
  var kv=function(r,label,formula,nf,note){
    sh.getRange(r,1).setValue(label).setFontWeight('bold').setBackground(C.SUBF);
    sh.getRange(r,2).setFormula(formula).setNumberFormat(nf||YEN).setFontWeight('bold');
    if(note) sh.getRange(r,3).setValue(note).setWrap(true);
  };
  // 過去
  block(4,'◆ 過去（実績ストック・BS連動）',C.P);
  kv(5,'純資産（自己資本）','='+bsRef(CFG.ASSET_ROWS)+'-'+bsRef(CFG.LIAB_ROWS),YEN,'世間中央値超だが流動性は薄い');
  kv(6,'総資産','='+bsRef(CFG.ASSET_ROWS));
  kv(7,'総負債','='+bsRef(CFG.LIAB_ROWS),YEN,'法人は債務超過に注意');
  // 現在
  block(9,'◆ 現在（今の体力）',C.N);
  kv(10,'現預金（流動性ハイ）','='+cashRef(),YEN,'即現金化できる額');
  sh.getRange(11,1).setValue('ランウェイ').setFontWeight('bold').setBackground(C.SUBF);
  sh.getRange(11,2).setFormula('=B10/'+CFG.BURN).setNumberFormat('0.0"ヶ月"').setFontWeight('bold');
  sh.getRange(11,3).setFormula('=IF(B11<3,"🔴守り:投資凍結",IF(B11<6,"🟡注意","🟢攻めOK"))').setFontWeight('bold');
  // 未来
  block(13,'◆ 未来（予測）',C.F);
  sh.getRange(14,1).setValue('谷／通期予測').setFontWeight('bold').setBackground(C.SUBF);
  sh.getRange(14,2,1,3).merge().setValue('詳細は既存「資金繰り」タブで6ヶ月先まで月末現金がプラスか確認。').setWrap(true);
  // 行動
  block(17,'◆ 行動（時間×金×家族・②連動）',C.A);
  sh.getRange(18,1).setValue('★営業直結比率(目標60%)').setFontWeight('bold').setBackground(C.SUBF);
  sh.getRange(18,2).setFormula('='+MX+'!C'+ME).setNumberFormat(PCT).setFontWeight('bold').setBackground(C.HI);
  sh.getRange(18,3).setValue('不動産+医療+テレアポ。内務に溶けてないか毎週確認。').setWrap(true);
  sh.getRange(19,1).setValue('今月 期待粗利→家族換算').setFontWeight('bold').setBackground(C.FAM);
  sh.getRange(19,2).setFormula('='+MX+'!G'+MT).setNumberFormat(YEN).setFontWeight('bold').setBackground(C.FAM);
  sh.getRange(19,3).setFormula('=CONCATENATE("＝家族 ",TEXT('+MX+'!I'+MT+',"0.0"),"ヶ月分の暮らし")').setFontWeight('bold').setBackground(C.FAM);
  sh.getRange(20,1).setValue('今月 家族時間(目的)／総実績h').setFontWeight('bold').setBackground(C.SUBF);
  sh.getRange(20,2).setFormula('='+MX+'!C13').setNumberFormat(HRS).setFontWeight('bold').setBackground(C.FAM);
  sh.getRange(20,3).setFormula('='+MX+'!C'+MT).setNumberFormat(HRS);
  sh.getRange(21,1).setValue('次の一手（手記入）').setFontWeight('bold').setBackground(C.SUBF);
  sh.getRange(21,2,1,3).merge().setValue('内務を削り営業へ／栄町6/20前倒し／医療もう1件').setWrap(true);
  // 判断ポイント
  sh.getRange(23,1).setValue('■ 今月の判断ポイント（絞る）').setFontWeight('bold');
  var pts=['□ 法人現預金の運転資金(EC仕入)と自由現金の切り分け＝真のランウェイ確定',
           '□ ランウェイ🟢6ヶ月以上で初めて資産拡大の投資に現金を回す',
           '□ 時間ROI(円/h)が高い活動に翌週の時間を寄せ、低ROI・内務は縮小'];
  for (var i=0;i<pts.length;i++) sh.getRange(24+i,1,1,5).merge().setValue(pts[i]).setWrap(true);
  sh.setColumnWidth(1,240); sh.setColumnWidth(2,150); sh.setColumnWidth(3,300);
  // ①を先頭へ
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(1);
}

// ============ ⑥ 日次ループ・使い方 ============
function buildDailyLoop(){
  var sh=freshSheet('⑥日次ループ・使い方');
  head(sh,'A1:B1','⑥ 日次ループ — 夜に抽出してタスク化 → 朝にブリーフして再配置');
  var rows=[
   ['🌙 夜ループ（仕込み・抽出→タスク化）','18時以降は家族最優先。家族明けの短時間で'],
   ['  1. その日の進捗・学び・気づきを抽出','秘書がnotes/learningsから抽出も可'],
   ['  2. 抽出をタスク化（誰を・いつまで・どう動かす）','Google Tasks(SSoT)へ。プッシュ営業視点'],
   ['  3. 翌日の最優先1〜3件を確定','明日の起点を1つ決めて寝る'],
   ['',''],
   ['🌅 朝ループ（ブリーフ→再配置）','feedback_morning_brief 準拠'],
   ['  1. カレンダー今週＋Google Tasks未完了全件を統合',''],
   ['  2. ①期限切れ ②今日 ③今週 ④将来 に仕分け','期限切れは件数明示'],
   ['  3. 配置換え：今日/今週/凍結/完了 をYES/NOで決定','営業60%死守'],
   ['  4. ①統合司令塔で過去→現在→未来＋行動を確認','数字を見てから動く'],
   ['',''],
   ['【月初ルーティン】','毎月初の5分'],
   ['  ・BSタブの当月列を更新（CONFIG.CUR_COLを翌列に）','①司令塔が自動連動'],
   ['  ・②本部マトリクスの見込み粗利・確度を最新化',''],
   ['  ・メニュー「実績hをカレンダーから更新」を実行','時間ROIで翌月配分を決める'],
   ['',''],
   ['【沼化の歯止め】','秘書が毎回チェック'],
   ['  ・判断を速くする表示層。配線の作り込みは本業を食う','完成磨き込み禁止'],
   ['',''],
   ['【データ元】','SSoT=このスプシ。財務=BSタブ／借入=0返済予定表が証票'],
   ['  ・実績h=Googleカレンダー(本部番号プレフィックスで自動仕分け)','xlsx廃止・GASネイティブ']
  ];
  sh.getRange(3,1,rows.length,2).setValues(rows).setWrap(true);
  for (var i=0;i<rows.length;i++){
    var a=rows[i][0];
    if(a.indexOf('🌙')===0||a.indexOf('🌅')===0||a.indexOf('【')===0)
      sh.getRange(3+i,1).setFontWeight('bold').setBackground(C.SUBF);
  }
  sh.setColumnWidth(1,360); sh.setColumnWidth(2,330);
}
