/**
 * 02作業DB クリーン化 ＋ 週次棚卸し(05_📊週次KPI)再構築 統合GAS
 * 使い方: script.new に貼付 → runAll を実行（何度でも安全＝冪等）
 *  ① 02をバックアップタブへ退避（当日分が無ければ作成）
 *  ② 本部/ステータス/営業直結の表記ゆれを統一＋ドロップダウン＋O/P/Q/R式を全行に付与
 *  ③ 05_📊週次KPI を「週次棚卸し」に作り替え（先週レビュー＋今週繰越を02から自動集約）
 *     旧週分は週次_アーカイブへ退避してから上書き＝積み上げも残す
 * 位置安全: 列はヘッダ名で特定（D/S/T・O/P/Q/Rの位置を動かさない）
 */
var SS_ID = '1ofLJOFuW5175OHDSy0J0561LzrfAU-M015o7ByGH5wc';

// 本部の表記ゆれ → 正規ラベル
var HONBU_MAP = {
  '0':'00家族','00':'00家族','00:00':'00家族','家族':'00家族',
  '1':'01経営','01':'01経営','経営':'01経営','秘書':'01経営','秘書室':'01経営',
  '2':'02資金','02':'02資金','資金':'02資金',
  '3':'03運営','03':'03運営','運営':'03運営',
  '4':'04コンサル','04':'04コンサル','コンサル':'04コンサル',
  '5':'05物件','05':'05物件','物件':'05物件'
};
var HONBU_LIST = ['00家族','01経営','02資金','03運営','04コンサル','05物件'];
// ステータスの表記ゆれ → 4値
var STATUS_MAP = {
  '未実施':'未完了','実施':'完了','完了(大幅超過)':'完了','統合':'完了','捨(意図的)':'完了',
  '着手':'進行','未':'未着手'
};
var STATUS_LIST = ['未着手','進行','完了','未完了'];

function runAll(){
  var ss = SpreadsheetApp.openById(SS_ID);
  backupDb02_(ss);
  cleanDb02_(ss);
  buildWeeklyReview_(ss);
  SpreadsheetApp.flush();
  Logger.log('runAll 完了');
}

// ── 02検出 ───────────────────────────────────────────────
function findDb02_(ss){
  var s=ss.getSheets();
  for(var i=0;i<s.length;i++){
    var r=Math.min(6,s[i].getLastRow()); if(r<1)continue;
    var v=s[i].getRange(1,1,r,Math.min(40,s[i].getLastColumn())).getValues();
    for(var x=0;x<v.length;x++){var j=v[x].join('|');
      if(j.indexOf('案件・相手')>=0&&j.indexOf('日付')>=0) return {sh:s[i],hRow:x+1};}
  }
  throw '02が見つからない';
}
function hmap_(sh,hRow){var h=sh.getRange(hRow,1,1,sh.getLastColumn()).getValues()[0];var m={};for(var i=0;i<h.length;i++)m[String(h[i]).trim()]=i+1;return m;}

// ① バックアップ ────────────────────────────────────────────
function backupDb02_(ss){
  var f=findDb02_(ss), sh=f.sh;
  var d=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMdd');
  var name='02_backup_'+d;
  if(ss.getSheetByName(name)){ Logger.log('本日バックアップ済: '+name); return; }
  var cp=sh.copyTo(ss); cp.setName(name); cp.hideSheet();
  Logger.log('バックアップ作成: '+name);
}

// ② クリーン化 ──────────────────────────────────────────────
function cleanDb02_(ss){
  var f=findDb02_(ss), sh=f.sh, hRow=f.hRow, m=hmap_(sh,hRow);
  var last=sh.getLastRow(); if(last<=hRow){Logger.log('データ無');return;}
  var n=last-hRow;
  var cHonbu=m['本部'], cStatus=m['ステータス'], cEig=m['営業直結'], cKubun=m['区分'];
  // 本部 正規化
  if(cHonbu){
    var rg=sh.getRange(hRow+1,cHonbu,n,1), v=rg.getValues();
    for(var i=0;i<n;i++){var x=String(v[i][0]).trim(); if(HONBU_MAP[x]) v[i][0]=HONBU_MAP[x];}
    rg.setValues(v);
    rg.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(HONBU_LIST,true).build());
  }
  // ステータス 正規化（空は 予定→未着手 のみ補完。実績の空は触らない）
  if(cStatus){
    var rs=sh.getRange(hRow+1,cStatus,n,1), vs=rs.getValues();
    var vk = cKubun ? sh.getRange(hRow+1,cKubun,n,1).getValues() : null;
    for(var j=0;j<n;j++){var y=String(vs[j][0]).trim();
      if(STATUS_MAP[y]) vs[j][0]=STATUS_MAP[y];
      else if(y==='' && vk && String(vk[j][0]).trim()==='予定') vs[j][0]='未着手';
    }
    rs.setValues(vs);
    rs.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUS_LIST,true).build());
  }
  // 営業直結 正規化（○ はそのまま、- は ✕ に、空はそのまま）＋ドロップダウン(○/✕)
  if(cEig){
    var re=sh.getRange(hRow+1,cEig,n,1), ve=re.getValues();
    for(var k=0;k<n;k++){var z=String(ve[k][0]).trim(); if(z==='-'||z==='×') ve[k][0]='✕';}
    re.setValues(ve);
    re.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['○','✕'],true).build());
  }
  // 区分 ドロップダウン
  if(cKubun) sh.getRange(hRow+1,cKubun,n,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['予定','実績'],true).build());
  // O/P/Q/R 式を全行に付与（実時間KPIの源泉）
  var cKs=m['予定開始'],cKe=m['予定終了'],cMs=m['実開始'],cMe=m['実終了'],
      cO=m['予定所要(分)'],cP=m['実所要(分)'],cQ=m['予実差分(分)'],cR=m['達成率%'];
  if(cO&&cP&&cQ&&cR&&cKs&&cKe&&cMs&&cMe){
    var L=function(c){return columnToLetter_(c);};
    var Ks=L(cKs),Ke=L(cKe),Ms=L(cMs),Me=L(cMe),Po=L(cO),Pp=L(cP);
    for(var r2=hRow+1;r2<=last;r2++){
      sh.getRange(r2,cO).setFormula('=IF('+Ks+r2+'="","",ROUND(('+Ke+r2+'-'+Ks+r2+')*1440,0))');
      sh.getRange(r2,cP).setFormula('=IF('+Ms+r2+'="","",ROUND(('+Me+r2+'-'+Ms+r2+')*1440,0))');
      sh.getRange(r2,cQ).setFormula('=IF(OR('+Po+r2+'="",'+Pp+r2+'="")," ",'+Po+r2+'-'+Pp+r2+')');
      sh.getRange(r2,cR).setFormula('=IF(OR('+Po+r2+'="",'+Pp+r2+'="",'+Pp+r2+'=0)," ",ROUND('+Pp+r2+'/'+Po+r2+'*100,1))');
    }
  }
  Logger.log('cleanDb02 完了 行数='+n);
}
function columnToLetter_(c){var s='';while(c>0){var t=(c-1)%26;s=String.fromCharCode(65+t)+s;c=(c-t-1)/26;}return s;}

// ③ 週次棚卸し ─────────────────────────────────────────────
function buildWeeklyReview_(ss){
  var f=findDb02_(ss), sh=f.sh, hRow=f.hRow, m=hmap_(sh,hRow);
  var last=sh.getLastRow(); var n=Math.max(0,last-hRow);
  var data = n>0 ? sh.getRange(hRow+1,1,n,sh.getLastColumn()).getValues() : [];
  var gv=function(row,name){var c=m[name];return c?row[c-1]:'';};
  var dstr=function(val){ if(val instanceof Date) return Utilities.formatDate(val,'Asia/Tokyo','yyyy-MM-dd'); return String(val).trim().replace(/\//g,'-').slice(0,10); };

  // 週の範囲：今日を含む週の月曜=今週開始／先週=その-7〜-1
  var today=new Date(); var dow=(today.getDay()+6)%7; // Mon=0
  var thisMon=new Date(today); thisMon.setDate(today.getDate()-dow);
  var lastMon=new Date(thisMon); lastMon.setDate(thisMon.getDate()-7);
  var lastSun=new Date(thisMon); lastSun.setDate(thisMon.getDate()-1);
  var fmt=function(d){return Utilities.formatDate(d,'Asia/Tokyo','yyyy-MM-dd');};
  var fmtMD=function(d){return Utilities.formatDate(d,'Asia/Tokyo','M/d');};
  var inRange=function(s,a,b){return s>=fmt(a)&&s<=fmt(b);};

  // 集計
  var agg={}; HONBU_LIST.forEach(function(h){agg[h]={min:0,emin:0,done:0,ng:0,prog:0,todo:0};});
  var doneList=[], carry=[];
  for(var i=0;i<data.length;i++){
    var row=data[i];
    var honbu=String(gv(row,'本部')).trim(); if(HONBU_LIST.indexOf(honbu)<0) honbu=null;
    var ds=dstr(gv(row,'日付'));
    var st=String(gv(row,'ステータス')).trim();
    var eig=String(gv(row,'営業直結')).trim();
    var p=Number(gv(row,'実所要(分)'))||0;
    var aite=String(gv(row,'案件・相手')).trim();
    var kekka=String(gv(row,'結果')).trim();
    var nextA=String(gv(row,'次アクション')).trim();
    var lim=dstr(gv(row,'期限'));
    var kakudo=String(gv(row,'確度ランク(A/B/C/D)')).trim();
    // 先週レビュー（日付が先週内）
    if(honbu && inRange(ds,lastMon,lastSun)){
      agg[honbu].min+=p; if(eig==='○')agg[honbu].emin+=p;
      if(st==='完了')agg[honbu].done++; else if(st==='未完了')agg[honbu].ng++;
      else if(st==='進行')agg[honbu].prog++; else if(st==='未着手')agg[honbu].todo++;
      if(st==='完了' && (kekka||aite)) doneList.push([honbu,ds.slice(5),aite,kekka||'(結果未記入)',eig]);
    }
    // 今週繰越（完了でない＋次アクション or 未完了/進行/未着手）：日付不問でオープンを集約
    if(['未着手','進行','未完了'].indexOf(st)>=0 && (nextA||aite) && honbu){
      carry.push([honbu, (eig==='○'?'○':''), kakudo, aite, nextA||'(次アクション未記入)', lim]);
    }
  }

  // タブ準備：05_📊週次KPI を探す→無ければ作成。旧内容は週次_アーカイブ(MMDD)へ退避
  var wsName='05_📊週次KPI';
  var ws=ss.getSheetByName(wsName);
  if(ws){
    var arch='週次_'+Utilities.formatDate(lastMon,'Asia/Tokyo','MMdd');
    if(!ss.getSheetByName(arch)){ var cp=ws.copyTo(ss); cp.setName(arch); cp.hideSheet(); }
    ws.clear(); ws.clearFormats();
  } else { ws=ss.insertSheet(wsName); }

  var IN='#FFF2CC',AUTO='#EFEFEF',HEAD='#AA2E26',SUB='#F3F3F3';
  var rno=1;
  // タイトル
  ws.getRange(rno,1).setValue('📊 週次棚卸し（先週レビュー '+fmtMD(lastMon)+'〜'+fmtMD(lastSun)+' ／ 今週 '+fmtMD(thisMon)+'〜）')
    .setFontWeight('bold').setFontSize(13); rno+=2;
  // 本部別サマリー
  ws.getRange(rno,1,1,7).setValues([['本部','実働(h)','営業直結(h)','営業比率','完了','未完了/進行','未着手']])
    .setBackground(HEAD).setFontColor('#FFF').setFontWeight('bold'); rno++;
  var tot={min:0,emin:0,done:0,ngprog:0,todo:0};
  HONBU_LIST.forEach(function(h){
    var a=agg[h]; var hh=(a.min/60), eh=(a.emin/60); var ratio=a.min?Math.round(a.emin/a.min*100):0;
    ws.getRange(rno,1,1,7).setValues([[h, round1_(hh), round1_(eh), ratio+'%', a.done, (a.ng+a.prog), a.todo]]);
    tot.min+=a.min; tot.emin+=a.emin; tot.done+=a.done; tot.ngprog+=(a.ng+a.prog); tot.todo+=a.todo; rno++;
  });
  ws.getRange(rno,1,1,7).setValues([['合計', round1_(tot.min/60), round1_(tot.emin/60),
    (tot.min?Math.round(tot.emin/tot.min*100):0)+'%', tot.done, tot.ngprog, tot.todo]])
    .setBackground(SUB).setFontWeight('bold'); rno+=2;
  ws.getRange(rno,1).setValue('※営業比率の目標=60%。先週の総実働'+round1_(tot.min/60)+'h・うち営業'+round1_(tot.emin/60)+'h').setFontColor('#666'); rno+=2;

  // 先週やったこと（完了・実績）
  ws.getRange(rno,1).setValue('✅ 先週やったこと（完了・営業直結◎優先）').setFontWeight('bold').setBackground(SUB); rno++;
  ws.getRange(rno,1,1,5).setValues([['本部','日付','案件・相手','結果','営業']]).setFontWeight('bold').setBackground('#D9EAD3'); rno++;
  doneList.sort(function(a,b){return (b[4]==='○')-(a[4]==='○');});
  if(doneList.length){ ws.getRange(rno,1,doneList.length,5).setValues(doneList); rno+=doneList.length; }
  else { ws.getRange(rno,1).setValue('（先週の完了実績なし）'); rno++; } rno++;

  // 今週へ繰越（オープンタスク自動集約）
  ws.getRange(rno,1).setValue('▶ 今週やること＝繰越（02の未完了/進行/未着手＋次アクション）').setFontWeight('bold').setBackground(SUB); rno++;
  ws.getRange(rno,1,1,7).setValues([['本部','営業','確度','案件・相手','次アクション','期限','🟡今週優先(S/A/B記入)']])
    .setFontWeight('bold').setBackground('#FCE5CD'); rno++;
  carry.sort(function(a,b){ var o=(b[1]==='○')-(a[1]==='○'); if(o)return o; return String(a[2]).localeCompare(String(b[2])); });
  if(carry.length){
    var crows=carry.map(function(c){return [c[0],c[1],c[2],c[3],c[4],c[5],'']});
    ws.getRange(rno,1,crows.length,7).setValues(crows);
    ws.getRange(rno,7,crows.length,1).setBackground(IN); rno+=crows.length;
  } else { ws.getRange(rno,1).setValue('（繰越タスクなし）'); rno++; } rno+=1;

  // 今週の重点3つ（手入力）
  ws.getRange(rno,1).setValue('🎯 今週の重点3つ（手入力）').setFontWeight('bold').setBackground(SUB); rno++;
  for(var z=1;z<=3;z++){ ws.getRange(rno,1).setValue(z+'.'); ws.getRange(rno,2,1,5).merge().setBackground(IN); rno++; }

  // 体裁
  ws.setColumnWidth(1,90); ws.setColumnWidth(4,260); ws.setColumnWidth(5,300);
  ws.getRange(1,1).setNote('毎週月曜に runAll を実行→このタブが先週レビュー＋今週繰越に自動更新。旧週は「週次_MMdd」へ退避（非表示）。');
  // タブを前方へ
  ss.setActiveSheet(ws); ss.moveActiveSheet(2);
  Logger.log('週次棚卸し 再構築完了 完了='+doneList.length+' 繰越='+carry.length);
}
function round1_(x){return Math.round(x*10)/10;}
