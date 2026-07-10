/**
 * 2026_KHD PJ一覧 統合整理用スプシ向け 整理スクリプト（復元保存 2026-06-03）
 * ※財務マスターとは別タスク。別スプシ NEW_SS_ID を操作する。必要時にこのファイルを貼って使う。
 * 機密シートの値はログに出さない。
 */
const NEW_SS_ID = '1uwbe2TTbkmanJAI6gmAFgC_gfyu65dHdtFqd4vLcGow';

/* ========= ユーティリティ ========= */
function ss_() { return SpreadsheetApp.openById(NEW_SS_ID); }
function getOrCreateSheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (sh) { sh.clear(); return sh; }
  return ss.insertSheet(name);
}
function rowsOf_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) return [];
  const r = sh.getLastRow(), c = sh.getLastColumn();
  if (r === 0 || c === 0) return [];
  return sh.getRange(1,1,r,c).getValues();
}

/* ========= (A) パス+解約 統合 ========= */
function reorganizePassKaiyaku() {
  const passRaw = rowsOf_('パス');
  const kaiyakuRaw = rowsOf_('解約');
  const out = getOrCreateSheet_('🔐パス・解約(機密)');

  const headers = ['区分','本部','家/事','分類','名(サービス名)','金額','出口','ID','PASS','連携先','ポイント','備考','元シート','元行番号','入力推定日','解約予定日(概算)'];
  out.getRange(1,1,1,headers.length).setValues([headers]);
  out.setFrozenRows(1);
  out.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#fce5cd');

  const today = new Date();
  const writeRows = [];

  for (let i = 2; i <= passRaw.length; i++) {
    const r = passRaw[i-1] || [];
    const name = (r[3]||'').toString();
    const money = r[4]||'';
    const exit_ = r[5]||'';
    let id = (r[6]||'').toString();
    let pass = (r[7]||'').toString();
    const conn = r[8]||'';
    const point = r[9]||'';
    const memo = r[10]||'';
    if (!name && !id && !pass) continue;

    if (!id && pass) {
      if (/@/.test(pass) || /^[A-Za-z0-9_.\-]+$/.test(pass.split('\n')[0]||'')) {
        const lines = pass.split('\n');
        if (lines.length >= 2) { id = lines[0]; pass = lines.slice(1).join('\n'); }
        else { id = pass; pass = ''; }
      }
    }
    if (id && !pass && /\n/.test(id)) {
      const lines = id.split('\n');
      if (/@/.test(lines[0]) || /^[A-Za-z0-9_.\-]+$/.test(lines[0])) {
        id = lines[0]; pass = lines.slice(1).join('\n');
      }
    }

    const total = passRaw.length - 1;
    const idx = i - 2;
    const startY = new Date(2022, 0, 1).getTime();
    const endY = new Date(2026, 4, 1).getTime();
    const est = total > 0 ? new Date(startY + (endY - startY) * (idx/total)) : '';
    const estStr = est ? Utilities.formatDate(est,'JST','yyyy/MM(推定)') : '';

    let exitEst = '';
    if (exit_) {
      const eDate = new Date(today.getFullYear(), 11, 31);
      exitEst = Utilities.formatDate(eDate,'JST','yyyy/MM(概算)');
    }

    writeRows.push(['パス',r[0]||'',r[1]||'',r[2]||'',name,money,exit_,id,pass,conn,point,memo,'パス',i,estStr,exitEst]);
  }

  for (let i = 2; i <= kaiyakuRaw.length; i++) {
    const r = kaiyakuRaw[i-1] || [];
    if (r.every(v=>v===''||v===null)) continue;
    const all = r.map(function(v,idx2){return '['+String.fromCharCode(65+idx2)+']'+(v===''||v===null?'':v);}).filter(function(s){return s.length>3;}).join(' / ');
    let name = '';
    for (let j=0;j<r.length;j++){if(r[j]){name=String(r[j]);break;}}
    const exitEst = Utilities.formatDate(new Date(today.getFullYear(), 11, 31),'JST','yyyy/MM(概算)');
    writeRows.push(['解約','','','',name,'','解約予定','','','','',all,'解約',i,'',exitEst]);
  }

  if (writeRows.length > 0) {
    out.getRange(2,1,writeRows.length,headers.length).setValues(writeRows);
  }
  out.autoResizeColumns(1, headers.length);
  Logger.log('🔐パス・解約(機密) 統合完了: ' + writeRows.length + '行 (パス'+ (passRaw.length-1) +'件+解約'+ (kaiyakuRaw.length-1) +'件)');
}

/* ========= (B) 社保税務法務ログ (Notion移行用) ========= */
function buildNotionLog() {
  const out = getOrCreateSheet_('📦NOTION移行用_社保税務法務ログ');
  const headers = ['カテゴリ','日付','項目','詳細','金額','期限','ステータス','備考','元シート行'];
  out.getRange(1,1,1,headers.length).setValues([headers]);
  out.setFrozenRows(1);
  out.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#d9ead3');

  const result = [];
  const srcs = [
    {name:'社保', cat:'社保'},
    {name:'税', cat:'税'},
    {name:'⑤税金', cat:'⑤税金'},
    {name:'法務', cat:'法務'}
  ];

  srcs.forEach(function(src){
    const rows = rowsOf_(src.name);
    if (rows.length === 0) return;
    const header = rows[0] || [];
    for (let i=1; i<rows.length; i++){
      const r = rows[i];
      if (!r || r.every(function(v){return v===''||v===null;})) continue;
      let date='',item='',detail='',money='',deadline='',status='',memo='';
      const others = [];
      r.forEach(function(v,idx){
        if (v===''||v===null) return;
        const s = String(v);
        if (v instanceof Date) { if(!date) date = Utilities.formatDate(v,'JST','yyyy/MM/dd'); else if(!deadline) deadline=Utilities.formatDate(v,'JST','yyyy/MM/dd'); else others.push(s); return; }
        if (/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}/.test(s) || /^\d{1,2}\/\d{1,2}/.test(s)) { if(!date) date = s; else if(!deadline) deadline=s; else others.push(s); return; }
        if (typeof v === 'number' && Math.abs(v) > 100) { if(!money) money = v; else others.push(s); return; }
        if (/^[0-9,]+円?$/.test(s)) { if(!money) money = s; else others.push(s); return; }
        if (/未着手|対応中|完了|確認中|済|未/.test(s)) { if(!status) status = s; else others.push(s); return; }
        if (!item) { item = s; return; }
        if (!detail) { detail = s; return; }
        others.push((header[idx]?header[idx]+':':'') + s);
      });
      memo = others.join(' / ');
      result.push([src.cat, date, item, detail, money, deadline, status, memo, src.name+'!行'+(i+1)]);
    }
  });

  if (result.length > 0) {
    out.getRange(2,1,result.length,headers.length).setValues(result);
  }
  out.autoResizeColumns(1, headers.length);
  Logger.log('📦NOTION移行用_社保税務法務ログ 作成完了: ' + result.length + '行');
}

/* ========= (C) 本部別シート + KPI表 ========= */
const KPI_TABLE = {
  '00_プライベート': {mission:'家族・健康・自己投資を最優先に確保（土地家屋調査士勉強・ゴルフ・育児・健康）', sakusen:'カレンダーに家族時間・勉強時間を強制ブロック。週次の積立時間で進捗管理。', kpi:['積立時間(週/月)','家族スマイル時間','勉強総時間/週平均']},
  '01_経営管理': {mission:'資金繰りの完全可視化 / 低利益事業の徹底損切り / スケジュールから内務排除し商談・仕込みで埋める(秘書室統合)', sakusen:'借入管理シート(SSoT)で金利リスク把握。EC事業(韓国輸出・国内転売)の撤退数値確定。朝ブリで「その開発タスク本当に必要?」とブレーキ。カレンダーに「追客作成枠」「セミナー台本練習枠」を強制ブロック。', kpi:['①撤退・損切り判断の実行数(件/月)','②キャッシュアウト削減額(円/月)','③営業・仕込み・商談比率(目標60%以上)']},
  '02_資金調達': {mission:'銀行を「顧客」と見立てた融資枠の最大化営業', sakusen:'税理士への記帳提出を最短化。一発OKを引き出す事業計画書ストーリーを磨く。', kpi:['①銀行への融資打診・面談件数(回/月)','②新規融資の内諾・確保額(円/月)']},
  '03_事業運営': {mission:'既存リストの泥臭い掘り起こし / YouTube等の自動集客罠の構築', sakusen:'cloud mil 11名リストを1人ずつ徹底解剖→個別提案テキストを1文字単位で推敲。YouTubeは技術自慢を捨てsite_v3への導線。', kpi:['①ホット客への個別追客・提案送信数(件/週)','②自社HP経由の有効問い合わせ獲得数(件/月)']},
  '04_コンサルタント': {mission:'セミナーを「教える場」ではなく高単価案件の「営業の場」にする', sakusen:'5/26のTTセミナー(医療専門不動産)台本を「参加者が現状に絶望→相談したくなる」構成へ。最後に限定3名の個別無料診断(商談アポ)へ誘導。', kpi:['①セミナーからの個別相談(本命商談)獲得率(%)','②コンサル案件の見積提示額(円/月)']},
  '05_物件調達': {mission:'バイセル等業者ルートの死守と栄町等PJの決済日程前倒し', sakusen:'主要業者へ定期プッシュし続ける。', kpi:['①主要業者からの新規物件情報獲得数(件/週)','②栄町等PJの決済日程の前倒し達成率(%)']}
};

function buildBuhonSheets() {
  const pjRaw = rowsOf_('PJ');
  let pjHeaderIdx = -1;
  for (let i=0;i<pjRaw.length;i++){
    if ((pjRaw[i]||[]).indexOf('事業部') >= 0) { pjHeaderIdx = i; break; }
  }
  const pjHeader = pjHeaderIdx>=0 ? pjRaw[pjHeaderIdx] : ['','事業部','大項目','小枠','行動タスク','期限','担当者'];
  const pjRows = pjHeaderIdx>=0 ? pjRaw.slice(pjHeaderIdx+1) : [];

  const wbsRaw = rowsOf_('WBS');
  const wbsByBuhon = {};
  let currentBuhon = '';
  for (let i=0;i<wbsRaw.length;i++){
    const a = (wbsRaw[i]&&wbsRaw[i][0])?String(wbsRaw[i][0]):'';
    if (/^0\d_/.test(a)) {
      currentBuhon = a;
      if (!wbsByBuhon[currentBuhon]) wbsByBuhon[currentBuhon] = [];
      continue;
    }
    if (currentBuhon && wbsRaw[i] && wbsRaw[i].some(function(v){return v!=='' && v!==null;})) {
      wbsByBuhon[currentBuhon].push({rowIdx:i+1, row:wbsRaw[i]});
    }
  }

  Object.keys(KPI_TABLE).forEach(function(buhon){
    const sh = getOrCreateSheet_('🏢'+buhon);
    const k = KPI_TABLE[buhon];

    sh.getRange(1,1).setValue('【本部】'+buhon).setFontWeight('bold').setFontSize(14).setBackground('#cfe2f3');
    sh.getRange(2,1).setValue('🎯 ミッション');
    sh.getRange(2,2,1,5).merge().setValue(k.mission).setWrap(true);
    sh.getRange(3,1).setValue('💪 泥臭い作戦');
    sh.getRange(3,2,1,5).merge().setValue(k.sakusen).setWrap(true);
    sh.getRange(4,1).setValue('📊 営業直結KPI');
    sh.getRange(4,2,1,5).merge().setValue(k.kpi.join(' / ')).setWrap(true);
    sh.getRange(1,1,4,1).setFontWeight('bold').setBackground('#fff2cc');
    sh.setColumnWidth(1, 150);
    sh.setRowHeights(2,3,40);

    let row = 6;
    sh.getRange(row,1).setValue('▼ PJ・タスク一覧').setFontWeight('bold').setBackground('#d9ead3');
    row++;
    sh.getRange(row,1,1,pjHeader.length).setValues([pjHeader]).setFontWeight('bold').setBackground('#ead1dc');
    row++;
    const matched = pjRows.filter(function(r){
      const v = (r[1]||'').toString();
      return v && (v.indexOf(buhon)>=0 || v.indexOf(buhon.replace(/^0\d_/,''))>=0);
    });
    if (matched.length>0) {
      sh.getRange(row,1,matched.length,pjHeader.length).setValues(matched);
      row += matched.length;
    } else {
      sh.getRange(row,1).setValue('(該当PJなし - PJシートの「事業部」列に「'+buhon+'」を入れると自動反映)');
      row++;
    }

    row += 2;
    sh.getRange(row,1).setValue('▼ WBS該当ブロック (行参照: WBSシート)').setFontWeight('bold').setBackground('#d9ead3');
    row++;
    const wbsList = wbsByBuhon[buhon] || [];
    if (wbsList.length>0) {
      const wbsHeader = ['行番号','工程','狙い','完了条件','担当','確認','予','実','進捗率','完了フラグ','From','TO'];
      sh.getRange(row,1,1,wbsHeader.length).setValues([wbsHeader]).setFontWeight('bold').setBackground('#ead1dc');
      row++;
      const wbsValues = wbsList.map(function(item){
        const r = item.row;
        return [item.rowIdx, r[0]||'', r[1]||'', r[2]||'', r[3]||'', r[4]||'', r[5]||'', r[6]||'', r[7]||'', r[8]||'', r[9]||'', r[10]||''];
      });
      sh.getRange(row,1,wbsValues.length,wbsHeader.length).setValues(wbsValues);
      row += wbsValues.length;
    } else {
      sh.getRange(row,1).setValue('(該当ブロックなし)');
      row++;
    }

    sh.setFrozenRows(5);
  });
  Logger.log('🏢本部別シート 作成完了: ' + Object.keys(KPI_TABLE).length + '本部');
}

/* ========= (D) 予実・収支・長期計画 集約 ========= */
function buildYojituShushi() {
  const y = getOrCreateSheet_('📊予実集約');
  const ySrcs = ['予実','2025','32-36','yozitu','3月'];
  let yRow = 1;
  ySrcs.forEach(function(name){
    const rows = rowsOf_(name);
    if (rows.length===0) return;
    y.getRange(yRow,1).setValue('======= ['+name+'] '+rows.length+'行 x '+(rows[0]?rows[0].length:0)+'列 =======').setFontWeight('bold').setBackground('#fce5cd');
    yRow++;
    const maxCol = Math.max.apply(null, rows.map(function(r){return r.length;}));
    const norm = rows.map(function(r){ const out=[]; for (let i=0;i<maxCol;i++) out.push(r[i]===undefined?'':r[i]); return out; });
    y.getRange(yRow,1,norm.length,maxCol).setValues(norm);
    yRow += norm.length + 2;
  });
  Logger.log('📊予実集約 完了');

  const s = getOrCreateSheet_('📊収支明細');
  const sSrcs = ['収支','①売上','②原価','③諸経費','クレカ用途'];
  let sRow = 1;
  sSrcs.forEach(function(name){
    const rows = rowsOf_(name);
    if (rows.length===0) return;
    s.getRange(sRow,1).setValue('======= ['+name+'] '+rows.length+'行 x '+(rows[0]?rows[0].length:0)+'列 =======').setFontWeight('bold').setBackground('#fce5cd');
    sRow++;
    const maxCol = Math.max.apply(null, rows.map(function(r){return r.length;}));
    const norm = rows.map(function(r){ const out=[]; for (let i=0;i<maxCol;i++) out.push(r[i]===undefined?'':r[i]); return out; });
    s.getRange(sRow,1,norm.length,maxCol).setValues(norm);
    sRow += norm.length + 2;
  });
  Logger.log('📊収支明細 完了');

  const l = getOrCreateSheet_('📈長期計画');
  const lSrcs = ['長期計画','損益計画','予算'];
  let lRow = 1;
  lSrcs.forEach(function(name){
    const rows = rowsOf_(name);
    if (rows.length===0) return;
    l.getRange(lRow,1).setValue('======= ['+name+'] '+rows.length+'行 x '+(rows[0]?rows[0].length:0)+'列 =======').setFontWeight('bold').setBackground('#fce5cd');
    lRow++;
    const maxCol = Math.max.apply(null, rows.map(function(r){return r.length;}));
    const norm = rows.map(function(r){ const out=[]; for (let i=0;i<maxCol;i++) out.push(r[i]===undefined?'':r[i]); return out; });
    l.getRange(lRow,1,norm.length,maxCol).setValues(norm);
    lRow += norm.length + 2;
  });
  Logger.log('📈長期計画 完了');
}

/* ========= (E) 欠損チェック ========= */
function verifyAndReport() {
  const ss = ss_();
  const all = ss.getSheets();
  Logger.log('=== 最終シート一覧 ===');
  all.forEach(function(sh){
    Logger.log(' - ' + sh.getName() + ' ('+sh.getLastRow()+'行 x '+sh.getLastColumn()+'列)');
  });
  Logger.log('=== 新スプシURL: ' + ss.getUrl() + ' ===');
}

/* ========= メイン: 一気通貫 ========= */
function runAll() {
  reorganizePassKaiyaku();
  buildNotionLog();
  buildBuhonSheets();
  buildYojituShushi();
  verifyAndReport();
}
