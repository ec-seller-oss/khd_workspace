/**
 * KHD 追客パイプライン自動化 GAS v1（2026-06-03）
 * -----------------------------------------------------------
 * 何をするか（沼回避のため v1 はこの2本に絞る）:
 *  1) syncFollowupTasks() … ①追客リストの各行を見て、
 *       「状態=対応中 / 次アクション・期限あり / まだタスク未作成」の行に
 *       Googleタスクを自動生成し、タスクIDを行に書き戻す（＝強制力の自動化）。
 *       状態が「成約/失注」に変わった行は、対応するタスクを自動で完了/削除。
 *  2) updateDashboard() … KPIダッシュボードtabを自動更新
 *       （今週の追客送信数 / LINE登録率 / ステージ別 / 温度別 / SLA期限切れ / 成約率）。
 *
 * Phase2（未実装・別途）: ポータル反響メール→新規行の自動追記。
 *
 * セットアップ:
 *  - スクリプトエディタにこのファイルを貼付
 *  - サービス＋「Tasks API」を追加（医療マッチングGASと同様）
 *  - setupTriggers() を1回実行して承認（10分ごとに2本が回る）
 * 実行アカウント: ec-seller@kikuchi-hd.net
 */

// ===== 設定 =====
const TK_SHEET_ID   = '1PWI3JyEVaFsTZqOA_DNijBIDtGHMZM2JRhdHORkKQhY';
const TK_TASKLIST   = 'MDI3NjUwNTM4MzM0MzY1Mjg1NzA6MDow'; // マイタスク
const TK_LIST_TAB   = '①追客リスト';   // 無ければ index 0 にフォールバック
const TK_LOG_TAB    = '②接触ログ';     // 無ければ index 1
const TK_DASH_TAB   = 'KPIダッシュボード'; // 無ければ自動作成
const TK_GID        = '1877162925';     // スプシリンクのgid（タスクnotes用）

// 列番号（1始まり）。①追客リストのヘッダ順に対応
const C = {
  id:1, name:2, channel:3, inq:4, stage:5, temp:6, line:7,
  must:8, want:9, pain:10, when:11, last:12,
  next:13, due:14, state:15, owner:16, memo:17, taskId:18 // taskId列は無ければ自動追加
};

function _sheet(name, idx) {
  const ss = SpreadsheetApp.openById(TK_SHEET_ID);
  return ss.getSheetByName(name) || ss.getSheets()[idx];
}
function _sheetUrl(rowNum) {
  return 'https://docs.google.com/spreadsheets/d/' + TK_SHEET_ID +
         '/edit?gid=' + TK_GID + '&range=A' + rowNum;
}

// ============================================================
// 1) 追客リスト → タスク自動生成（強制力）
// ============================================================
function syncFollowupTasks() {
  const sh = _sheet(TK_LIST_TAB, 0);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  // taskId列のヘッダを保証（無ければ作る）
  const header = sh.getRange(1, 1, 1, Math.max(C.taskId, sh.getLastColumn())).getValues()[0];
  if (!header[C.taskId - 1]) sh.getRange(1, C.taskId).setValue('タスクID');

  const width = C.taskId;
  const data = sh.getRange(2, 1, lastRow - 1, width).getValues();

  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const rowNum = i + 2;
    const name = (r[C.name-1]||'').toString().trim();
    if (!name || name.indexOf('（例）') === 0) continue; // 例行はスキップ

    const state  = (r[C.state-1]||'').toString().trim();
    const taskId = (r[C.taskId-1]||'').toString().trim();
    const next   = (r[C.next-1]||'').toString().trim();
    const due    = r[C.due-1];

    // (a) 成約/失注 → 既存タスクをクローズ
    if (taskId && (state === '成約' || state === '失注')) {
      try {
        if (state === '成約') {
          const t = Tasks.Tasks.get(TK_TASKLIST, taskId);
          t.status = 'completed';
          Tasks.Tasks.update(t, TK_TASKLIST, taskId);
        } else {
          Tasks.Tasks.remove(TK_TASKLIST, taskId);
        }
        sh.getRange(rowNum, C.taskId).setValue(''); // 二重処理防止でクリア
      } catch (e) { /* 既に消えている等は無視 */ }
      continue;
    }

    // (a2) 対応中でタスク完了済み → 「1接触」として記録（KPIはタスク側で集計）
    //      メモに履歴を残し・最終接触日を更新・行を解放（次アクションを書けば次のタスクが湧く）
    if (taskId && state === '対応中') {
      try {
        const t = Tasks.Tasks.get(TK_TASKLIST, taskId);
        if (t && t.status === 'completed') {
          const memoCell = sh.getRange(rowNum, C.memo);
          const stamp = Utilities.formatDate(new Date(),'Asia/Tokyo','MM/dd');
          const hist = '[' + stamp + '済]' + (next||'');
          memoCell.setValue((((memoCell.getValue()||'') + ' / ' + hist)).slice(0,500));
          sh.getRange(rowNum, C.last).setValue(_tkToday()); // 最終接触日を自動更新
          sh.getRange(rowNum, C.taskId).setValue('');        // 行を解放
          sh.getRange(rowNum, C.next).setValue('');          // 次アクション完了→クリア
          sh.getRange(rowNum, C.due).setValue('');
          continue;
        }
      } catch (e) { sh.getRange(rowNum, C.taskId).setValue(''); continue; } // タスク消失なら解放
    }

    // (b) 対応中 / 次アクション・期限あり / タスク未作成 → 新規タスク
    if (state === '対応中' && next && due && !taskId) {
      const stage = (r[C.stage-1]||'').toString().trim();
      const temp  = (r[C.temp-1]||'').toString().trim();
      const title = ('【追客】' + name + ' ' + stage + '｜' + next).slice(0, 1024);
      const notes = '温度:' + temp +
                    '\nMUST:' + (r[C.must-1]||'') +
                    '\nPAIN:' + (r[C.pain-1]||'') +
                    '\n担当:' + (r[C.owner-1]||'') +
                    '\n追客管理シート(' + (r[C.id-1]||'') + '行): ' + _sheetUrl(rowNum);
      const task = { title: title, notes: notes, status: 'needsAction' };
      const dueDate = _toDue(due);
      if (dueDate) task.due = dueDate;
      try {
        const created = Tasks.Tasks.insert(task, TK_TASKLIST);
        sh.getRange(rowNum, C.taskId).setValue(created.id);
      } catch (e) {
        Logger.log('task insert失敗 row' + rowNum + ': ' + e);
      }
    }
  }
}

// 期限セル(Date or "YYYY-MM-DD")→RFC3339(午前0時UTC)
function _toDue(v) {
  let d;
  if (Object.prototype.toString.call(v) === '[object Date]') d = v;
  else {
    const s = (v||'').toString().trim();
    if (!s) return null;
    d = new Date(s.replace(/\//g,'-') + 'T00:00:00Z');
  }
  if (isNaN(d.getTime())) return null;
  return Utilities.formatDate(d, 'UTC', "yyyy-MM-dd'T'00:00:00'Z'");
}

// ============================================================
// 2) KPIダッシュボード自動更新
// ============================================================
function updateDashboard() {
  const ss = SpreadsheetApp.openById(TK_SHEET_ID);
  let dash = ss.getSheetByName(TK_DASH_TAB);
  if (!dash) dash = ss.insertSheet(TK_DASH_TAB, 0); // 先頭に作成

  const list = _sheet(TK_LIST_TAB, 0);
  const lr = list.getLastRow();
  const rows = lr >= 2 ? list.getRange(2, 1, lr-1, C.taskId).getValues()
                          .filter(r => (r[C.name-1]||'').toString().trim() &&
                                       (r[C.name-1]||'').toString().indexOf('（例）') !== 0) : [];

  // 集計
  const today = new Date(); today.setHours(0,0,0,0);
  const cnt = {stage:{}, temp:{}, state:{}, line:{done:0,total:0}, slaOver:0};
  rows.forEach(r => {
    const stg=(r[C.stage-1]||'').toString().trim();
    const tp =(r[C.temp-1]||'').toString().split('(')[0].trim();
    const st =(r[C.state-1]||'').toString().trim();
    const ln =(r[C.line-1]||'').toString().trim();
    const due= r[C.due-1];
    if(stg) cnt.stage[stg]=(cnt.stage[stg]||0)+1;
    if(tp)  cnt.temp[tp]=(cnt.temp[tp]||0)+1;
    if(st)  cnt.state[st]=(cnt.state[st]||0)+1;
    cnt.line.total++; if(ln==='済') cnt.line.done++;
    if(st==='対応中' && due){ const d=_dateOf(due); if(d && d<today) cnt.slaOver++; }
  });

  // 今週(月〜)の追客数 ＝ 完了した【追客】タスク件数（チャネル不問・B方式）
  const monday = _mondayOf(today);
  const mondayIso = Utilities.formatDate(monday,'UTC',"yyyy-MM-dd'T'00:00:00'Z'");
  let weekSends = 0;
  try {
    let pageToken = null;
    do {
      const resp = Tasks.Tasks.list(TK_TASKLIST, {showCompleted:true, showHidden:true, completedMin:mondayIso, maxResults:100, pageToken:pageToken});
      (resp.items||[]).forEach(t => { if(t.title && t.title.indexOf('【追客】')===0 && t.status==='completed') weekSends++; });
      pageToken = resp.nextPageToken;
    } while (pageToken);
  } catch(e){ Logger.log('完了タスク集計失敗:'+e); }

  const lineRate = cnt.line.total ? Math.round(cnt.line.done/cnt.line.total*100) : 0;
  const won = cnt.state['成約']||0, lost = cnt.state['失注']||0;
  const winRate = (won+lost) ? Math.round(won/(won+lost)*100) : 0;

  // 描画
  dash.clear();
  const out = [];
  out.push(['📊 追客KPIダッシュボード（自動更新）', '', Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd HH:mm')+' 時点']);
  out.push(['', '', '']);
  out.push(['■ 先行KPI', '値', '']);
  out.push(['今週の追客数（完了タスク）', weekSends, '']);
  out.push(['LINE登録率', lineRate + '%', cnt.line.done + '/' + cnt.line.total]);
  out.push(['アクティブ追客数（対応中）', cnt.state['対応中']||0, '']);
  out.push(['🔴 SLA期限切れ（対応中で期限超過）', cnt.slaOver, '']);
  out.push(['成約率（成約/(成約+失注)）', winRate + '%', '成約'+won+'/失注'+lost]);
  out.push(['', '', '']);
  out.push(['■ ステージ別件数', '', '']);
  ['①即時接触','②LINE移行','③悩みヒアリング','④物件提案','⑤アポ調整','⑥内見','⑦クロージング','⑧申込〜成約','⑨入居後/紹介']
    .forEach(s => out.push([s, cnt.stage[s]||0, '']));
  out.push(['', '', '']);
  out.push(['■ 温度別件数', '', '']);
  ['HOT','WARM','COLD'].forEach(t => out.push([t, cnt.temp[t]||0, '']));

  dash.getRange(1,1,out.length,3).setValues(out);
  dash.getRange(1,1,1,3).merge().setBackground('#AA2E26').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(13);
  dash.getRange(3,1,1,2).setBackground('#F9F6EF').setFontWeight('bold');
  dash.getRange(10,1).setFontWeight('bold');
  dash.getRange(21,1).setFontWeight('bold');
  dash.setColumnWidth(1,260); dash.setColumnWidth(2,120); dash.setColumnWidth(3,200);
}

function _dateOf(v){
  if(Object.prototype.toString.call(v)==='[object Date]') { const d=new Date(v); d.setHours(0,0,0,0); return d; }
  const s=(v||'').toString().trim(); if(!s) return null;
  const d=new Date(s.replace(/\//g,'-').slice(0,10)+'T00:00:00'); return isNaN(d)?null:d;
}
function _mondayOf(d){ const x=new Date(d); const day=(x.getDay()+6)%7; x.setDate(x.getDate()-day); x.setHours(0,0,0,0); return x; }
function _tkToday(){ return Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd'); }

// ============================================================
// セットアップ（1回だけ実行して承認）
// ============================================================
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['syncFollowupTasks','updateDashboard'].indexOf(t.getHandlerFunction())>=0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncFollowupTasks').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('updateDashboard').timeBased().everyHours(1).create();
  // 初回手動実行
  syncFollowupTasks();
  updateDashboard();
}

// ============================================================
// ↓↓↓ 賃貸反響 P1(費用/ヒアリング/返信)＋P2(マイソク抽出)＋P3(ニーズ判定の脳・菊池ボイス) ↓↓↓
// ============================================================

/**
 * 賃貸反響 自動対応 P1（2026-06-05）─ Claude API不要のMVP
 * -----------------------------------------------------------
 * 「賃貸反響」タブに主要値を1行入力 → メニュー実行で:
 *   ① 初期費用概算（計算式を再現）  ② 仲介ヒアリング事項
 *   ③ 返信文（羽鳥流・遅延謝罪/誠実） ④ 管理会社確認リスト（マイソク欠落項目）
 * を同じ行の出力列に自動生成し、さらに「①追客リスト」へ顧客を集約ログする。
 *
 * 設計原則（抜け漏れ防止）:
 *   マイソクに明記された値だけ計算に入れる。鍵交換・消毒・書類作成費・保証会社指定・
 *   更新料などは金額を入れず「管理会社確認」として④に列挙する。
 *
 * 完成例で検証: 中目黒502 / 6-29入居 → 合計 551,833円 と一致。
 *
 * セットアップ: スクリプトエディタに貼付 → 保存 → シートを再読込 → 上部メニュー
 *   「賃貸反響P1」→「この行を生成」。Tasks APIは集約タスク生成に使用（追加推奨）。
 */

const CT_SHEET_ID = '1PWI3JyEVaFsTZqOA_DNijBIDtGHMZM2JRhdHORkKQhY';
const CT_IN_TAB   = '賃貸反響';        // 入力＋出力（無ければ自動作成）
const CT_LIST_TAB = '①追客リスト';     // 集約先（無ければ index0）
const CT_TASKLIST = 'MDI3NjUwNTM4MzM0MzY1Mjg1NzA6MDow';

// 賃貸反響タブの列定義（1始まり）
const T = {
  date:1, name:2, channel:3, bukken:4, chin:5, kanri:6,
  shiki:7, rei:8, hosho:9, hoshoRate:10, kasai:11, nyukyo:12, years:13, youbou:14,
  // 出力
  outTotal:15, outMeisai:16, outHearing:17, outReply:18, outKakunin:19, outDate:20,
  // P2: マイソクPDFのDriveリンク（ここを埋めると物件・費用が自動抽出される）
  maisoku:21,
  // P3: 文面からのAI判定（自走の脳）
  inquiry:22,   // 客の問合せ文面（生）
  aiKakudo:23,  // 確度(温度+%)
  aiNeeds:24,   // ニーズの本質
  aiKubun:25,   // 区分(新規開拓/追客)
  aiApproach:26,// 推奨アプローチ
  aiReplies:27  // 複数返信案
};
const CT_HEADERS = ['受信日','顧客名','流入','物件名・号室','賃料','管理費',
  '敷(ヶ月)','礼(ヶ月)','保証金(ヶ月)','初回保証料%','火災保険','入居日','契約年数','お客様の要望',
  '初期費用合計','初期費用明細','ヒアリング事項','返信文','管理会社確認','生成日','マイソクPDF(Driveリンク)',
  '問合せ文面','AI確度','ニーズ本質','区分','推奨アプローチ','複数返信案'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('賃貸反響P1')
    .addItem('★マイソクから自動入力＋生成', 'ct_autoFromMaisoku')
    .addItem('★文面からニーズ判定＋複数返信案', 'ct_aiAdvise')
    .addSeparator()
    .addItem('この行を生成（手入力後）', 'ct_generateActiveRow')
    .addItem('未処理を一括生成', 'ct_generateAll')
    .addToUi();
}

function _ctSheet() {
  const ss = SpreadsheetApp.openById(CT_SHEET_ID);
  let sh = ss.getSheetByName(CT_IN_TAB);
  if (!sh) {
    sh = ss.insertSheet(CT_IN_TAB);
    sh.getRange(1,1,1,CT_HEADERS.length).setValues([CT_HEADERS])
      .setFontWeight('bold').setBackground('#AA2E26').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ct_generateActiveRow() {
  const sh = _ctSheet();
  const r = sh.getActiveRange().getRow();
  if (r < 2) { SpreadsheetApp.getUi().alert('データ行を選択してください'); return; }
  _ctProcess(sh, r);
  SpreadsheetApp.getUi().alert('生成しました（行' + r + '）。返信文・初期費用・確認事項を出力列に記入しました。');
}

function ct_generateAll() {
  const sh = _ctSheet();
  const last = sh.getLastRow();
  let n = 0;
  for (let r = 2; r <= last; r++) {
    const name = sh.getRange(r, T.name).getValue();
    const done = sh.getRange(r, T.outDate).getValue();
    if (name && !done) { _ctProcess(sh, r); n++; }
  }
  SpreadsheetApp.getUi().alert(n + '行を生成しました。');
}

function _num(v){ const n = Number(String(v).replace(/[,，円]/g,'').trim()); return isNaN(n)?0:n; }
function _daysInMonth(d){ return new Date(d.getFullYear(), d.getMonth()+1, 0).getDate(); }

function _ctProcess(sh, r) {
  const g = c => sh.getRange(r, c).getValue();
  const name = g(T.name), bukken = g(T.bukken);
  const chin = _num(g(T.chin)), kanri = _num(g(T.kanri));
  const shikiM = _num(g(T.shiki)), reiM = _num(g(T.rei)), hoshoM = _num(g(T.hosho));
  const rate = _num(g(T.hoshoRate)), kasai = _num(g(T.kasai));
  const years = _num(g(T.years)) || 2;
  const youbou = (g(T.youbou)||'').toString();

  // 入居日
  let nyukyo = g(T.nyukyo);
  if (!(Object.prototype.toString.call(nyukyo)==='[object Date]')) {
    const s = (nyukyo||'').toString().replace(/[年月]/g,'-').replace(/日/,'').replace(/\//g,'-');
    nyukyo = s ? new Date(s) : null;
  }

  // === 初期費用 計算（マイソク明記分のみ） ===
  const shiki = chin*shikiM, rei = chin*reiM, hosho = chin*hoshoM;
  const hoshoryo = Math.round((chin+kanri)*rate/100);
  const chukai = Math.round(chin*1.1);
  const chukaiTax = Math.round(chin*0.1);
  let maeChinTo=0, maeChinNext=0, maeKanTo=0, maeKanNext=0, dRows='';
  if (nyukyo && !isNaN(nyukyo.getTime())) {
    const dim = _daysInMonth(nyukyo);
    const days = dim - nyukyo.getDate() + 1; // 入居日〜月末
    maeChinTo = Math.round(chin/dim*days);
    maeKanTo  = Math.round(kanri/dim*days);
    maeChinNext = chin; maeKanNext = kanri;
    dRows = '・前払家賃 当月('+days+'日)：'+maeChinTo.toLocaleString()+'　翌月：'+maeChinNext.toLocaleString()+'\n'
          + '・前払共益 当月：'+maeKanTo.toLocaleString()+'　翌月：'+maeKanNext.toLocaleString()+'\n';
  } else {
    dRows = '・前払家賃／共益：入居日 未入力（要記入）\n';
  }
  const total = shiki+rei+hosho+maeChinTo+maeChinNext+maeKanTo+maeKanNext+hoshoryo+kasai+chukai;

  const meisai =
    '【初期費用概算】物件:'+bukken+'\n'+
    '・敷金('+shikiM+'ヶ月)：'+shiki.toLocaleString()+'\n'+
    '・礼金('+reiM+'ヶ月)：'+rei.toLocaleString()+'\n'+
    (hoshoM? '・保証金('+hoshoM+'ヶ月)：'+hosho.toLocaleString()+'\n':'')+
    dRows+
    '・初回保証料('+rate+'%)：'+hoshoryo.toLocaleString()+'\n'+
    '・火災保険：'+kasai.toLocaleString()+'\n'+
    '・仲介手数料(1ヶ月+税)：'+chukai.toLocaleString()+'（内税'+chukaiTax.toLocaleString()+'）\n'+
    '━━ 概算合計：'+total.toLocaleString()+' 円 ━━\n'+
    '※鍵交換・消毒等は含めていません（管理会社確認）';

  // === ヒアリング事項（羽鳥②＋業界ルール） ===
  const hearing =
    '■ MUST（絶対条件）：入居時期／予算上限／エリア・駅距離／人数\n'+
    '■ WANT（あれば）：階数・日当たり・設備の希望\n'+
    '■ PAIN（本音）：前のお部屋で困った点・引越し理由（←魔法の質問）\n'+
    '■ 申込前提：保証会社審査・必要書類・勤務先\n'+
    (youbou? '■ 既出の要望：'+youbou+'\n':'')+
    '■ 内見：居住中のため日程調整。候補日2〜3を確保';

  // === 返信文（①即時接触・遅延謝罪・誠実） ===
  const reply =
    (name||'○○')+'様\n\n'+
    'はじめまして。Claud mil（クラウドミル）の菊池と申します。\n'+
    '「'+bukken+'」へお問い合わせいただきありがとうございます。ご連絡が遅くなり申し訳ございません。\n\n'+
    'お部屋の概要と、参考の初期費用概算（約'+Math.round(total/10000)+'万円）をお送りします。\n'+
    '・賃料 '+chin.toLocaleString()+'円＋管理費 '+kanri.toLocaleString()+'円'+(reiM===0?'／礼金なし':'')+'\n\n'+
    '差し支えなければ、ご入居時期・人数・特に重視される点を教えていただけますか。よりぴったりのご提案ができます。\n'+
    '内見もご案内できます。居住中のため、ご希望の候補日を2〜3いただければすぐ先方に確認します。LINEでもお気軽にどうぞ。\n\n'+
    'よろしくお願いいたします。\nClaud mil　菊池';

  // === 管理会社確認リスト（マイソク欠落の定番） ===
  const kakunin =
    '【管理会社（元付）確認】\n'+
    '□ 空き状況・申込1番手の有無\n□ 入居可能日の確定\n□ 鍵交換代\n□ 室内消毒・抗菌費\n'+
    '□ 書類作成費・事務手数料\n□ 指定保証会社・料率(初回/月額/更新)\n□ 更新料\n'+
    '□ 内見可否・段取り（居住中）\n□ ペット／事務所利用／楽器 可否\n□ 駐車場の有無・賃料';

  // 出力
  sh.getRange(r, T.outTotal).setValue(total);
  sh.getRange(r, T.outMeisai).setValue(meisai);
  sh.getRange(r, T.outHearing).setValue(hearing);
  sh.getRange(r, T.outReply).setValue(reply);
  sh.getRange(r, T.outKakunin).setValue(kakunin);
  sh.getRange(r, T.outDate).setValue(Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd HH:mm'));

  // === ①追客リストへ集約ログ（賃貸も統一顧客一覧に） ===
  _ctLogToList(sh, r, name, g(T.channel), g(T.date), bukken, chin, total);
}

function _ctLogToList(srcSh, r, name, channel, date, bukken, chin, total) {
  const ss = SpreadsheetApp.openById(CT_SHEET_ID);
  const list = ss.getSheetByName(CT_LIST_TAB) || ss.getSheets()[0];
  // 既に同名×同物件が居れば二重登録しない（簡易）
  const last = list.getLastRow();
  if (last >= 2) {
    const ex = list.getRange(2,2,last-1,1).getValues().map(x=>x[0]);
    for (let i=0;i<ex.length;i++){
      if (ex[i] && bukken && (list.getRange(i+2,17).getValue()||'').toString().indexOf(bukken)>=0
          && ex[i].toString()===name.toString()) return;
    }
  }
  const today = Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd');
  const due   = Utilities.formatDate(new Date(Date.now()+86400000),'Asia/Tokyo','yyyy-MM-dd');
  // ①追客リスト列順: 顧客ID,顧客名,流入,反響日時,ステージ,温度,LINE登録,MUST,WANT,PAIN,希望時期,最終接触,次アクション,期限,状態,担当,メモ
  list.appendRow([
    'R'+Utilities.formatDate(new Date(),'Asia/Tokyo','MMddHHmm'),
    name, channel||'ITANDI', date||today, '①即時接触', 'WARM(週1-2)', '未',
    bukken+'／賃料'+_num(chin).toLocaleString(), '', '', '', today,
    '初回返信を送る→内見打診', due, '対応中', '菊池(Claud mil)',
    '【賃貸】初期費用概算 約'+Math.round(total/10000)+'万円。元付に空き/1番手確認'
  ]);
}

// ============================================================
// P2: マイソクPDF → Claude API(vision)で自動抽出 → 入力欄を自動充填 → 生成
//   使い方: 賃貸反響タブの行に「顧客名」と「マイソクPDF(Driveリンク)(U列)」だけ入れ、
//           その行を選んでメニュー「★マイソクから自動入力＋生成」。
//   必要: スクリプトプロパティ ANTHROPIC_API_KEY（平文でコードに書かない）。
//         入居日は概算の日割りに必要なので、空なら後で手入力（マイソクは"下旬"等で曖昧なため）。
// ============================================================
function ct_autoFromMaisoku() {
  const sh = _ctSheet();
  const r = sh.getActiveRange().getRow();
  if (r < 2) { SpreadsheetApp.getUi().alert('賃貸反響タブのデータ行を選択してください'); return; }
  const link = (sh.getRange(r, T.maisoku).getValue()||'').toString().trim();
  if (!link) { SpreadsheetApp.getUi().alert('U列「マイソクPDF(Driveリンク)」を入れてください'); return; }
  const fileId = _fileIdFromLink(link);
  if (!fileId) { SpreadsheetApp.getUi().alert('DriveのPDFリンクが解釈できません'); return; }

  let d;
  try { d = _extractMaisoku(fileId); }
  catch (e) { SpreadsheetApp.getUi().alert('抽出失敗: ' + e + '\n（ANTHROPIC_API_KEY 未設定なら設定してください）'); return; }

  // 抽出値を入力欄へ（空欄のみ上書きしない＝既入力を尊重）
  const set = (col, val) => { if (val!=='' && val!=null && !sh.getRange(r,col).getValue()) sh.getRange(r,col).setValue(val); };
  set(T.bukken,   d['物件名']);
  set(T.chin,     d['賃料']);
  set(T.kanri,    d['管理費']);
  set(T.shiki,    d['敷金ヶ月']);
  set(T.rei,      d['礼金ヶ月']);
  set(T.hosho,    d['保証金ヶ月']);
  set(T.hoshoRate,d['初回保証料率']);
  set(T.kasai,    d['火災保険']);
  if (d['入居可能'] && !sh.getRange(r,T.nyukyo).getValue()) {
    sh.getRange(r,T.youbou).setValue((sh.getRange(r,T.youbou).getValue()||'')+' [入居可能:'+d['入居可能']+'(要・契約開始日確定)]');
  }

  _ctProcess(sh, r);
  const miss = sh.getRange(r,T.nyukyo).getValue() ? '' : '\n※入居日(契約開始日)が空です。日割り精算のため入力後にもう一度「この行を生成」を。';
  SpreadsheetApp.getUi().alert('マイソクから自動入力＋生成しました（行'+r+'）。'+miss);
}

function _fileIdFromLink(s) {
  let m = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/);  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);     if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;
  return null;
}

function _extractMaisoku(fileId) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('ANTHROPIC_API_KEY 未設定');
  const blob = DriveApp.getFileById(fileId).getBlob();
  const b64  = Utilities.base64Encode(blob.getBytes());
  const prompt =
    'これは日本の賃貸物件のマイソク(物件チラシ)PDFです。次の項目を読み取り、JSONのみで返答してください。' +
    '数値は半角整数(カンマ無し)。記載が無い項目はnull。礼金/敷金は「ヶ月数」、初回保証料は「%の数値」。' +
    'キー: {"物件名":string(物件名+号室),"賃料":number,"管理費":number,"敷金ヶ月":number,"礼金ヶ月":number,' +
    '"保証金ヶ月":number,"初回保証料率":number,"火災保険":number,"入居可能":string}';
  const body = {
    model: 'claude-sonnet-4-6',  // ← 必要に応じて変更可
    max_tokens: 1024,
    messages: [{ role:'user', content:[
      { type:'document', source:{ type:'base64', media_type:'application/pdf', data:b64 } },
      { type:'text', text: prompt }
    ]}]
  };
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:'post', contentType:'application/json',
    headers:{ 'x-api-key':key, 'anthropic-version':'2023-06-01' },
    payload: JSON.stringify(body), muteHttpExceptions:true
  });
  if (res.getResponseCode() !== 200) throw new Error('API '+res.getResponseCode()+': '+res.getContentText().slice(0,300));
  const j = JSON.parse(res.getContentText());
  const text = (j.content && j.content[0] && j.content[0].text) || '';
  const mm = text.match(/\{[\s\S]*\}/);
  if (!mm) throw new Error('JSON抽出失敗: '+text.slice(0,200));
  return JSON.parse(mm[0]);
}

// ============================================================
// P3: 文面からニーズ本質・確度を判定し、複数アプローチ返信案を出す（自走の脳）
//   入力: 流入経路 + 問合せ文面(V列、無ければお客様の要望) + 物件
//   出力: AI確度/ニーズ本質/区分(新規開拓・追客)/推奨アプローチ/複数返信案
//   要: ANTHROPIC_API_KEY
// ============================================================
function ct_aiAdvise() {
  const sh = _ctSheet();
  const r = sh.getActiveRange().getRow();
  if (r < 2) { SpreadsheetApp.getUi().alert('データ行を選択してください'); return; }
  const channel = (sh.getRange(r, T.channel).getValue()||'').toString();
  const inquiry = (sh.getRange(r, T.inquiry).getValue() || sh.getRange(r, T.youbou).getValue() || '').toString();
  const bukken  = (sh.getRange(r, T.bukken).getValue()||'').toString();
  if (!inquiry) { SpreadsheetApp.getUi().alert('V列「問合せ文面」（または「お客様の要望」）を入れてください'); return; }

  let a;
  try { a = _aiAdvise(channel, inquiry, bukken); }
  catch (e) { SpreadsheetApp.getUi().alert('AI判定失敗: ' + e); return; }

  sh.getRange(r, T.aiKakudo).setValue((a['確度']||'')+(a['確度パーセント']!=null?'（'+a['確度パーセント']+'%）':''));
  sh.getRange(r, T.aiNeeds).setValue(a['ニーズ本質']||'');
  sh.getRange(r, T.aiKubun).setValue(a['区分']||'');
  sh.getRange(r, T.aiApproach).setValue(a['推奨アプローチ']||'');
  const reps = (a['返信案']||[]).map((x,i)=>('【案'+(i+1)+'：'+(x['型']||'')+'】\n'+(x['文面']||''))).join('\n\n──────\n\n');
  sh.getRange(r, T.aiReplies).setValue(reps);

  // ①追客リストの温度に反映（確度を温度として持たせたい時の手掛かり）
  SpreadsheetApp.getUi().alert('判定しました（行'+r+'）。確度='+(a['確度']||'?')+' / 区分='+(a['区分']||'?')+' / 推奨='+(a['推奨アプローチ']||'?')+'\n返信案を'+(a['返信案']||[]).length+'パターン出力しました。');
}

function _aiAdvise(channel, inquiry, bukken) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('ANTHROPIC_API_KEY 未設定');
  const sys =
    'あなたは賃貸仲介トップセールス(羽鳥流)の判断脳。中核信条=売り込まない/嘘の緊急性NG/相手目線でGIVE。' +
    '反響の[流入経路]と[問合せ文面]から、相手のニーズの本質と確度を判定し、確度・状況に応じた' +
    '複数の返信アプローチを出す。\n' +
    '◆確度の読み: 紹介/電話/物件名指定=高、ポータル一斉=中〜低、「とりあえず/条件なし/一行/まだ先/深夜衝動」=低(微妙→確度を下げる)。' +
    '時期明記/内見したい/条件具体=本気度高。\n' +
    '◆ニーズ本質: 表面条件でなく「なぜ動くか」(通勤/同棲/在宅/現状不満)を推定。\n' +
    '◆アプローチ: A.ストレート提案(HOT:物件+初期費用+内見即提案+事実ベースの後押し) / ' +
    'B.目線ヒアリング(WARM・ニーズ曖昧:商品を先に出さず角度を変えた質問で本音を引き出す) / ' +
    'C.軽接触GIVE(COLD・微妙:売り込まずLINE登録/有益情報/焦らず長期枠) / D.絞り込み(条件が広い:MUST確認)。\n' +
    '◆区分: 初回反響=「新規開拓」(第一印象・スピード・ニーズ把握が最優先、いきなり商品はNG)。' +
    '既存client継続=「追客」(関係維持・新着提案・タイミングで背中押し、放置と追いすぎがNG)。文面から推定。\n' +
    '◆【菊池の声＝俺らしさ】新規の第一接触は"相手に興味を持たせる"が勝負＝人間にしかできない火種。' +
    '必ずこの声で書く: 正直/誠実(できない事は言わない・損も正直)、売り込まない/お願い感を出さない余裕、' +
    '丁寧だが堅すぎず親しみ、相手目線でGIVE先行、エッジ=「医療・土地家屋調査士・AIにも明るい不動産屋」という' +
    '他にない掛け算をさりげなく匂わせ"他と違う"と感じさせる、謙虚に学ぶ姿勢、盛らず数字で語る。\n' +
    '◆【新規の逆算設計】ゴール=「この人は他と違う、話してみたい」と思わせる。そこから逆算で' +
    '①フック(さりげないエッジ or 相手のニーズを言い当てる一言) ②安心(売り込まない・焦らせない) ' +
    '③菊池の正直で親しみある声 ④返しやすい問い1つ(ハードル低)。長文/いきなり商品/お願い感/嘘の緊急性はNG。' +
    '区分が「新規開拓」のときは返信案を必ずこの菊池の声＋逆算で書く。「追客」は関係前提で簡潔に。';
  const user =
    '[流入経路]'+channel+'\n[物件]'+bukken+'\n[問合せ文面]\n'+inquiry+'\n\n' +
    '次のJSONのみで返答:{"確度":"HOT|WARM|COLD","確度パーセント":0-100,"ニーズ本質":"…",' +
    '"区分":"新規開拓|追客","推奨アプローチ":"A.ストレート提案|B.目線ヒアリング|C.軽接触GIVE|D.絞り込み",' +
    '"返信案":[{"型":"推奨の型名","文面":"そのまま送れる本文"},{"型":"代替の型名","文面":"…"}]}' +
    '（返信案は推奨を先頭に2〜3案。微妙な相手ほどCを混ぜる。売り込まない。）';
  const body = {
    model: 'claude-sonnet-4-6',  // 必要に応じ変更可
    max_tokens: 2000,
    system: sys,
    messages: [{ role:'user', content: user }]
  };
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:'post', contentType:'application/json',
    headers:{ 'x-api-key':key, 'anthropic-version':'2023-06-01' },
    payload: JSON.stringify(body), muteHttpExceptions:true
  });
  if (res.getResponseCode() !== 200) throw new Error('API '+res.getResponseCode()+': '+res.getContentText().slice(0,300));
  const j = JSON.parse(res.getContentText());
  const text = (j.content && j.content[0] && j.content[0].text) || '';
  const mm = text.match(/\{[\s\S]*\}/);
  if (!mm) throw new Error('JSON抽出失敗: '+text.slice(0,200));
  return JSON.parse(mm[0]);
}
