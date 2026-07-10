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
