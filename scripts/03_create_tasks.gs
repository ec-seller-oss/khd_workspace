/**
 * GAS③ 受信ログ「振分済」行から具体的アクションを生成し
 *   - Google Tasks（@default リスト）にタスクとして登録
 *   - 期限がある書類は task.due に設定
 *   - 書類情報（取引先・金額・Driveリンク）は notes に格納
 *
 * 前提:
 *   appsscript.json の dependencies.enabledAdvancedServices に Tasks を追加
 *   oauthScopes に https://www.googleapis.com/auth/tasks を追加
 *
 * ルール:
 *   - 受信ログ L列=「振分済」 かつ N列=FALSE の行を対象
 *   - M列が「（不要）」または空 → 登録不要としてN列=TRUE
 *   - action_title はM列に格納済み（期限あれば「（期限:YYYY-MM-DD）」を末尾に埋め込み）
 */

const TASK_LIST_ID = '@default';  // メインタスクリスト

function createTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('受信ログ');

  const last = logSheet.getLastRow();
  if (last < 2) { console.log('受信ログが空'); return; }
  const data = logSheet.getRange(2, 1, last - 1, 15).getValues();

  let registered = 0;
  let skipped = 0;
  let errored = 0;

  data.forEach((row, idx) => {
    const rowNum = idx + 2;
    const ingestCell = row[0];
    const fileId = row[1];
    const fileName = row[2];
    const docType = String(row[3] || '').trim();
    const dateCell = row[4];
    const amount = row[5];
    const vendor = String(row[6] || '').trim();
    const status = row[11];
    const nextActionCell = String(row[12] || '').trim();
    const tasksDone = row[13];

    if (status !== '振分済' && status !== '要確認') return;
    if (tasksDone === true || String(tasksDone).toUpperCase() === 'TRUE') return;

    try {
      // M列が「（不要）」または空白なら登録不要
      if (nextActionCell === '（不要）' || !nextActionCell) {
        logSheet.getRange(rowNum, 14).setValue(true);
        skipped++;
        return;
      }

      // 期限抽出: M列に「（期限:YYYY-MM-DD）」が埋め込まれていれば取り出す
      const deadlineMatch = nextActionCell.match(/期限:(\d{4}-\d{2}-\d{2})/);
      const rawTitle = nextActionCell.replace(/（期限:\d{4}-\d{2}-\d{2}）/, '').trim();

      // 期限決定: 明示期限があればそれ、なければ取込日+3営業日（放置防止デフォルト）
      let dueDate;
      if (deadlineMatch) {
        dueDate = new Date(deadlineMatch[1] + 'T00:00:00.000Z');
      } else {
        const baseDate = (ingestCell instanceof Date && !isNaN(ingestCell.getTime()))
          ? new Date(ingestCell.getTime()) : new Date();
        dueDate = addBusinessDays_(baseDate, 3);
      }

      // 優先度マーク: 🔴=期限3日以内 or 契約書/通知書、🟡=その他
      const daysUntilDue = Math.ceil((dueDate.getTime() - new Date().getTime()) / (1000 * 60 * 60 * 24));
      const isUrgent = daysUntilDue <= 3 || docType === '契約書' || docType === '通知書';
      const mark = isUrgent ? '🔴' : '🟡';
      const title = mark + ' ' + rawTitle;

      // notes欄: 書類情報 + リンク
      const linkUrl = fileId ? 'https://drive.google.com/file/d/' + fileId + '/view' : '';
      const notesParts = [
        '【書類】' + (fileName || ''),
        linkUrl ? '【リンク】' + linkUrl : '',
        '【書類種類】' + (docType || ''),
        '【取引先】' + (vendor || '不明'),
        amount ? '【金額】' + Number(amount).toLocaleString() + '円' : '',
        '【書類日付】' + (formatDateCell_(dateCell) || '不明'),
        '',
        '— 自動生成 (受信ログ行' + rowNum + ') —'
      ].filter(Boolean);

      const task = {
        title: title,
        notes: notesParts.join('\n'),
        status: 'needsAction',
        due: dueDate.toISOString()
      };

      const inserted = Tasks.Tasks.insert(task, TASK_LIST_ID);

      logSheet.getRange(rowNum, 14).setValue(true);
      const dueLabel = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        + (deadlineMatch ? '' : '(自動+3営業日)');
      logSheet.getRange(rowNum, 15).setValue(String(row[14] || '') + ' [Tasks登録] ' + mark + ' 期限' + dueLabel);
      registered++;
    } catch (e) {
      console.error('行' + rowNum + ' Tasks登録失敗: ' + e.message);
      logSheet.getRange(rowNum, 15).setValue(String(row[14] || '') + ' [Tasks失敗] ' + e.message);
      errored++;
    }
  });

  console.log('Tasks登録完了: 登録=' + registered + ' / スキップ=' + skipped + ' / エラー=' + errored);
}

/** 起点日からN営業日後を返す（土日スキップ） */
function addBusinessDays_(from, n) {
  const d = new Date(from.getTime());
  let added = 0;
  while (added < n) {
    d.setDate(d.getDate() + 1);
    const dow = d.getDay();
    if (dow !== 0 && dow !== 6) added++;  // 0=日,6=土をスキップ
  }
  d.setHours(0, 0, 0, 0);
  return d;
}

/** 日付セル整形 */
function formatDateCell_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (!v) return '';
  const s = String(v).trim();
  const m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[0] : s;
}

/** Ingest + Route + Tasks をまとめて実行（フルパイプライン） */
function processAllPdfsAndTasks() {
  ingestNewPdfs();
  routeAndRename();
  createTasks();
}

/** トリガーを「フルパイプライン版」に切り替え */
function installFullPipelineTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'ingestNewPdfs' || fn === 'processAllPdfs' || fn === 'routeAndRename'
        || fn === 'createCalendarEvents' || fn === 'createTasks' || fn === 'processAllPdfsAndTasks') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processAllPdfsAndTasks')
    .timeBased()
    .everyMinutes(10)
    .create();
  Browser.msgBox('完了', '10分ごとに ingest+route+tasks のフルパイプラインを自動実行します。', Browser.Buttons.OK);
}

/** 動作確認: 受信ログの全「振分済」行をTasksに登録（手動実行用） */
function testCreateTasks() {
  createTasks();
}

/** テスト用: ダミータスクを1件強制的にマイタスクに書き込む（動作証明） */
function testInsertTask() {
  const task = {
    title: '【テスト】Day5動作確認 2026-05-23',
    notes: '03_create_tasks.gs の動作証明用テストタスク。\nこのタスクが「マイタスク」に表示されればGoogle Tasks書き込み成功。\n— 自動生成 (testInsertTask) —',
    status: 'needsAction',
    due: new Date('2026-05-25T00:00:00.000Z').toISOString()
  };
  const result = Tasks.Tasks.insert(task, '@default');
  console.log('✅ Tasks書き込み成功');
  console.log('タスクID: ' + result.id);
  console.log('タイトル: ' + result.title);
  console.log('期限: ' + (result.due || 'なし'));
  console.log('URL: https://tasks.google.com/');
  console.log('→ Google Tasks(https://tasks.google.com/)を開いて「マイタスク」に表示されているか確認してください');
}

/** Google Tasks のリスト一覧を確認（デバッグ用） */
function listTaskLists() {
  const lists = Tasks.Tasklists.list();
  if (lists.items) {
    lists.items.forEach(l => console.log(l.id + ' / ' + l.title));
  } else {
    console.log('タスクリストなし');
  }
}
