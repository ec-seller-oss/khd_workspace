/**
 * GAS② 受信ログ解析済の行を自動振分
 *
 * 振分ロジック:
 *   1. property_keyword（I列）あり → Drive検索で物件フォルダ特定
 *      - 一致した物件フォルダの 01〜08 サブフォルダへ振分
 *      - 一致しない → 受信箱に残す（L列=要確認）
 *   2. property_keyword なし → business_guess（I列）で判定
 *      - 「法人経費」 → KHD法人経費フォルダ
 *      - 「EC」 → ECフォルダ
 *      - 「不明」 → 受信箱に残す（L列=要確認）
 *
 * 設定シート必須キー:
 *   - KHD_HOJIN_KEIHI_FOLDER_ID: 法人経費フォルダID
 *   - EC_FOLDER_ID: EC用フォルダID
 *   - PROPERTY_PARENT_FOLDER_IDS: 物件フォルダ検索範囲（空ならドライブ全体）
 */

// 書類種類 → 物件サブフォルダ番号のマッピング
const SUBFOLDER_MAP = {
  '請求書': '06_領収書',
  '領収書': '06_領収書',
  '見積書': '05_工事',
  '契約書': '04_仲介',
  '納品書': '05_工事',
  '注文書': '05_工事',
  '通知書': '01_概要',
  '図面': '01_概要',
  'その他': '01_概要'
};

function routeAndRename() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('受信ログ');
  const settings = readSettings_(ss);
  const businessMap = readBusinessMap_(ss);

  const last = logSheet.getLastRow();
  if (last < 2) { console.log('受信ログが空'); return; }

  const data = logSheet.getRange(2, 1, last - 1, 15).getValues();

  let processed = 0;
  let needCheck = 0;
  let errored = 0;

  data.forEach((row, idx) => {
    const rowNum = idx + 2;
    const fileId = row[1];
    const docType = row[3];
    const dateStr = row[4];
    const amount = row[5];
    const vendor = row[6];
    const biz = String(row[8] || '').trim(); // I列: property_keyword or business_guess
    const status = row[11];

    if (status !== '解析済' && status !== '要確認') return;
    if (!fileId) {
      logSheet.getRange(rowNum, 12).setValue('エラー');
      logSheet.getRange(rowNum, 15).setValue('ファイルIDが空');
      errored++;
      return;
    }

    try {
      const file = DriveApp.getFileById(fileId);
      const newName = buildFileName_(docType, dateStr, amount, vendor, file.getName());

      // 振分先決定
      const dest = decideDestination_(biz, docType, settings, businessMap);

      if (dest) {
        // 振分先確定 → リネーム＋移動
        const folder = DriveApp.getFolderById(dest.folderId);
        const finalName = ensureUniqueName_(folder, newName);
        file.setName(finalName);
        moveFileTo_(file, folder);
        const folderUrl = 'https://drive.google.com/drive/folders/' + dest.folderId;
        const linkLabel = (dest.label || 'フォルダ').replace(/"/g, '');
        logSheet.getRange(rowNum, 10).setFormula('=HYPERLINK("' + folderUrl + '","' + linkLabel + '")');
        // ファイル名も Drive で開けるよう HYPERLINK 化
        const fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
        const fileLabel = finalName.replace(/"/g, '');
        logSheet.getRange(rowNum, 11).setFormula('=HYPERLINK("' + fileUrl + '","' + fileLabel + '")');
        logSheet.getRange(rowNum, 12).setValue('振分済');
        logSheet.getRange(rowNum, 15).setValue('[振分先] ' + (dest.label || ''));
        processed++;
      } else {
        // 振分不能 → 受信箱に残す＋リネームのみ実施
        const inboxId = extractFolderId_(settings.INBOX_FOLDER_ID);
        const inbox = DriveApp.getFolderById(inboxId);
        const finalName = ensureUniqueName_(inbox, newName);
        file.setName(finalName);
        const inboxUrl = 'https://drive.google.com/drive/folders/' + inboxId;
        logSheet.getRange(rowNum, 10).setFormula('=HYPERLINK("' + inboxUrl + '","00_受信箱")');
        const fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
        const fileLabel = finalName.replace(/"/g, '');
        logSheet.getRange(rowNum, 11).setFormula('=HYPERLINK("' + fileUrl + '","' + fileLabel + '")');
        logSheet.getRange(rowNum, 12).setValue('要確認');
        logSheet.getRange(rowNum, 15).setValue('[受信箱に残留・リネーム済] 判別:' + biz);
        needCheck++;
      }
    } catch (e) {
      console.error('行' + rowNum + ' 失敗: ' + e.message);
      logSheet.getRange(rowNum, 15).setValue('[振分失敗] ' + e.message);
      logSheet.getRange(rowNum, 12).setValue('エラー');
      errored++;
    }
  });

  console.log('振分完了: 振分済=' + processed + ' / 要確認=' + needCheck + ' / エラー=' + errored);
}

/**
 * 振分先を決定
 * @returns {{folderId: string, label: string} | null}
 */
function decideDestination_(biz, docType, settings, businessMap) {
  if (!biz) return null;

  // ① 設定シートの「事業マッピング」セクションを優先参照
  if (businessMap.has(biz)) {
    const url = businessMap.get(biz);
    if (!url) return null; // 振分先空 = 受信箱残留
    const id = extractFolderId_(url);
    if (id) return { folderId: id, label: biz };
    return null;
  }

  // ② マッピング外＝物件名扱い → Drive検索
  const propFolder = findPropertyFolder_(biz);
  if (propFolder) {
    const subName = SUBFOLDER_MAP[docType] || '01_概要';
    const subFolder = getOrCreateSubFolder_(propFolder, subName);
    return { folderId: subFolder.getId(), label: propFolder.getName() + '/' + subName };
  }
  return null;
}

/** 設定シートの「事業マッピング」セクションを読む */
function readBusinessMap_(ss) {
  const sh = ss.getSheetByName('設定');
  const all = sh.getRange('A1:C100').getValues();
  const map = new Map();
  let inSection = false;
  for (let i = 0; i < all.length; i++) {
    const a = String(all[i][0] || '').trim();
    const b = String(all[i][1] || '').trim();
    if (a === '事業マッピング' || a === '事業名') { inSection = true; continue; }
    if (!inSection) continue;
    if (!a) break; // 空行で終了
    map.set(a, b); // bが空文字でもMapには登録（=受信箱残留扱い）
  }
  return map;
}

/** URL or ID どちらが入っていても フォルダID を抽出 */
function extractFolderId_(v) {
  const s = String(v || '').trim();
  const m = s.match(/folders\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : s;
}

/** 受信ログのデータ入力規則を「警告のみ」に緩和（AI書き込みでrejectされないように） */
function relaxValidations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('受信ログ');
  ['D2:D1000', 'H2:H1000', 'I2:I1000'].forEach(rangeA1 => {
    const range = sh.getRange(rangeA1);
    const current = range.getDataValidation();
    if (!current) return;
    const newRule = current.copy().setAllowInvalid(true).build();
    range.setDataValidation(newRule);
  });
  console.log('D/H/I列のデータ入力規則を「警告のみ」に変更しました');
}

/** 物件キーワードでDrive検索して NNN_ プレフィックス付きフォルダを探す */
function findPropertyFolder_(keyword) {
  if (!keyword) return null;
  // 1) 「NNN_xxxxx_キーワード」形式の物件フォルダを優先
  const escaped = keyword.replace(/'/g, "\\'");
  let query = "mimeType = 'application/vnd.google-apps.folder' and title contains '" + escaped + "' and trashed = false";
  const files = DriveApp.searchFolders(query);
  while (files.hasNext()) {
    const f = files.next();
    if (/^\d{1,4}_/.test(f.getName())) return f; // NNN_ プレフィックスあり = 物件フォルダ
  }
  return null;
}

/** 物件フォルダ内のサブフォルダを取得（無ければ作成） */
function getOrCreateSubFolder_(propertyFolder, subFolderName) {
  const it = propertyFolder.getFoldersByName(subFolderName);
  if (it.hasNext()) return it.next();
  return propertyFolder.createFolder(subFolderName);
}

/** ファイル名生成: YYMMDD_書類種類（取引先_金額円）.pdf */
function buildFileName_(docType, dateStr, amount, vendor, originalName) {
  let ymd = '';
  if (dateStr instanceof Date && !isNaN(dateStr.getTime())) {
    ymd = Utilities.formatDate(dateStr, Session.getScriptTimeZone(), 'yyMMdd');
  } else if (dateStr) {
    const m = String(dateStr).match(/(\d{4})-(\d{2})-(\d{2})/);
    if (m) ymd = m[1].slice(2) + m[2] + m[3];
  }
  if (!ymd) {
    const today = new Date();
    ymd = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyMMdd');
  }

  const type = sanitize_(docType || '書類');
  const vend = sanitize_(vendor || '');
  const amt = amount ? String(Math.round(Number(amount))).replace(/\B(?=(\d{3})+(?!\d))/g, ',') : '';

  const inner = [vend, amt ? amt + '円' : ''].filter(Boolean).join('_');
  const paren = inner ? '（' + inner + '）' : '';

  const ext = (originalName && originalName.match(/\.[a-zA-Z0-9]+$/)) ? originalName.match(/\.[a-zA-Z0-9]+$/)[0] : '.pdf';

  return ymd + '_' + type + paren + ext;
}

function sanitize_(s) {
  return String(s || '').replace(/[\/\\:\*\?"<>\|]/g, '').trim();
}

function ensureUniqueName_(folder, name) {
  const base = name.replace(/\.[a-zA-Z0-9]+$/, '');
  const ext = (name.match(/\.[a-zA-Z0-9]+$/) || ['.pdf'])[0];
  let candidate = name;
  let n = 2;
  while (true) {
    const it = folder.getFilesByName(candidate);
    if (!it.hasNext()) return candidate;
    candidate = base + '_' + n + ext;
    n++;
    if (n > 100) return candidate;
  }
}

function moveFileTo_(file, targetFolder) {
  const parents = file.getParents();
  targetFolder.addFile(file);
  while (parents.hasNext()) {
    const p = parents.next();
    if (p.getId() !== targetFolder.getId()) {
      p.removeFile(file);
    }
  }
}

/** Ingest + Routeをまとめて実行 */
function processAllPdfs() {
  ingestNewPdfs();
  routeAndRename();
}

function installFullTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'ingestNewPdfs' || fn === 'processAllPdfs' || fn === 'routeAndRename') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processAllPdfs')
    .timeBased()
    .everyMinutes(10)
    .create();
  Browser.msgBox('完了', '10分ごとに ingest+route を自動実行するトリガーを設定しました。', Browser.Buttons.OK);
}

function uninstallAllTriggers() {
  let count = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'ingestNewPdfs' || fn === 'processAllPdfs' || fn === 'routeAndRename') {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  });
  Browser.msgBox('完了', count + '件のトリガーを削除しました。', Browser.Buttons.OK);
}

/** 設定シートに「事業マッピング」セクションをセットアップ */
function setupBusinessMapSection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('設定');
  // 既存セクション確認
  const all = sh.getRange('A1:A100').getValues().flat();
  if (all.some(v => String(v).trim() === '事業マッピング')) {
    console.log('既に事業マッピングセクション存在。スキップ');
    return;
  }
  const lastRow = sh.getLastRow();
  const data = [
    ['', '', ''],
    ['事業マッピング', '', '（事業名→振分先フォルダURL。空欄にすると受信箱残留）'],
    ['事業名', '振分先URL', 'メモ'],
    ['KHD法人', 'https://drive.google.com/drive/folders/15YzBKPksFV5oq50w3Vi0Tko-2SHjnUKz', '法人経費（社保・税理士・諸経費）'],
    ['法人経費', 'https://drive.google.com/drive/folders/15YzBKPksFV5oq50w3Vi0Tko-2SHjnUKz', '同上（AI判定エイリアス）'],
    ['EC', 'https://drive.google.com/drive/folders/17VbU2ZOx5m_2YgZa64MuicENV4QIgH7q', '韓国輸出・物販'],
    ['不動産共通', '', '受信箱に残留（後でフォルダ指定したくなったらURLを入れる）'],
    ['物件関連', '', '受信箱に残留'],
    ['不明', '', '受信箱に残留'],
    ['その他', '', '受信箱に残留']
  ];
  sh.getRange(lastRow + 1, 1, data.length, 3).setValues(data);
  // ヘッダーをハイライト
  sh.getRange(lastRow + 2, 1, 1, 3).setBackground('#FFF3E0').setFontWeight('bold');
  sh.getRange(lastRow + 3, 1, 1, 3).setBackground('#E8F0FE').setFontWeight('bold');
  console.log('事業マッピングセクションを追加しました');
}

/** ステータス問わず全行を再振分（'エラー'と'振分済'以外も全部対象） */
function routeAndRenameAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('受信ログ');
  const settings = readSettings_(ss);
  const businessMap = readBusinessMap_(ss);

  const last = logSheet.getLastRow();
  if (last < 2) { console.log('受信ログが空'); return; }
  const data = logSheet.getRange(2, 1, last - 1, 15).getValues();

  let processed = 0;
  let needCheck = 0;
  let errored = 0;
  let skipped = 0;

  data.forEach((row, idx) => {
    const rowNum = idx + 2;
    const fileId = row[1];
    const docType = row[3];
    const dateStr = row[4];
    const amount = row[5];
    const vendor = row[6];
    const biz = String(row[8] || '').trim();
    const status = row[11];

    // 'エラー'は対象外。それ以外は全部対象（'振分済'も再振分）
    if (status === 'エラー') { skipped++; return; }
    if (!fileId) { skipped++; return; }

    try {
      const file = DriveApp.getFileById(fileId);
      const newName = buildFileName_(docType, dateStr, amount, vendor, file.getName());
      const dest = decideDestination_(biz, docType, settings, businessMap);

      if (dest) {
        const folder = DriveApp.getFolderById(dest.folderId);
        const finalName = ensureUniqueName_(folder, newName);
        file.setName(finalName);
        moveFileTo_(file, folder);
        const folderUrl = 'https://drive.google.com/drive/folders/' + dest.folderId;
        const linkLabel = (dest.label || 'フォルダ').replace(/"/g, '');
        logSheet.getRange(rowNum, 10).setFormula('=HYPERLINK("' + folderUrl + '","' + linkLabel + '")');
        const fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
        const fileLabel = finalName.replace(/"/g, '');
        logSheet.getRange(rowNum, 11).setFormula('=HYPERLINK("' + fileUrl + '","' + fileLabel + '")');
        logSheet.getRange(rowNum, 12).setValue('振分済');
        logSheet.getRange(rowNum, 15).setValue('[強制振分] ' + (dest.label || ''));
        processed++;
      } else {
        const inboxId = extractFolderId_(settings.INBOX_FOLDER_ID);
        const inbox = DriveApp.getFolderById(inboxId);
        const finalName = ensureUniqueName_(inbox, newName);
        file.setName(finalName);
        const inboxUrl = 'https://drive.google.com/drive/folders/' + inboxId;
        logSheet.getRange(rowNum, 10).setFormula('=HYPERLINK("' + inboxUrl + '","00_受信箱")');
        const fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
        const fileLabel = finalName.replace(/"/g, '');
        logSheet.getRange(rowNum, 11).setFormula('=HYPERLINK("' + fileUrl + '","' + fileLabel + '")');
        logSheet.getRange(rowNum, 12).setValue('要確認');
        logSheet.getRange(rowNum, 15).setValue('[受信箱残留・リネーム済] 判別:' + biz);
        needCheck++;
      }
    } catch (e) {
      console.error('行' + rowNum + ' 失敗: ' + e.message);
      logSheet.getRange(rowNum, 15).setValue('[振分失敗] ' + e.message);
      logSheet.getRange(rowNum, 12).setValue('エラー');
      errored++;
    }
  });

  console.log('強制振分完了: 振分済=' + processed + ' / 要確認=' + needCheck + ' / エラー=' + errored + ' / スキップ=' + skipped);
}

/** 設定シートに新キーを追加（既存設定の末尾に追記） */
function addSettingsKeys() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('設定');
  const data = sh.getRange('A1:A30').getValues().flat();
  const existing = new Set(data.filter(v => v).map(v => String(v).trim()));

  // フォルダはURL形式で書き込む（クリックして飛べるように）
  const additions = [
    ['KHD_HOJIN_KEIHI_FOLDER_ID', 'https://drive.google.com/drive/folders/15YzBKPksFV5oq50w3Vi0Tko-2SHjnUKz', 'KHD法人経費フォルダ（社保・税理士・諸経費）'],
    ['EC_FOLDER_ID', 'https://drive.google.com/drive/folders/17VbU2ZOx5m_2YgZa64MuicENV4QIgH7q', 'EC（個人物販）フォルダ']
  ];

  let appended = 0;
  let updated = 0;
  additions.forEach(([k, v, desc]) => {
    if (existing.has(k)) {
      // 既存なら値をURLに更新する
      const data = sh.getRange('A1:B30').getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() === k) {
          sh.getRange(i + 1, 2).setValue(v);
          updated++;
          break;
        }
      }
      return;
    }
    const lastRow = sh.getLastRow();
    sh.getRange(lastRow + 1, 1, 1, 3).setValues([[k, v, desc]]);
    appended++;
  });
  console.log('追加: ' + appended + '件, URL更新: ' + updated + '件');
  Browser.msgBox('完了', appended + '件の設定キーを追加しました。', Browser.Buttons.OK);
}
