/**
 * GAS① 00_受信箱 PDF検知 → Gemini 2.5 Flash 解析 → 受信ログ転記
 *
 * 抽出項目:
 *   - 書類情報: doc_type, date, amount, vendor, account, summary
 *   - 自動振分用: property_keyword, business_guess
 *
 * 使い方:
 *   1. 「設定」シートの GEMINI_API_KEY を埋める
 *   2. ingestNewPdfs() を一度手動実行して動作確認
 *   3. installFullTrigger() を1回だけ実行 → 10分ごとに自動巡回
 */

const GEMINI_MODEL = 'gemini-2.5-flash';
const MAX_FILES_PER_RUN = 10;

function ingestNewPdfs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('受信ログ');
  const settings = readSettings_(ss);

  if (!settings.GEMINI_API_KEY) {
    throw new Error('設定シートのGEMINI_API_KEYが空です。');
  }
  if (!settings.INBOX_FOLDER_ID) {
    throw new Error('設定シートのINBOX_FOLDER_IDが空です。');
  }

  const processedIds = collectProcessedIds_(logSheet);

  const rawInbox = String(settings.INBOX_FOLDER_ID || '').trim();
  const m = rawInbox.match(/folders\/([a-zA-Z0-9_-]+)/);
  const folderId = m ? m[1] : rawInbox;
  console.log('使用フォルダID: [' + folderId + '] (長さ=' + folderId.length + ')');
  if (!folderId) throw new Error('INBOX_FOLDER_IDが空です。設定シートB3を確認してください。');
  const folder = DriveApp.getFolderById(folderId);
  const targets = [];
  const it = folder.getFilesByType(MimeType.PDF);
  while (it.hasNext()) {
    const f = it.next();
    if (!processedIds.has(f.getId())) targets.push(f);
    if (targets.length >= MAX_FILES_PER_RUN) break;
  }

  if (targets.length === 0) {
    console.log('処理対象なし');
    return;
  }
  console.log('処理対象: ' + targets.length + '件');

  targets.forEach(file => {
    deleteErrorRowsForFile_(logSheet, file.getId());
    try {
      const result = analyzeWithGemini_(file, settings.GEMINI_API_KEY);
      appendLogRow_(logSheet, file, result, '解析済', '');
    } catch (e) {
      console.error('ファイル解析失敗: ' + file.getName() + ' / ' + e.message);
      appendLogRow_(logSheet, file, {}, 'エラー', e.message);
    }
  });
}

/** 「設定」シートをKey-Valueで読み込み（前後空白を自動除去） */
function readSettings_(ss) {
  const sh = ss.getSheetByName('設定');
  const values = sh.getRange('A2:B30').getValues();
  const obj = {};
  values.forEach(([k, v]) => {
    if (k && typeof k === 'string') {
      const key = k.trim();
      const val = (typeof v === 'string') ? v.trim() : v;
      obj[key] = val;
    }
  });
  return obj;
}

/** B列(元ファイルID) と L列(ステータス) からエラー行以外の処理済IDを収集 */
function collectProcessedIds_(logSheet) {
  const last = logSheet.getLastRow();
  const set = new Set();
  if (last < 2) return set;
  const data = logSheet.getRange(2, 2, last - 1, 11).getValues();
  data.forEach((row) => {
    const id = row[0];
    const status = row[10];
    // エラー行は処理済み扱いしない（リトライ対象に含める）
    if (id && status !== 'エラー') set.add(String(id));
  });
  return set;
}

/** 指定fileIdのエラー行を削除 */
function deleteErrorRowsForFile_(logSheet, fileId) {
  const last = logSheet.getLastRow();
  if (last < 2) return;
  const data = logSheet.getRange(2, 2, last - 1, 11).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    const id = data[i][0];
    const status = data[i][10];
    if (id === fileId && status === 'エラー') {
      logSheet.deleteRow(i + 2);
    }
  }
}

/** Gemini APIにPDFを投げて構造化JSONを取得 */
function analyzeWithGemini_(file, apiKey) {
  const pdfBytes = file.getBlob().getBytes();
  const base64 = Utilities.base64Encode(pdfBytes);

  const prompt = [
    'あなたは日本の経理書類解析の専門家です。添付PDFを読み取り、以下のJSONスキーマで応答してください。',
    '不明な項目は空文字"" で返してください。憶測で埋めず、判別不能はそのまま空に。',
    '',
    'スキーマ:',
    '{',
    '  "doc_type": "請求書|見積書|契約書|領収書|納品書|注文書|通知書|図面|その他",',
    '  "date": "YYYY-MM-DD",',
    '  "amount": 数値（税込総額。複数あれば合計）,',
    '  "vendor": "発行元・取引先名",',
    '  "account": "想定される勘定科目（修繕費/外注費/仕入高/地代家賃/支払手数料/租税公課/通信費/消耗品費/広告宣伝費/交通費/会議費/水道光熱費/保険料/雑費/建物/土地/建設仮勘定/未払金/売上高/受取手数料/要確認 から最も近いもの1つ）",',
    '  "summary": "1行50字以内の内容要約",',
    '  "property_keyword": "書類内に出てくる物件名・住所・エリア（飯山満町/鎌ケ谷市東鎌ケ谷/船橋市/東陽町/四谷 等）。複数あれば最も特定度が高いもの1つ。物件と関係ない書類は空文字",',
    '  "business_guess": "法人経費 | EC | 物件関連 | 不明",',
    '  "action_required": true/false（具体的アクションが必要か。requires_human_decisionと同じ判定でよい）,',
    '  "action_title": "20字以内の具体的タスク名。「対応」「確認」だけはNG、何をするか動詞で書く（例: 司法書士に必要書類を送付 / 工事見積を比較し発注先決定 / ○○へ請求額を支払い）",',
    '  "action_detail": "やるべき具体手順を3-5行で。証憑番号・連絡先・参照URLが書類にあれば含める",',
    '  "payment_deadline": "YYYY-MM-DD"（支払期限・契約期限・回答期限のいずれかが読み取れたら。なければ空文字）,',
    '  "requires_human_decision": true/false（菊池さん本人がタスクとして対応すべきか。下記基準で判定）',
    '}',
    '',
    'requires_human_decision の判断基準（タスク化要否。2026-05-24改定ルール）:',
    '- true（タスク化する）にすべき:',
    '    ・契約書 / 見積書 / 通知書 / 注文書 → 書類種類だけで無条件にtrue',
    '    ・請求書のうち「定例ではない」もの（スポット・新規取引先・単発の工事/外注等）→ 金額の大小は問わない',
    '- false（タスク化しない。記帳・保管のみ）にすべき:',
    '    ・領収書 / 納品書 / 図面 / その他',
    '    ・請求書のうち「定例」のもの = 下記の定例取引先からの請求書',
    '      【定例リスト】税理士 / リベ大オンラインサロン / GWS(Google Workspace) / 通信サービス系(携帯・回線・SaaS月額) / 家賃・地代 / 社会保険(年金機構・健康保険組合) / 水道光熱費',
    '- 「定例」とは毎月ほぼ同額で菊池さんの判断が不要な固定費。上記リストに該当しない請求書は、たとえ少額でもtrue（タスク化）にする。',
    '- false の場合はTasks登録されない（記帳のみで自動完了扱い）',
    '',
    'doc_type の判断基準:',
    '- 領収書/請求書で「工事」「修繕」→ 修繕費 or 建設仮勘定（高額・資産化）',
    '- 「家賃」「賃料」→ 地代家賃',
    '- 「印紙」「登録免許税」→ 租税公課',
    '- 「司法書士」「税理士」報酬→ 支払手数料',
    '- 不明確なら "要確認"',
    '',
    'business_guess の判断基準:',
    '- 取引先が「税理士」「年金機構」「健康保険組合」「労基」「税務署」「市町村役場（市県民税等）」→ 法人経費',
    '- 韓国コスメ・服飾・物販・Amazon・楽天・Shopify → EC',
    '- 物件名・所在地・不動産業者・建築会社・司法書士（登記関連）→ 物件関連',
    '- 判別不能 → 不明',
    '',
    'property_keyword の判断基準（厳密に）:',
    '- 「個別物件への参照」のみ抽出する：契約・取引・所有・登記・工事対象物件として明確に書類対象になっている物件名',
    '- 「飯山満町」「鎌ケ谷市東鎌ケ谷」「東陽町」「四谷」のように、物件フォルダ名と照合できる固有名詞',
    '- ❌ 単なる発行元住所（税理士事務所の住所等）は対象外',
    '- ❌ 不動産教材・営業ノウハウ資料内で一般的に地名が言及されているだけの場合は対象外（教材なら空文字）',
    '- ❌ ニュース記事内の地名言及も対象外',
    '- 判別がつかない一般地名なら空文字を返す。「請求書/契約書/見積書等が、その物件のための書類」と確信できる場合のみ抽出。'
  ].join('\n');

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inlineData: { mimeType: 'application/pdf', data: base64 } }
      ]
    }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json'
    }
  };

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(apiKey);

  const delays = [2000, 5000, 10000, 15000];
  let code, body;
  for (let attempt = 0; attempt <= delays.length; attempt++) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    code = res.getResponseCode();
    body = res.getContentText();
    if (code === 200) break;
    if (code !== 503 && code !== 429) break;
    if (attempt === delays.length) break;
    console.log('Gemini ' + code + ' → ' + delays[attempt] + 'ms 後にリトライ (試行' + (attempt + 1) + ')');
    Utilities.sleep(delays[attempt]);
  }
  if (code !== 200) {
    throw new Error('Gemini API HTTP ' + code + ': ' + body.substring(0, 500));
  }

  const json = JSON.parse(body);
  const text = json.candidates && json.candidates[0] && json.candidates[0].content
    && json.candidates[0].content.parts && json.candidates[0].content.parts[0].text;
  if (!text) throw new Error('Gemini応答にtextが含まれない: ' + body.substring(0, 300));

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    throw new Error('JSON parse失敗: ' + text.substring(0, 300));
  }
  return parsed;
}

/** マスタシートからプルダウン候補リストを取得（キャッシュ） */
let _docTypeListCache = null;
let _accountListCache = null;
function getDocTypeList_(ss) {
  if (_docTypeListCache) return _docTypeListCache;
  const sh = ss.getSheetByName('勘定科目マスタ');
  const vals = sh.getRange('A2:A10').getValues().flat().filter(v => v && typeof v === 'string').map(v => v.trim());
  _docTypeListCache = new Set(vals);
  return _docTypeListCache;
}
function getAccountList_(ss) {
  if (_accountListCache) return _accountListCache;
  const sh = ss.getSheetByName('勘定科目マスタ');
  const vals = sh.getRange('A13:A33').getValues().flat().filter(v => v && typeof v === 'string').map(v => v.trim());
  _accountListCache = new Set(vals);
  return _accountListCache;
}

/** Gemini応答をマスタリストに合わせて正規化 */
function normalizeGeminiResult_(ss, r) {
  const docTypes = getDocTypeList_(ss);
  const accounts = getAccountList_(ss);
  const docType = String(r.doc_type || '').trim();
  const account = String(r.account || '').trim();
  return {
    doc_type: docTypes.has(docType) ? docType : 'その他',
    date: String(r.date || '').trim(),
    amount: r.amount || '',
    vendor: String(r.vendor || '').trim(),
    account: accounts.has(account) ? account : '要確認',
    summary: String(r.summary || '').trim(),
    property_keyword: String(r.property_keyword || '').trim(),
    business_guess: String(r.business_guess || '').trim(),
    action_required: r.action_required === true || String(r.action_required).toLowerCase() === 'true',
    action_title: String(r.action_title || '').trim(),
    action_detail: String(r.action_detail || '').trim(),
    payment_deadline: String(r.payment_deadline || '').trim(),
    requires_human_decision: r.requires_human_decision === true || String(r.requires_human_decision).toLowerCase() === 'true'
  };
}

/** 受信ログに1行追加 */
function appendLogRow_(logSheet, file, r, status, errorNote) {
  const ss = logSheet.getParent();
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const nowStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');

  const norm = (status === '解析済') ? normalizeGeminiResult_(ss, r) : r;

  // I列の「物件/事業」: 物件キーワードがあれば優先、なければ事業推定
  const autoBiz = norm.property_keyword || norm.business_guess || '';

  // 列順: 取込日時/元ファイルID/元ファイル名/書類種類/日付/金額/取引先/勘定科目/物件事業/保存先ID/リネーム後/ステータス/次アクション/Tasks済/メモ
  // 「人手判断不要」と「具体アクション不要」のどちらも false なら（不要）扱い
  // 人手判断が必要な場合のみカレンダー登録対象とする
  let nextActionCell = '';
  if (status === '解析済') {
    if (norm.requires_human_decision && norm.action_title) {
      // タスク名に期限を埋め込む（カレンダー側で抽出）
      nextActionCell = norm.payment_deadline
        ? norm.action_title + '（期限:' + norm.payment_deadline + '）'
        : norm.action_title;
    } else {
      nextActionCell = '（不要）';
    }
  }

  const row = [
    nowStr,
    file.getId(),
    file.getName(),
    norm.doc_type || '',
    norm.date || '',
    norm.amount || '',
    norm.vendor || '',
    norm.account || '',
    autoBiz,                // I列: 自動判定（物件キーワード or 事業推定）
    '',                     // J列: 保存先ID（routeで埋める）
    '',                     // K列: リネーム後
    status,
    nextActionCell,         // M列: 次アクション
    'FALSE',
    errorNote || (norm.summary || '')
  ];

  const lastRow = logSheet.getLastRow();
  const targetRow = lastRow + 1;
  const before = row.slice(0, 9);
  const after = row.slice(10);
  logSheet.getRange(targetRow, 1, 1, 9).setValues([before]);
  logSheet.getRange(targetRow, 11, 1, after.length).setValues([after]);
}

/** 00_受信箱の全PDFを新ルールで強制再解析（既存ログ行を削除して入れ直す） */
function reprocessInboxFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('受信ログ');
  const settings = readSettings_(ss);
  if (!settings.GEMINI_API_KEY) throw new Error('設定シートのGEMINI_API_KEYが空です。');
  if (!settings.INBOX_FOLDER_ID) throw new Error('設定シートのINBOX_FOLDER_IDが空です。');

  const rawInbox = String(settings.INBOX_FOLDER_ID || '').trim();
  const m = rawInbox.match(/folders\/([a-zA-Z0-9_-]+)/);
  const folderId = m ? m[1] : rawInbox;
  const folder = DriveApp.getFolderById(folderId);

  const targets = [];
  const it = folder.getFilesByType(MimeType.PDF);
  while (it.hasNext()) targets.push(it.next());

  if (targets.length === 0) { console.log('受信箱にPDFなし'); return; }
  console.log('受信箱の強制再解析対象: ' + targets.length + '件');

  let ok = 0, ng = 0;
  targets.forEach(file => {
    deleteAllRowsForFile_(logSheet, file.getId());
    try {
      const result = analyzeWithGemini_(file, settings.GEMINI_API_KEY);
      appendLogRow_(logSheet, file, result, '解析済', '');
      ok++;
    } catch (e) {
      console.error('再解析失敗: ' + file.getName() + ' / ' + e.message);
      appendLogRow_(logSheet, file, {}, 'エラー', e.message);
      ng++;
    }
  });
  console.log('強制再解析完了: 成功=' + ok + ' / 失敗=' + ng);
}

/** 指定fileIdの全行を削除（ステータス問わず） */
function deleteAllRowsForFile_(logSheet, fileId) {
  const last = logSheet.getLastRow();
  if (last < 2) return;
  const data = logSheet.getRange(2, 2, last - 1, 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]) === String(fileId)) {
      logSheet.deleteRow(i + 2);
    }
  }
}

/** 動作確認用：1ファイルだけ手動指定で試す */
function testSingleFile(fileId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = readSettings_(ss);
  const file = DriveApp.getFileById(fileId);
  const result = analyzeWithGemini_(file, settings.GEMINI_API_KEY);
  console.log(JSON.stringify(result, null, 2));
  return result;
}

/* ===== 診断用 ===== */
function diag1_driveApp() {
  const root = DriveApp.getRootFolder();
  console.log('Root folder name: ' + root.getName());
}

function diag2_targetFolder() {
  const id = '1k80UlkueNAABF8rUg7HV_ExdPobnZX3C';
  try {
    const f = DriveApp.getFolderById(id);
    console.log('Folder name: ' + f.getName());
    console.log('Folder owner: ' + f.getOwner().getEmail());
    const it = f.getFiles();
    let count = 0;
    while (it.hasNext()) { it.next(); count++; }
    console.log('Files in folder: ' + count);
  } catch (e) {
    console.error('Error: ' + e.message);
  }
}
