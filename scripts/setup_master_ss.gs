/**
 * 00_書類処理マスターDB 初期化スクリプト
 *
 * 使い方:
 *   1. マスターSSを開く
 *   2. 拡張機能 → Apps Script
 *   3. このファイルの中身を全部貼り付け
 *   4. 関数 setupAll() を選択して実行
 *   5. 権限を許可
 *   6. 完了後、このスクリプト自体は残しておく（再実行で再初期化可能）
 *
 * 構成:
 *   - 受信ログ:       PDFを取り込んだ一覧。各PDFが1行。
 *   - 振分ルール:     書類種類×事業 → 保存先フォルダID の対応表（編集可能）
 *   - 勘定科目マスタ: 書類種類リスト＋勘定科目リスト＋自動推測ルール
 *   - 設定:           事業リスト・各種フォルダID・Gemini APIキー等
 */

function setupAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存シートを一旦消す（Sheet1 含む）
  const existingSheets = ss.getSheets();
  const tempSheet = ss.insertSheet('__temp__');
  existingSheets.forEach(s => ss.deleteSheet(s));

  setupSheetSetting(ss);
  setupSheetAccount(ss);
  setupSheetRule(ss);
  setupSheetLog(ss);

  ss.deleteSheet(tempSheet);

  // シート順を整える
  ss.setActiveSheet(ss.getSheetByName('受信ログ'));
  SpreadsheetApp.flush();
  Browser.msgBox('セットアップ完了', '4シート作成・プルダウン設定完了。\n\n次は「設定」シートのGemini APIキーと各種フォルダIDを埋めてください。', Browser.Buttons.OK);
}

function setupSheetSetting(ss) {
  const sh = ss.insertSheet('設定');
  const data = [
    ['キー', '値', '説明'],
    ['GEMINI_API_KEY', '', 'Gemini APIキー（https://aistudio.google.com/app/apikey）'],
    ['INBOX_FOLDER_ID', '1k80UlkueNAABF8rUg7HV_ExdPobnZX3C', '監視対象フォルダ(00_受信箱)のID'],
    ['DEFAULT_ARCHIVE_FOLDER_ID', '', '振分ルールに該当しなかった場合の保存先'],
    ['', '', ''],
    ['事業リスト', '', '（受信ログ H列のプルダウン選択肢）'],
    ['飯山満町', '', ''],
    ['不動産共通', '', ''],
    ['EC', '', ''],
    ['医療コンサル', '', ''],
    ['調査士事務所', '', ''],
    ['個人', '', ''],
    ['その他', '', ''],
  ];
  sh.getRange(1, 1, data.length, 3).setValues(data);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#E8F0FE');
  sh.getRange(6, 1, 1, 3).setFontWeight('bold').setBackground('#FFF3E0');
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 380);
  sh.setColumnWidth(3, 380);
}

function setupSheetAccount(ss) {
  const sh = ss.insertSheet('勘定科目マスタ');
  const data = [
    ['書類種類', '想定勘定科目', '備考'],
    ['請求書', '（取引内容による）', 'AIに取引内容から推測させる'],
    ['見積書', '（参考のみ）', '実発生時に確定'],
    ['契約書', '（参考のみ）', '内容により多岐'],
    ['領収書', '（取引内容による）', 'AIに取引内容から推測させる'],
    ['納品書', '（参考のみ）', ''],
    ['注文書', '（参考のみ）', ''],
    ['通知書', '（記録のみ）', '行政・税務関係'],
    ['図面', '（記録のみ）', '物件関連'],
    ['その他', '（要手動）', ''],
    ['', '', ''],
    ['勘定科目候補', '', '（受信ログ G列のプルダウン選択肢）'],
    ['修繕費', '', ''],
    ['外注費', '', ''],
    ['仕入高', '', ''],
    ['地代家賃', '', ''],
    ['支払手数料', '', ''],
    ['租税公課', '', ''],
    ['通信費', '', ''],
    ['消耗品費', '', ''],
    ['広告宣伝費', '', ''],
    ['交通費', '', ''],
    ['会議費', '', ''],
    ['新聞図書費', '', ''],
    ['水道光熱費', '', ''],
    ['保険料', '', ''],
    ['雑費', '', ''],
    ['建物', '', ''],
    ['土地', '', ''],
    ['建設仮勘定', '', ''],
    ['未払金', '', ''],
    ['売上高', '', ''],
    ['受取手数料', '', ''],
    ['要確認', '', '判別できない場合'],
  ];
  sh.getRange(1, 1, data.length, 3).setValues(data);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#E8F0FE');
  sh.getRange(12, 1, 1, 3).setFontWeight('bold').setBackground('#FFF3E0');
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 320);
}

function setupSheetRule(ss) {
  const sh = ss.insertSheet('振分ルール');
  const data = [
    ['書類種類', '事業', '保存先フォルダID', '保存先フォルダ名（メモ）'],
    ['請求書', '飯山満町', '', '飯山満町/05_請求書'],
    ['見積書', '飯山満町', '', '飯山満町/04_見積書'],
    ['契約書', '飯山満町', '', '飯山満町/02_契約書'],
    ['図面', '飯山満町', '', '飯山満町/03_図面'],
    ['請求書', '不動産共通', '', '不動産共通/請求書'],
    ['見積書', '不動産共通', '', '不動産共通/見積書'],
    ['請求書', 'EC', '', 'EC/経理'],
    ['請求書', '個人', '', '個人/家計'],
    ['その他', 'その他', '', '一旦ここに集約'],
  ];
  sh.getRange(1, 1, data.length, 4).setValues(data);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#E8F0FE');
  sh.setColumnWidth(1, 120);
  sh.setColumnWidth(2, 150);
  sh.setColumnWidth(3, 380);
  sh.setColumnWidth(4, 280);

  // 書類種類プルダウン（C列のソースを勘定科目マスタ A2:A10 から）
  const docTypeRange = ss.getSheetByName('勘定科目マスタ').getRange('A2:A10');
  const ruleDocTypeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(docTypeRange, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('A2:A1000').setDataValidation(ruleDocTypeValidation);

  // 事業プルダウン（B列のソースを設定 A7:A13 から）
  const bizRange = ss.getSheetByName('設定').getRange('A7:A13');
  const bizValidation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(bizRange, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('B2:B1000').setDataValidation(bizValidation);
}

function setupSheetLog(ss) {
  const sh = ss.insertSheet('受信ログ');
  const headers = [
    '取込日時', '元ファイルID', '元ファイル名', '書類種類', '日付', '金額',
    '取引先', '勘定科目', '物件/事業', '保存先フォルダID', 'リネーム後ファイル名',
    'ステータス', '次アクション', 'Tasks登録済み', 'メモ'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E8F0FE');
  sh.setFrozenRows(1);
  [140, 220, 220, 110, 110, 100, 160, 140, 130, 280, 260, 110, 240, 120, 220].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // 書類種類プルダウン（D列）
  const docTypeRange = ss.getSheetByName('勘定科目マスタ').getRange('A2:A10');
  sh.getRange('D2:D1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(docTypeRange, true).setAllowInvalid(false).build()
  );

  // 勘定科目プルダウン（H列）
  const accountRange = ss.getSheetByName('勘定科目マスタ').getRange('A13:A33');
  sh.getRange('H2:H1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(accountRange, true).setAllowInvalid(false).build()
  );

  // 事業プルダウン（I列）
  const bizRange = ss.getSheetByName('設定').getRange('A7:A13');
  sh.getRange('I2:I1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(bizRange, true).setAllowInvalid(false).build()
  );

  // ステータスプルダウン（L列）
  sh.getRange('L2:L1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['未処理', '解析済', 'リネーム済', '振分済', 'DB登録済', 'Tasks登録済', '完了', '要確認', 'エラー'], true).setAllowInvalid(false).build()
  );

  // J列（保存先フォルダID）を VLOOKUP で振分ルールから自動引き当て
  // =IFERROR(VLOOKUP(D2&I2, ARRAYFORMULA(振分ルール!A2:A1000 & 振分ルール!B2:B1000), 1, FALSE)... ではなく
  // ARRAYFORMULA + INDEX/MATCH 方式が安定
  const formula = '=IFERROR(INDEX(振分ルール!C:C, MATCH(D2 & "|" & I2, ARRAYFORMULA(振分ルール!A:A & "|" & 振分ルール!B:B), 0)), "")';
  sh.getRange('J2').setFormula(formula);
  // 下方向にコピー（1000行）
  sh.getRange('J2').copyTo(sh.getRange('J3:J1000'));
}
