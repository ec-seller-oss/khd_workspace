/**
 * 💴 日繰り資金レーダー v3 【2026-06-12 根本解決版＝口座別日繰り予測】
 * ------------------------------------------------------------------
 * v2までの問題: 「3日後にAMEX引落です」と知らせるだけ＝結局その場で焦って振替。
 * v3の設計: A口座残高 × B1/B2出金予定 × B3入金予定 → D日繰り予報を毎朝自動計算し、
 *   「6/25 東京ベイTBが▲23,832不足 → 6/24(水)までにSBI銀行→TBへ20万振替」
 *   という"いつまでに・どこへ・いくら"の具体指示まで出す。
 *
 * 役割分担（確定）:
 *   MF        = 残高の事実（日々の正）
 *   このシート = 期日×口座×金額の未来予測と振替指示
 *   メール     = 毎朝5時のトリガー
 *
 * データの流れ:
 *   A表(口座残高・月初orMF確認時に手更新・更新日付き)
 *   + B1表(毎月の引落) + B2表(年数回の引落) + B3表(入金予定★確度=高のみ算入)
 *   → dailyCashAlert() が口座ごとに35日先まで日繰りシミュレーション
 *   → D表に書き込み＋不足があれば振替指示メール
 *
 * セットアップ（1回だけ）:
 *   1. replaceRadarWithV3() 実行（旧レーダーがあれば自動削除→v3生成）
 *   2. 🟡セル（引落日・口座・金額・A表残高）を実物で1回確認
 *   3. setupCashTrigger() 実行 → 毎朝5時に自動予測＋メール開始
 *
 * 運用（日々）:
 *   - 振替や大きい入出金をしたら A表の残高と更新日を直す（30秒）
 *   - 入金予定が確定したら B3 の確度を「高」にする（日付も実日に）
 *   - 残りは全自動
 */

var RADAR = {
  TAB: '00_操縦席',
  MAIL_TO: 'k.kenta0917@gmail.com',
  HORIZON_DAYS: 35,      // 日繰り予測の先読み日数
  ALERT_DAYS: 14,        // この日数以内の不足はメールで指示
  MARKER: '💴 日繰り資金レーダー',
  D_MARKER: 'D. 日繰り資金予報',
};

/** ============================================================
 * 旧レーダー（v1/v2）を自動削除 → v3を生成。これ1回でOK
 * ============================================================ */
function replaceRadarWithV3() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(RADAR.TAB);
  if (!sh) { SpreadsheetApp.getUi().alert('00_操縦席が見つかりません'); return; }

  var found = ss.createTextFinder(RADAR.MARKER).findNext();
  if (found && found.getSheet().getName() === RADAR.TAB) {
    var startRow = found.getRow();
    var endRow = startRow, emptyStreak = 0;
    for (var r = startRow + 1; r <= startRow + 200 && r <= sh.getMaxRows(); r++) {
      var vals = sh.getRange(r, 1, 1, 8).getValues()[0];
      var isEmpty = vals.every(function(v){ return v === '' || v === null; });
      if (isEmpty) { emptyStreak++; if (emptyStreak >= 5) break; }
      else { emptyStreak = 0; endRow = r; }
    }
    sh.deleteRows(startRow, endRow - startRow + 1);
  }
  buildCashRadar();
}

/** 互換用（前の案内名）。中身はv3置き換え */
function replaceRadarWithV2() { replaceRadarWithV3(); }

/** ============================================================
 * レーダー本体の生成
 * ============================================================ */
function buildCashRadar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(RADAR.TAB);
  if (!sh) { SpreadsheetApp.getUi().alert('00_操縦席タブが見つかりません'); return; }
  if (ss.createTextFinder(RADAR.MARKER).findNext()) {
    SpreadsheetApp.getUi().alert('旧レーダーがあります。replaceRadarWithV3() を実行してください（自動削除→再生成）'); return;
  }
  var r = sh.getLastRow() + 3;
  var C_HEAD='#1F3864', C_YEL='#FFF2CC', C_RED='#FCE4D6', C_GRAY='#F2F2F2', C_GREEN='#E2EFDA';

  sh.getRange(r,1).setValue(RADAR.MARKER + ' v3（毎朝5時に口座別日繰り予測→不足日と振替指示をメール）')
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(C_HEAD).setFontSize(12);
  sh.getRange(r,1,1,8).merge();

  // ============ A. 口座残高（更新日つき。振替したらここを直す＝30秒） ============
  sh.getRange(r+2,1).setValue('A. 口座残高 ★振替・大きい入出金をしたら残高と更新日を直す（MFの数字をそのまま転記）').setFontWeight('bold').setBackground(C_GRAY);
  sh.getRange(r+3,1,1,4).setValues([['口座（この名前がB1/B3の口座名と一致している必要あり）','残高','更新日','メモ']]).setFontWeight('bold').setBackground(C_HEAD).setFontColor('#FFFFFF');
  var bals = [
    ['住信SBI（個人）', 139320, '2026/06/01', '🔴AMEX10日＋住宅ローン27日＋PayPay27日の引落口座。月29万出る'],
    ['SBI銀行（個人）', 1224610, '2026/06/01', '振替の母艦（不足時はここから出す）'],
    ['東京ベイTB（個人事業）', 100000, '2026/06/11', '🔴6/25返済123,832'],
    ['ゆうちょ', 137132, '2026/06/01', '公庫5日'],
    ['楽天銀行RB1', 92400, '2026/06/01', '⚠️楽天カード27日の口座ならここ大穴'],
    ['法人TB', 1677563, '2026/06/01', '法人返済・社保'],
    ['城北信金（法人）', 281012, '2026/06/01', ''],
    ['麻梨奈SMBC（対象外）', '', '', '麻梨奈名義＝KHD資金繰りの監視対象外。NL/Amazonカード/管理費はここから引落（家計按分）'],
  ];
  sh.getRange(r+4,1,bals.length,4).setValues(bals).setBackground(C_YEL);
  sh.getRange(r+4,2,bals.length,1).setNumberFormat('#,##0"円"');

  // ============ B1. 毎月の引落マスタ ============
  var mr = r + 4 + bals.length + 2;
  sh.getRange(mr,1).setValue('B1. 毎月の引落マスタ ★🟡は実物で要確認。口座名はA表と同じ表記にする（予測の紐付けキー）').setFontWeight('bold').setBackground(C_GRAY);
  sh.getRange(mr+1,1,1,6).setValues([['毎月の日','項目','引落口座（A表と同名）','金額(目安)','確認','備考']]).setFontWeight('bold').setBackground(C_HEAD).setFontColor('#FFFFFF');
  var master = [
    [5,  '公庫返済（個人）', 'ゆうちょ', 21464, '🟡要確認', ''],
    [10, 'AMEXカード', '住信SBI（個人）', 150000, '🟡要確認', '🔴事故元。金額変動・明細D-5確認'],
    [10, 'NP掛け払い（タイミー等）', 'SBI銀行（個人）', 0, '🟡要確認', '🔴督促済。利用月のみ・振込なら口座=振込元'],
    [25, '東京ベイ返済（個人事業2本）', '東京ベイTB（個人事業）', 123832, '🟡要確認', '🔴毎月の本丸'],
    [26, '三井住友NLカード', '麻梨奈SMBC（対象外）', 121000, '対象外', '麻梨奈口座から引落＝予測対象外（家計按分で管理）'],
    [26, 'Amazonカード（三井住友発行）', '麻梨奈SMBC（対象外）', 30000, '対象外', '同上'],
    [27, '楽天カード（生活費+積立NISA）', '楽天銀行RB1', 280000, '🟡要確認', '生活18万+積立10万。口座どこか要確認'],
    [27, 'JALカード（法人切替済）', '法人TB', 30000, '🟡要確認', '5月法人移行・引落日/口座要確認'],
    [27, 'PayPay', '住信SBI（個人）', 10000, '🟡要確認', ''],
    [27, '住宅ローン', '住信SBI（個人）', 130668, '✅確認済', 'SBI引落を明細で確認済(2026-06-12)。AMEX10日と同口座＝27日前の残高に注意'],
    [27, 'SMBC管理費', '麻梨奈SMBC（対象外）', 22130, '対象外', '麻梨奈口座から引落'],
    [27, '税理士（法人+個人）', '法人TB', 65083, '🟡要確認', '個人分が別口座なら行を分ける'],
    [27, 'EC仕入・ツール支払', 'SBI銀行（個人）', 20731, '🟡要確認', '仕入カードどれか要特定'],
    [28, '朝日信金（2ヶ月毎）', '法人TB', 28847, '🟡要確認', '偶数月?頻度要確認。隔月なら該当月以外は0行を別管理'],
    [31, '社会保険（法人）', '法人TB', 22167, '🟡要確認', '末日'],
    [31, '法人返済（城北+朝日+大東京）', '城北信金（法人）', 81600, '🟡要確認', '日付バラなら行を分解'],
  ];
  sh.getRange(mr+2,1,master.length,6).setValues(master);
  sh.getRange(mr+2,4,master.length,1).setNumberFormat('#,##0"円"');
  for (var i=0;i<master.length;i++){
    if (String(master[i][5]).indexOf('🔴')===0) sh.getRange(mr+2+i,1,1,6).setBackground(C_RED);
  }

  // ============ B2. 年数回の引落 ============
  var ar = mr + 2 + master.length + 2;
  sh.getRange(ar,1).setValue('B2. 年数回の引落（月・日指定）★退職初年度の罠。通知書が来たら金額を入れる→予測に自動算入').setFontWeight('bold').setBackground(C_GRAY);
  sh.getRange(ar+1,1,1,6).setValues([['月','日','項目','引落口座（A表と同名）/納付書','金額(目安)','備考']]).setFontWeight('bold').setBackground(C_HEAD).setFontColor('#FFFFFF');
  var annual = [
    [6, 30, '住民税 普通徴収 第1期', 'SBI銀行（個人）', 0, '🔴アストン退職→特別徴収切れ。6月に通知書到着。納付書払いなら払う口座を書く'],
    [8, 31, '住民税 第2期', 'SBI銀行（個人）', 0, '🟡通知書で確定'],
    [10, 31, '住民税 第3期', 'SBI銀行（個人）', 0, '🟡'],
    [1, 31, '住民税 第4期', 'SBI銀行（個人）', 0, '🟡'],
    [6, 30, '固定資産税 第1期（北千住等）', 'SBI銀行（個人）', 0, '🔴賃貸保有=必ず来る。納税通知書で金額確定'],
    [9, 30, '固定資産税 第2期', 'SBI銀行（個人）', 0, '🟡'],
    [12, 28, '固定資産税 第3期', 'SBI銀行（個人）', 0, '🟡'],
    [2, 28, '固定資産税 第4期', 'SBI銀行（個人）', 0, '🟡'],
    [7, 31, '所得税 予定納税 第1期', 'SBI銀行（個人）', 0, '🟡該当なら6月通知。EC還付ポジションなら無しの可能性'],
    [11, 30, '所得税 予定納税 第2期', 'SBI銀行（個人）', 0, '🟡'],
  ];
  sh.getRange(ar+2,1,annual.length,6).setValues(annual);
  sh.getRange(ar+2,5,annual.length,1).setNumberFormat('#,##0"円"');
  sh.getRange(ar+2,1,1,6).setBackground(C_RED);
  sh.getRange(ar+6,1,1,6).setBackground(C_RED);

  // ============ B3. 入金予定（★確度=高 だけが予測に算入される） ============
  var br = ar + 2 + annual.length + 2;
  sh.getRange(br,1).setValue('B3. 入金予定 ★「確度」が 高 の行だけ予測に算入（保守設計）。確定したら 高 に変える＋日付を実日に').setFontWeight('bold').setBackground(C_GRAY);
  sh.getRange(br+1,1,1,6).setValues([['入金日','項目','入金口座（A表と同名）','金額','確度(高のみ算入)','備考']]).setFontWeight('bold').setBackground(C_HEAD).setFontColor('#FFFFFF');
  var inflows = [
    ['2026/07/15', '栄町 売却入金（隣地大工550万の手残り分）', '東京ベイTB（個人事業）', 1805000, '中', '登記律速。決済日が固まったら日付更新→確度=高'],
    ['2026/06/30', 'EC クーパン入金', 'SBI銀行（個人）', 0, '低', 'KYC解錠待ち。金額確定で更新'],
    ['2026/06/30', '辻堂 関連入金（100万当て込み分）', '東京ベイTB（個人事業）', 1000000, '中', 'CFに当て込み済の分。確定したら高に'],
    ['', '', '', '', '', '←03_売上予定で月が確定した案件は、ここに"日付"で転記する（月次→日次への変換は手動1行）'],
  ];
  sh.getRange(br+2,1,inflows.length,6).setValues(inflows);
  sh.getRange(br+2,4,inflows.length,1).setNumberFormat('#,##0"円"');
  sh.getRange(br+2,1,inflows.length,6).setBackground(C_GREEN);

  // ============ C. 現金化ルール（不足が大きい時のエスカレーション） ============
  var cr = br + 2 + inflows.length + 2;
  sh.getRange(cr,1).setValue('C. 現金化ルール（L0で足りない時に上の層へ。予測メールが自動でどの層か言う）').setFontWeight('bold').setBackground(C_GRAY);
  sh.getRange(cr+1,1,1,5).setValues([['層','原資','金額(概算)','現金化リードタイム','発動期限ルール']]).setFontWeight('bold').setBackground(C_HEAD).setFontColor('#FFFFFF');
  var rules = [
    ['L0','口座間振替（SBI母艦→各口座）', 1500000, '即日（平日15時まで）', '不足日の前営業日まで'],
    ['L1','楽天証券 株売却（高配当10銘柄）', 2716810, '約3営業日', '不足日の4営業日前までに売却判断'],
    ['L2','楽天証券 投信売却', 3860818, '約5営業日', '不足日の7営業日前まで'],
    ['L3','野村', 248468, '約1週間', '7営業日前まで'],
    ['L4','保険解約', 1100000, '2〜3週間', '最後の手段。21日前まで'],
  ];
  sh.getRange(cr+2,1,rules.length,5).setValues(rules);
  sh.getRange(cr+2,3,rules.length,1).setNumberFormat('#,##0"円"');
  sh.getRange(cr+2+rules.length,1).setValue('⚠️ L1/L2はNISA＝売ると枠が翌年まで戻らない。営業入金で買い戻す前提の橋渡し限定').setFontStyle('italic');

  // ============ D. 日繰り資金予報（自動更新エリア・手で触らない） ============
  var dr = cr + 2 + rules.length + 3;
  sh.getRange(dr,1).setValue(RADAR.D_MARKER + '（毎朝5時 自動更新・手で編集しない＝灰色運用）').setFontWeight('bold').setFontColor('#FFFFFF').setBackground(C_HEAD);
  sh.getRange(dr,1,1,8).merge();
  sh.getRange(dr+1,1).setValue('（dailyCashAlert 実行でここに口座別の不足予報と振替指示が書き込まれます）').setFontStyle('italic').setBackground(C_GRAY);
  sh.getRange(dr+1,1,1,8).merge();

  SpreadsheetApp.getUi().alert(
    '✅ レーダーv3生成完了\n\n' +
    '次にやること:\n' +
    '① B1の🟡（引落日・口座・金額）を実物確認\n' +
    '② dailyCashAlert() を1回手動実行 → D予報が書き込まれるのを確認\n' +
    '③ setupCashTrigger() 実行 → 毎朝5時自動化'
  );
}

/** ============================================================
 * 日繰り予測エンジン（毎朝5時）
 *  口座ごとに今日→35日先まで残高をシミュレーション
 *  → D表更新 ＋ 不足があれば「いつまでに・どこから・いくら」をメール
 * ============================================================ */
function dailyCashAlert() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(RADAR.TAB);
  if (!sh) return;
  var today = new Date(); today.setHours(0,0,0,0);

  // ---- マスタ読み込み ----
  var accounts = readTable_(ss, 'A. 口座残高', 4);          // [口座, 残高, 更新日, メモ]
  var monthly  = readTable_(ss, 'B1. 毎月の引落マスタ', 6); // [日, 項目, 口座, 金額, 確認, 備考]
  var annual   = readTable_(ss, 'B2. 年数回の引落', 6);     // [月, 日, 項目, 口座, 金額, 備考]
  var inflows  = readTable_(ss, 'B3. 入金予定', 6);         // [日付, 項目, 口座, 金額, 確度, 備考]
  if (!accounts.length || !monthly.length) return;

  // ---- 口座ごとにシミュレーション ----
  var horizon = new Date(today.getTime() + RADAR.HORIZON_DAYS * 86400000);
  var report = [];   // D表に書く行
  var alerts = [];   // メール指示
  var unknownAccts = {};

  // 口座名の正規化マッチ
  function norm(s){ return String(s||'').replace(/\s/g,''); }
  function findAcct(label) {
    var l = norm(label);
    if (!l) return null;
    for (var i=0;i<accounts.length;i++){
      var a = norm(accounts[i][0]);
      if (a === l || a.indexOf(l) === 0 || l.indexOf(a) === 0) return accounts[i][0];
    }
    return null;
  }

  // 出金・入金イベントを日付つきで展開
  function expandEvents(acctName) {
    var ev = []; // {date, amt(+入金/-出金), label}
    // B1 毎月: 更新日翌日〜horizon の各月に展開
    for (var m=0;m<monthly.length;m++){
      var day = Number(monthly[m][0]); if (!day) continue;
      if (String(monthly[m][2]).indexOf('対象外') !== -1) continue; // 麻梨奈口座等＝監視対象外
      var acct = findAcct(monthly[m][2]);
      if (!acct) { unknownAccts[String(monthly[m][2])] = true; continue; }
      if (acct !== acctName) continue;
      var amt = Number(monthly[m][3]) || 0; if (!amt) continue;
      for (var mm=-1; mm<3; mm++) {
        var base = new Date(today.getFullYear(), today.getMonth()+mm, 1);
        var last = new Date(base.getFullYear(), base.getMonth()+1, 0).getDate();
        var d = new Date(base.getFullYear(), base.getMonth(), Math.min(day, last));
        ev.push({date:d, amt:-amt, label:String(monthly[m][1])});
      }
    }
    // B2 年数回
    for (var a2=0;a2<annual.length;a2++){
      var mo = Number(annual[a2][0]), dy = Number(annual[a2][1])||28;
      if (!mo) continue;
      var acct2 = findAcct(annual[a2][3]); if (!acct2 || acct2 !== acctName) continue;
      var amt2 = Number(annual[a2][4]) || 0; if (!amt2) continue;
      var d2 = new Date(today.getFullYear(), mo-1, dy);
      if (d2 < new Date(today.getTime() - 32*86400000)) d2 = new Date(today.getFullYear()+1, mo-1, dy);
      ev.push({date:d2, amt:-amt2, label:String(annual[a2][2])+'（年数回）'});
    }
    // B3 入金（確度=高のみ）
    for (var f=0;f<inflows.length;f++){
      var dv = inflows[f][0]; if (!dv) continue;
      if (String(inflows[f][4]).indexOf('高') === -1) continue;
      var acct3 = findAcct(inflows[f][2]); if (!acct3 || acct3 !== acctName) continue;
      var d3 = (dv instanceof Date) ? new Date(dv) : new Date(String(dv));
      if (isNaN(d3)) continue;
      d3.setHours(0,0,0,0);
      ev.push({date:d3, amt:Number(inflows[f][3])||0, label:String(inflows[f][1])+'（入金）'});
    }
    ev.sort(function(x,y){ return x.date - y.date; });
    return ev;
  }

  for (var k=0;k<accounts.length;k++){
    var name = accounts[k][0];
    if (String(name).indexOf('対象外') !== -1) { report.push([name, '—', '—', '—', '対象外', '麻梨奈名義＝監視しない（家計按分で管理）']); continue; }
    var bal = accounts[k][1];
    if (bal === '' || bal === null) { report.push([name, '残高未入力', '', '', '', '🟡A表に残高を入れると予測対象になる']); continue; }
    bal = Number(bal) || 0;
    var asOf = accounts[k][2] ? new Date(accounts[k][2]) : today;
    asOf.setHours(0,0,0,0);

    var events = expandEvents(name);
    var running = bal, minBal = bal, minDate = null, firstShortDate = null, shortAmt = 0;
    for (var e=0;e<events.length;e++){
      var ed = events[e].date;
      if (ed <= asOf) continue;          // 更新日以前のイベントは残高に織込済とみなす
      if (ed > horizon) break;
      running += events[e].amt;
      if (ed >= today && running < minBal) { minBal = running; minDate = ed; }
      if (ed >= today && running < 0 && !firstShortDate) { firstShortDate = ed; shortAmt = -running; }
    }

    if (firstShortDate) {
      var deadline = prevBizDay_(firstShortDate);
      var rec = Math.max(Math.ceil((shortAmt + 10000)/10000)*10000, 50000); // 不足+1万を万単位切上げ・最低5万
      report.push([name, fmtYen_(bal), fmtYen_(minBal), fmtDate_(firstShortDate), fmtYen_(shortAmt)+' 不足',
                   '🔴 ' + fmtDate_(deadline) + 'までに SBI母艦→' + name + ' へ ' + fmtYen_(rec) + ' 振替（L0）']);
      var daysLeft = Math.floor((firstShortDate - today)/86400000);
      if (daysLeft <= RADAR.ALERT_DAYS) {
        alerts.push('🔴 ' + fmtDate_(firstShortDate) + ' ' + name + ' が ' + fmtYen_(shortAmt) + ' 不足見込み\n' +
                    '   → ' + fmtDate_(deadline) + '（前営業日）までに SBI銀行→' + name + ' へ ' + fmtYen_(rec) + ' 振替');
      }
    } else {
      report.push([name, fmtYen_(bal), fmtYen_(minBal), '—', '不足なし', '✅ 35日内は持つ見込み']);
    }
  }

  // 口座未特定の引落（紐付けできない＝予測から漏れている）を警告
  var unknownList = Object.keys(unknownAccts);
  if (unknownList.length) {
    report.push(['（未紐付け）', '', '', '', '', '⚠️ B1の口座名がA表と不一致で予測除外: ' + unknownList.join(' / ')]);
  }

  // ---- D表へ書き込み ----
  writeForecast_(ss, report);

  // ---- メール（不足がある時は指示、無い時も週1月曜は安全報告） ----
  if (alerts.length) {
    MailApp.sendEmail(RADAR.MAIL_TO,
      '🔴【資金レーダー】不足予報 ' + alerts.length + '口座 — 振替指示あり',
      alerts.join('\n\n') +
      '\n\n※残高の正はMF。振替を実行したらA表の残高と更新日を直す（30秒）。' +
      '\n※L0（振替）で足りない規模なら C表の期限ルール（株=4営業日前/投信=7営業日前）で売却判断。');
  } else if (today.getDay() === 1) {
    MailApp.sendEmail(RADAR.MAIL_TO,
      '✅【資金レーダー】35日内の不足なし（週次報告）',
      '全口座、向こう35日の引落をシミュレーションして不足なし。\nD表（操縦席）に口座別の最低残高を更新済み。');
  }
}

/** ---- D表クリア＆書き込み ---- */
function writeForecast_(ss, rows) {
  var found = ss.createTextFinder(RADAR.D_MARKER).findNext();
  if (!found) return;
  var sh = found.getSheet();
  var top = found.getRow() + 1;
  // 既存の予報エリアをクリア（最大30行）
  sh.getRange(top, 1, 30, 8).clearContent().setBackground(null);
  var headers = ['口座','現在残高(A表)','35日内 最低残高','不足日','不足額','指示'];
  sh.getRange(top, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
  if (rows.length) {
    sh.getRange(top+1, 1, rows.length, 6).setValues(rows);
    for (var i=0;i<rows.length;i++){
      var inst = String(rows[i][5]);
      var bg = inst.indexOf('🔴')===0 ? '#FCE4D6' : inst.indexOf('✅')===0 ? '#E2EFDA' : '#FFF2CC';
      sh.getRange(top+1+i, 1, 1, 6).setBackground(bg);
    }
  }
  sh.getRange(top+1+rows.length+1, 1).setValue('最終更新: ' + fmtDate_(new Date()) + '（毎朝5時自動）').setFontStyle('italic');
}

/** ---- テーブル読み込みヘルパ：見出しテキストを探し、ヘッダ行の下から空行まで読む ---- */
function readTable_(ss, headerText, cols) {
  var found = ss.createTextFinder(headerText).findNext();
  if (!found) return [];
  var sh = found.getSheet();
  var start = found.getRow() + 2; // 見出し行+列ヘッダ行の次
  var out = [];
  for (var r = start; r < start + 60; r++) {
    var vals = sh.getRange(r, 1, 1, cols).getValues()[0];
    var isEmpty = vals.every(function(v){ return v === '' || v === null; });
    if (isEmpty) break;
    out.push(vals);
  }
  return out;
}

/** ---- 前営業日（土日のみ考慮。祝日は手動で1日前倒し判断） ---- */
function prevBizDay_(date) {
  var d = new Date(date.getTime() - 86400000);
  while (d.getDay() === 0 || d.getDay() === 6) d = new Date(d.getTime() - 86400000);
  return d;
}

function fmtYen_(n) { return Utilities.formatString('%s円', Number(n).toLocaleString('ja-JP')); }
function fmtDate_(d) { return Utilities.formatDate(d, 'JST', 'M/d(E)'); }

/** ---- 毎朝5時トリガー ---- */
function setupCashTrigger() {
  var ts = ScriptApp.getProjectTriggers();
  for (var i=0;i<ts.length;i++) if (ts[i].getHandlerFunction()==='dailyCashAlert') ScriptApp.deleteTrigger(ts[i]);
  ScriptApp.newTrigger('dailyCashAlert').timeBased().everyDays(1).atHour(5).create();
  SpreadsheetApp.getUi().alert('✅ 毎朝5時の日繰り予測＋振替指示メール開始（' + RADAR.MAIL_TO + ' 宛）');
}
