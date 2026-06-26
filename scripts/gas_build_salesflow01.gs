/**
 * 01_営業フロー（現在地＋将来予測）タブ生成 【2026-06-11】
 * ------------------------------------------------------------------
 * 菊池指示：営業フローを見える化し、現在地と将来予測をセットで1タブに。
 * 設計判断（タブ増殖防止）：
 *  - 日々動くベース＝00_操縦席（今日の3件）＋02作業DB（朝予定/夜実績）に既にある→新設しない
 *  - このタブは「パイプライン現在地と月着地予測」専用。更新は週1（私が朝ブリーフで反映提案）
 *  - 既存タブ非破壊。同名あれば中止
 * 使い方：Apps Scriptに貼り → buildSalesflow01() 実行
 */
function buildSalesflow01() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var TAB = '01_営業フロー';
  if (ss.getSheetByName(TAB)) { SpreadsheetApp.getUi().alert(TAB + ' は既存です。削除してから再実行を。'); return; }
  var d = ss.insertSheet(TAB, 1);
  var HEAD='#1F3864', YEL='#FFF2CC', GRN='#E2EFDA', RED='#FCE4D6', GRY='#F2F2F2';

  d.getRange('A1').setValue('🔭 01_営業フロー ── パイプライン現在地と月着地予測（週1更新・日々の実働は00操縦席＋02作業DB）')
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(HEAD).setFontSize(12);
  d.getRange('A1:I1').merge();

  // ===== ファネル現在地 =====
  d.getRange('A3').setValue('▼ ファネル現在地（2026-06-11）').setFontWeight('bold').setBackground(GRY); d.getRange('A3:I3').merge();
  d.getRange(4,1,2,5).setValues([
    ['追客中(リード)','アポ確定','提案中','決済待ち','成約済(6月)'],
    ['15人リスト+的リスト35名','山崎先生(6/15週)','曾我/内山/門平(稲城)','栄町(持倉・171万)','0件 ← ここを月1.6件に'],
  ]);
  d.getRange(4,1,1,5).setFontWeight('bold').setBackground(HEAD).setFontColor('#FFFFFF');
  d.getRange(5,1,1,5).setBackground(YEL);

  // ===== パイプライン明細（確度加重） =====
  d.getRange('A7').setValue('▼ パイプライン明細（期待値順）　🟡=今週動かす').setFontWeight('bold').setBackground(GRY); d.getRange('A7:I7').merge();
  var hdr = ['案件','相手(経路)','ステージ','金額','確度','期待値','次アクション(誰が何を)','期限','着地月'];
  d.getRange(8,1,1,hdr.length).setValues([hdr]).setFontWeight('bold').setBackground(HEAD).setFontColor('#FFFFFF');
  var rows = [
    ['栄町売却','持倉さん(直)','決済待ち',1805000,0.95,'','司法書士岩崎の登記進捗→決済日確定','6月末','2026/07'],
    ['高松2丁目(仕入→出口)','トラステン阿部様','融資審査中',1200000,0.40,'','阿部様へ審査進捗ひと押し','今週','2026/09'],
    ['曾我先生 物件提案','曾我先生(福井経由)','提案中',440000,0.30,'','福井下書きの送信判断を締める(6/5繰越)','今日','2026/08'],
    ['京橋・山崎先生 AIコンサル','山崎先生(福井経由)','アポ確定',660000,0.30,'','★一人目：会食日程の確定LINE','今日','2026/08'],
    ['稲城建貸・門平先生','福井さん→門平','返信待ち',1100000,0.20,'','福井返信きたらGO文面2本送付(準備済)','今週','2026/09'],
    ['内山先生 物件提案','内山先生(福井経由)','提案中',440000,0.20,'','福井下書きの送信判断を締める','今週','2026/09'],
    ['横浜NPO 障害者GH','叔父TAW','座組検討',500000,0.20,'','R1 TAW契約・R2市の可否の2急所','6月','2026/10'],
    ['天野先生 AI-GIVE','にゃほにゃほ','COLD',0,0.00,'','謝罪+GIVEで入り直し(売らない)','余力日','—'],
  ];
  d.getRange(9,1,rows.length,hdr.length).setValues(rows);
  // 期待値=金額×確度（式）
  for (var i=0;i<rows.length;i++){ d.getRange(9+i,6).setFormula('=D'+(9+i)+'*E'+(9+i)); }
  d.getRange(9,4,rows.length,1).setNumberFormat('#,##0"円"');
  d.getRange(9,5,rows.length,1).setNumberFormat('0%');
  d.getRange(9,6,rows.length,1).setNumberFormat('#,##0"円"');
  d.getRange(9,1,2,hdr.length).setBackground(GRN);  // 栄町・高松
  d.getRange(11,1,2,hdr.length).setBackground(YEL); // 今日動かす2件
  var sumRow = 9 + rows.length;
  d.getRange(sumRow,5).setValue('期待値計').setFontWeight('bold');
  d.getRange(sumRow,6).setFormula('=SUM(F9:F'+(sumRow-1)+')').setNumberFormat('#,##0"円"').setFontWeight('bold').setBackground(GRN);

  // ===== 将来予測（月着地 vs 必要56万） =====
  var f = sumRow + 2;
  d.getRange(f,1).setValue('▼ 将来予測（期待値の月着地 vs 毎月の不足560,742円）').setFontWeight('bold').setBackground(GRY);
  d.getRange(f,1,1,9).merge();
  d.getRange(f+1,1,4,5).setValues([
    ['月','着地見込み(期待値)','不足56万との差','状態',''],
    ['2026/07','栄町171万','+115万 🟢','栄町で食える',''],
    ['2026/08','曾我13万+山崎20万=33万','▲23万 🔴','★今日の追客2.4件がここの飯のタネ',''],
    ['2026/09','高松48万+門平22万+内山9万=79万','+23万 🟡','6月の仕込みが効けば',''],
  ]);
  d.getRange(f+1,1,1,5).setFontWeight('bold').setBackground(HEAD).setFontColor('#FFFFFF');
  d.getRange(f+3,1,1,5).setBackground(RED);

  // ===== 運用ルール =====
  var u = f + 6;
  d.getRange(u,1).setValue('運用：①週1（月曜朝ブリーフ）で確度・ステージを秘書と更新 ②新規リードは02作業DBに入れる→ここはサマリのみ ③確度の定義: 0.95=契約済/0.4=審査中/0.3=提案済/0.2=接触中')
    .setFontStyle('italic');
  d.getRange(u,1,1,9).merge();

  d.setColumnWidth(1,200); d.setColumnWidth(2,160); d.setColumnWidth(7,300);
  d.getRange('A8:I'+sumRow).setBorder(true,true,true,true,true,true);
  SpreadsheetApp.getUi().alert('✅ 01_営業フロー を作成しました（期待値計と月着地予測つき）');
}
