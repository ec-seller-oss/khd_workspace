/**
 * 賃貸反響 スプシ内ガイド書き込み（2026-06-05）
 * 既存の統合GASとは別の「新しいスクリプトファイル」に貼って、setupChintaiGuide() を1回実行。
 *  ① 賃貸反響タブのヘッダーを色分け（緑=入力 / 赤=GAS自動 / 青=AI）＋A1に使い方メモ
 *  ② スプシ内に「📖使い方」タブを自動生成（色分けガイドを本体に書く）
 * 変数・関数名は GUIDE_ 接頭辞で統合GASと衝突しません。
 */
const GUIDE_SHEET_ID = '1PWI3JyEVaFsTZqOA_DNijBIDtGHMZM2JRhdHORkKQhY';
const GUIDE_IN_TAB   = '賃貸反響';

function setupChintaiGuide() {
  const ss = SpreadsheetApp.openById(GUIDE_SHEET_ID);
  _guideColorizeHeaders(ss);
  _guideBuildSheet(ss);
  try { ss.toast('色分け＋📖使い方タブを作成しました', '賃貸反響ガイド', 5); } catch (e) {}
}

// ① 賃貸反響タブのヘッダー色分け
function _guideColorizeHeaders(ss) {
  const sh = ss.getSheetByName(GUIDE_IN_TAB);
  if (!sh) return;
  const GREEN='#C6EFCE', GREENT='#006100', RED='#FFC7CE', REDT='#9C0006', BLUE='#BDD7EE', BLUET='#1F4E78';
  const paint=(c1,c2,bg,fc)=>{ sh.getRange(1,c1,1,c2-c1+1).setBackground(bg).setFontColor(fc).setFontWeight('bold'); };
  paint(1,14,GREEN,GREENT);   // 入力
  paint(15,20,RED,REDT);      // GAS自動
  paint(21,27,BLUE,BLUET);    // AI（入力/出力）
  sh.setFrozenRows(1);
  sh.getRange(1,1).setNote(
    '【使い方】\n🟢緑=あなたが入力 / 🔴赤=GASが自動 / 🔵青=AI(要APIキー)\n' +
    '手順:緑を埋める→その行を選び メニュー「賃貸反響P1→この行を生成」\n' +
    'メニューが出ない時:Apps Scriptで ct_generateAll を実行');
}

// ② 📖使い方タブを本体に作る
function _guideBuildSheet(ss) {
  const name='📖使い方';
  let g = ss.getSheetByName(name);
  if (g) ss.deleteSheet(g);
  g = ss.insertSheet(name, 0);
  g.setHiddenGridlines(true);
  g.setColumnWidth(1,210); g.setColumnWidth(2,560);
  const RED='#AA2E26', CREAM='#F9F6EF', GREEN='#C6EFCE', REDL='#FFC7CE', BLUE='#BDD7EE', GRAY='#ECECEC';
  let r=1;
  const put=(a,b,bg,bold,sz,fc)=>{ g.getRange(r,1).setValue(a); g.getRange(r,2).setValue(b||'');
    const rng=g.getRange(r,1,1,2); if(bg)rng.setBackground(bg);
    rng.setFontWeight(bold?'bold':'normal').setFontSize(sz||11).setFontColor(fc||'#1A1A1A').setVerticalAlignment('middle').setWrap(true); r++; };
  const sec=(t)=>{ g.getRange(r,1,1,2).merge().setValue(t).setBackground(RED).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12); r++; };

  g.getRange(r,1,1,2).merge().setValue('📖 賃貸反響システム 使い方（全社共通ルール）').setFontWeight('bold').setFontSize(15); r++;
  g.getRange(r,1,1,2).merge().setValue('使えてこそ意味がある。緑=入力／赤=自動／青=AI。').setFontColor('#6E6E6E'); r++; r++;

  sec('① 毎日の流れ（3ステップ）');
  put('STEP1 入力','賃貸反響タブの新しい行に、🟢緑の列を埋める（顧客名・物件・賃料 等）',GREEN,true);
  put('STEP2 生成','その行を選び メニュー「賃貸反響P1 → ★この行を生成」',CREAM,true);
  put('STEP3 確認・送信','🔴赤の列に出た 初期費用・返信文・管理会社確認 を見て送る',REDL,true); r++;

  sec('② 列の色分け');
  put('🟢 緑＝あなたが入力','受信日/顧客名/流入/物件名/賃料/管理費/敷・礼・保証金(ヶ月)/保証料%/火災保険/入居日/契約年数/要望',GREEN);
  put('🔴 赤＝GASが自動(触らない)','初期費用合計/明細/ヒアリング事項/返信文/管理会社確認/生成日',REDL);
  put('🔵 青＝AI(要APIキー)','入力=マイソクPDFリンク/問合せ文面 ／ 出力=AI確度/ニーズ本質/区分/推奨アプローチ/複数返信案',BLUE); r++;

  sec('③ メニュー操作 ＆ 出ない時');
  put('★マイソクから自動入力＋生成','マイソクPDFリンク＋顧客名だけで自動抽出して生成（要APIキー）',CREAM);
  put('★文面からニーズ判定＋複数返信案','問合せ文面から 確度/ニーズ/新規・追客/菊池ボイスの複数案（要APIキー）',CREAM);
  put('★この行を生成','手入力済みの行から生成（APIキー不要）',CREAM);
  put('⚠️メニューが出ない時','Apps Scriptエディタで関数「ct_generateAll」を選んで▶実行＝メニュー無しで全行生成',GRAY,true); r++;

  sec('④ フル稼働の追加設定（使う機能だけ）');
  put('AI機能を使う','スクリプトプロパティに ANTHROPIC_API_KEY を設定（console.anthropic.com/settings/keys）※キーはチャットに貼らない',BLUE);
  put('追客タスク自動化・KPI','サービス＋で Tasks API 追加 → 関数 setupTriggers を1回実行',BLUE); r++;

  sec('⑤ 出力の見方');
  put('初期費用合計','お客様に出す初期費用（例:中目黒502=551,833円）。明細列に内訳',GRAY);
  put('返信文','そのまま送れる①即時接触の文面（菊池ボイス）',GRAY);
  put('管理会社確認','鍵交換/消毒/保証会社/更新料 等 マイソクに無い項目＝元付に要確認',GRAY);
  put('複数返信案(AI)','確度・状況別。新規はエッジで興味づけ、追客は関係前提で簡潔',GRAY);
}
