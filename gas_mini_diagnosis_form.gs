/**
 * KHD 立地・商圏 無料ミニ診断フォーム 自動生成スクリプト
 * 顔は医療で尖らせ、Q1で3本柱(医療開業/訪問マッサージ等開業/不動産売買)を拾う拡張版。
 *
 * 使い方：
 *  1. https://script.google.com → 新しいプロジェクト
 *  2. このコードを全部貼り付け
 *  3. 関数 createMiniDiagnosisForm を実行（初回は承認ダイアログでGoogleアカウント許可）
 *  4. 実行ログ(表示→ログ)に「公開URL」「編集URL」が出る → 公開URLをHP/X/noteのCTAに設定
 */
function createMiniDiagnosisForm() {
  var form = FormApp.create('立地・商圏 無料ミニ診断｜KHD（医療×不動産×AI）');

  form.setDescription(
    '候補地と相談内容を教えていただければ、概算の「立地・商圏」を無料でお返しします。\n' +
    '宅建士が診療圏調査・商圏分析から物件・契約までワンストップで対応します。\n' +
    '売り込みはしません。まず判断材料をお渡しします（48時間以内に返信）。'
  );

  // 設定：メール収集なし・回答1回制限なし・進捗バーなし
  form.setCollectEmail(false);
  form.setLimitOneResponsePerUser(false);
  form.setProgressBar(false);
  form.setAllowResponseEdits(false);

  // Q1【必須】ご相談内容（複数選択可）＝3本柱を拾う
  form.addCheckboxItem()
    .setTitle('ご相談内容（複数選択可）')
    .setChoiceValues([
      'クリニック開業の立地・診療圏を知りたい',
      '訪問マッサージ／治療院などの開業・商圏を知りたい',
      '不動産の売買（買いたい／売りたい）を相談したい',
      'AI活用・その他'
    ])
    .setRequired(true);

  // Q2【必須】候補地・対象エリア
  form.addTextItem()
    .setTitle('候補地・対象エリア（住所 または 最寄駅。なるべく具体的に）')
    .setHelpText('例：千葉県船橋市本町／JR船橋駅 徒歩5分 など')
    .setRequired(true);

  // Q3【必須】ご連絡先
  form.addTextItem()
    .setTitle('ご連絡先（メール または X のDM可・IDを記入）')
    .setRequired(true);

  // Q4 業種・診療科目（任意）
  form.addTextItem()
    .setTitle('業種・診療科目（任意）')
    .setHelpText('例：産婦人科／内科／訪問マッサージ／戸建売却 など')
    .setRequired(false);

  // Q5 時期の目安（任意）
  form.addMultipleChoiceItem()
    .setTitle('時期の目安')
    .setChoiceValues(['未定', '半年以内', '1年以内', '1〜3年'])
    .setRequired(false);

  // Q6【魔法の質問】いま一番不安なこと
  form.addParagraphTextItem()
    .setTitle('いま一番"不安なこと"を一言で（自由記述）')
    .setHelpText('ここが一番知りたいところです。率直にどうぞ。')
    .setRequired(false);

  // 送信後メッセージ
  form.setConfirmationMessage(
    '送信ありがとうございます。48時間以内に概算の立地・商圏をお返しします。\n' +
    '売り込みはしません。まず材料をお渡しします。'
  );

  var pubUrl = form.getPublishedUrl();
  var editUrl = form.getEditUrl();
  Logger.log('=== 作成完了 ===');
  Logger.log('公開URL（CTAに貼る）: ' + pubUrl);
  Logger.log('編集URL（中身を直す）: ' + editUrl);
  Logger.log('回答はこのフォームの「回答」タブ→スプレッドシート連携で蓄積できます。');
  return { published: pubUrl, edit: editUrl };
}
