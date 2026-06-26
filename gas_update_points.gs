/** ポイント/チャージ/マイルの確定値をBSのBI列(6/1)へ反映 2026-06-03 */
function updatePointValues(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('③ 資産負債（BS）'); if(!sh) throw new Error('BS無し');
  var v={  // 行: 6/1値
    82:87379,    // アマギフ1
    83:29327,    // アマギフ2
    84:90670,    // A1(Apple)
    86:1000,     // PayPay1（PayPay合計1,000）
    87:0,        // PayPay2
    89:22566,    // WAON
    94:112500,   // UA
    97:38097,    // えきねっと(JRE通常)
    98:84741,    // MB(Marriott)
    102:63535,   // RP1 楽天市場ポイント（MF確定・キャッシュ混入を是正）
    103:19,      // RP2 楽天市場ポイント2
    104:0        // B1（1/1有効期限切れ）
    // ※r90 RC=楽天キャッシュは楽天ログイン後に別途（今は据置）
  };
  var log=[];
  for(var r in v){ sh.getRange('BI'+r).setValue(v[r]); log.push('BI'+r+'='+v[r]); }
  // r110 モッピー(麻梨奈)=据置・要麻梨奈確認（触らない）
  Logger.log('BS BI列 更新: '+log.join(' / '));
  ss.toast(log.length+'件をBSへ入力（モッピー麻梨奈は据置）','ポイント値反映',6);
}
