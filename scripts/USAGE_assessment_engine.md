# KHD AI査定エンジン 使い方（2026-06-02 完成版）

物件情報を入れる → 玉川式で判定 → 物DB蓄積 → 銀行融資資料デッキ。
入口は4つ。すべて `scripts/` で実行。出力は `out_screener/`。

---

## 🟦 入口① マイソクPDF
```
python3 scripts/screen_property.py 物件.pdf
```
- PDFから 価格/土地面積/容積率/所在地/年賃料 を自動抽出（athome対応済）
- 画像PDFで抽出できない時は手入力で補完: `--price 3000万 --land 200 --yoseki 200`

## 🟩 入口② 住所だけ（番地入力モード）
```
python3 scripts/screen_property.py --address "岩手県盛岡市高松2丁目34-5" --price 2500万 --land 335 --nenchin 352万
```
- 住所から都道府県/市区町村を自動分解 → 路線価を自動取得
- マイソク不要。価格・面積・年賃料は手入力

## 🟧 入口③ ポータルURL（健美家・楽待）
```
python3 scripts/screen_property.py --url "https://www.kenbiya.com/pp2/s/.../re_xxxx/"
python3 scripts/screen_property.py --url "https://www.rakumachi.jp/.../show.html"
```
- URL 1本で 土地面積/価格/利回り/用途等を自動取得 → 路線価自動 → 本査定（担保まで）
- 健美家=フル取得／楽待=価格・土地・容積のみ（利回り等はログイン制限）

## 🟪 入口④ 業者メール（自動収集）
```
python3 scripts/queue_assess.py        # スプシの業者メールを件名で一次査定→書戻し
python3 scripts/auto_pipeline.py --sheet   # スプシO列URLを本査定→🟢通知（無人用）
```

---

## ⚙️ よく足すオプション（①〜③共通）
| オプション | 意味 |
|---|---|
| `--bank 岩手銀行` | デッキ表紙の提出先銀行 |
| `--deck` | 融資資料デッキ(pptx)を生成 |
| `--db` | 物DB(SSoT)へ査定結果を自動蓄積 |
| `--nenchin 473万` | 年間賃料 → 収益物件モード（利回り/CF/CCR） |
| `--loan 7480万 --rate 2.5 --years 25` | CF試算の融資条件 |
| `--jiko 1000万` | 自己資金 → CCR算出 |
| `--uridashi 6500万 --kenchikuhi 2000万` | 再販(キャピタル)の粗利率算出 |
| `--rosenka 95000` / `--kosho 120000` | 路線価を手で固定（自動より優先） |

### フル例（実弾：盛岡の収益物件を判定＋DB＋銀行デッキ）
```
python3 scripts/screen_property.py --address "岩手県盛岡市高松2丁目34-5" \
   --name "アーバンキューブ" --price 2500万 --land 335.66 --nenchin 352.8万 \
   --loan 2500万 --rate 2.5 --years 20 --bank 岩手銀行 --db --deck
```

---

## 🔎 物件を選ぶ・貯める（補助ツール）
| コマンド | 役割 |
|---|---|
| `python3 scripts/pick_good.py` | 物DBから🟢買い候補だけ抽出（オートピック） |
| `python3 scripts/queue_assess.py` | 業者メールを一次査定（件名ベース） |
| `python3 scripts/auto_pipeline.py <URL…>` | URL群をまとめて本査定→🟢通知 |

## 📂 出力先
- 判定シート / 抽出JSON / 融資資料デッキ → `out_screener/`
- 物DB（SSoT）→ Google Sheets「物」タブ（`--db`時）
- 査定キュー/オートピック結果 → `.company/secretary/notes/`

---

## 🤖 完全無人（設定すれば毎朝自動）
1. GAS `gas_property_url_extract.gs` を貼付 → listing URLをスプシO列へ
2. `com.khd.autopipeline.plist` を `~/Library/LaunchAgents/` に置いて `launchctl load`
→ 毎朝7時に収集物件を本査定し、🟢だけ通知

---

## 「最初」と何が変わったか
| | 当初(6/2朝) | 今(完成) |
|---|---|---|
| 入口 | PDFのみ | PDF / 住所 / URL / 業者メール |
| 路線価 | 手入力必須 | reinfolib自動 |
| 評価軸 | キャピタル(粗利率)のみ | ＋インカム(利回り/CF/CCR) |
| 出力 | 判定シートのみ | ＋物DB蓄積＋銀行デッキ |
| 収集 | なし | 業者メール自動収集→査定キュー |
| 運用 | 都度手動 | オートピック＋無人化 |
