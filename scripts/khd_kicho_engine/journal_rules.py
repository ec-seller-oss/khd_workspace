# -*- coding: utf-8 -*-
"""
KHD自動記帳エンジン — 仕訳ルール辞書（編集可・SSoT連動）
=====================================================
土台（既存資産）:
  - MF仕訳帳インポート19列ヘッダ … scripts/amazon_to_mf_journal.py（確定済）
  - 勘定科目マスタ / 受信ログ(取引先→科目の実績) … 書類自動処理パイプライン マスターDB
      (スプシ 12pDC8ffXtNay2TuHCy50T0HHgeHqM0HIRZW4WJdsNPw)
  - 区分(法人/個人事業/家計/投資)＋特殊処理 … Notion「KHD 税務・記帳ルール（全社共通SSoT）」
      (DB 4bfe5ed8-3e8b-4025-ad06-2dc877060e44)

設計原則:
  - 「ルール変更が無い限り全自動」。ここ(RULES)を直せば仕訳が変わる＝唯一の真実源。
  - ルール未HIT＝推測しない。区分=不明 として「要確認」キューへ回す（7割自動＋3割人間）。
  - 家計カードの家計支出は事業会計の記帳対象外（＝MF事業仕訳に出さない。家計集計のみ）。
"""
import re
import unicodedata

# ===== MF仕訳帳インポート 19列ヘッダ（amazon_to_mf_journal.py と完全一致）=====
MF_HEADER = [
    "取引No", "取引日", "借方勘定科目", "借方補助科目", "借方部門", "借方取引先",
    "借方税区分", "借方インボイス", "借方金額(円)", "貸方勘定科目", "貸方補助科目",
    "貸方部門", "貸方取引先", "貸方税区分", "貸方インボイス", "貸方金額(円)",
    "摘要", "タグ", "メモ",
]

# ===== カード／口座プロファイル =====
# default_kubun: そのカードの既定の帰属。家計カードは家計支出が大半。
# kashikata: 貸方科目（カード払い＝未払金 or その口座名）
CARD_PROFILES = {
    "楽天カード(8991)": {
        "default_kubun": "家計",       # 個人の家計カード（食料品・育児が主）
        "kashikata": "未払金",
        "kashikata_hojo": "楽天カード",
        "mf_account": "楽天カード",     # MF一括CSVの「保有金融機関」名（カード抽出キー）
        "note": "楽天3枚合算注意。事業利用が出たらRULESで個別に事業科目へ。",
    },
    # MB＝三井住友ビジネス系カード（事業経費の本丸：SaaS/通信/交通/会費）
    "MB": {
        "default_kubun": "個人事業",   # 事業カード。未HITは家計でなく要確認に寄せる
        "kashikata": "未払金",
        "kashikata_hojo": "MB",
        "mf_account": "MB",
        "note": "ドコモ/ChatGPT/Claude/MS/freee/MFクラウド/Workspace/リベ/note/タイムズ等の定例。",
    },
    # === 家計カード（事業会計の記帳対象外。資金繰りの生活費へ集計）===
    "麻梨奈三井住友カード": {"default_kubun": "家計", "mf_account": "麻梨奈三井住友カード",
                              "kashikata": "未払金", "kashikata_hojo": "麻梨奈三井住友"},
    "d払い": {"default_kubun": "家計", "mf_account": "d払い", "kashikata": "未払金", "kashikata_hojo": "d払い"},
    "VIEW CARD": {"default_kubun": "家計", "mf_account": "VIEW CARD",
                  "kashikata": "未払金", "kashikata_hojo": "VIEW"},

    # === Amazon物販＝仕入。amazon_to_mf_journal.py(注文履歴)が担当→このエンジンは触らない ===
    "Amazon.co.jp": {"default_kubun": "家計", "mf_account": "Amazon.co.jp", "skip": True,
                     "note": "物販仕入は注文履歴pipeline。ギフトチャージ=振替。"},
    "Amazonビジネス": {"default_kubun": "法人", "mf_account": "Amazonビジネス", "skip": True,
                       "note": "クーパン2仕入。注文履歴pipeline担当。"},

    # === 法人口座（経費・借入返済。証書貸付ルールは今後）===
    "【法人】住信SBIネット銀行": {"default_kubun": "法人", "mf_account": "【法人】住信SBIネット銀行",
                                   "kashikata": "普通預金", "kashikata_hojo": "法人SBI"},
    "【法人】城北信用金庫": {"default_kubun": "法人", "mf_account": "【法人】城北信用金庫",
                            "kashikata": "普通預金", "kashikata_hojo": "城北信金"},
}

# ===== 振替（経費でない）パターン … 投資・口座間移動・カード引落し =====
# これらは P/L 経費に入れない。事業なら事業主貸/預け金、家計なら家計外(投資)。
TRANSFER_PATTERNS = [
    (r"投信|投資信託|証券|つみたて|積立.*0．5|楽天証券", "投資積立"),
    (r"カード(ご利用|引落|お支払|自動振替)|口座振替", "カード引落"),
    (r"ＡＴＭ|ATM|入金|振込(?!手数料)", "資金移動"),
]

# ===== 摘要 → 仕訳ルール（上から評価、最初にHITしたものを採用）=====
# fields: pattern(正規表現), 区分, 勘定科目, 税区分, メモ
# 区分: 法人 / 個人事業 / 家計 / 投資
# 税区分: 課仕10% / 課仕10%（経過措置80%）/ 対象外 / 非課税  (amazonと同体系)
# ※ pattern は NFKC正規化後のテキストに対して評価される（半角カナ→全角、ＶＩＳＡ→VISA）。
#   評価順は重要：specificな店（セリア＝100均）を、含意の広いスーパー語(セイユウ)より先に置く。
RULES = [
    # --- 家計（事業会計の記帳対象外。家計簿/資金繰り側で集計）---
    # 育児・保育（最優先で確定させる）
    (r"ポピンズ|POPPINS|保育|シツター|ベビーシッター",
     "家計", "教育・保育", "対象外", "育児・保育(ポピンズ)"),
    # 日用品・育児・家電（セリアは『セイユウ』を含むので食費より先に判定）
    (r"セリア|ダイソー|キャンドゥ|ノジマ|ミルクコウボウ|ベビー館|赤ちゃん本舗|アカチヤン",
     "家計", "日用品", "対象外", "日用品・育児"),
    # スーパー（食料品）
    (r"マイバスケツト|マイバスケット|西友|セイユウ|イト-ヨ-カド|イトーヨーカド|ヨ-カド|ライフ|まいばす",
     "家計", "食費", "対象外", "食料品(スーパー)"),
    (r"ピザ|ﾋﾟｻﾞ|マクドナルド|スターバックス|スタバ|ガスト|サイゼ|外食|レストラン|ハツト",
     "家計", "外食費", "対象外", "外食"),
    (r"GU MODE|GU |ジーユー|UNIQLO|ユニクロ|衣料|アパレル",
     "家計", "被服費", "対象外", "衣料品"),

    # --- 事業（個人事業/法人）: MB(ビジネスカード)実データで育成 ---
    # 通信費（携帯・回線・サーバー・Workspace）
    (r"ドコモ|docomo|au |ソフトバンク|楽天モバイル|AWS|AmazonWeb|さくらインターネット|エックスサーバー|ドメイン|ムームー|GOOGLE.?WORKSPACE|Google.*Work|GWS",
     "個人事業", "通信費", "課仕10%", "通信・サーバー・Workspace"),
    # SaaS・ツール（生成AI/会計/MS等のサブスク）
    (r"Adobe|ChatGPT|CHATGPT|OpenAI|OPENAI|Anthropic|CLAUDE|Claude|Notion|MICROSOFT|Microsoft|FREEE|freee|マネ.?フォワ.?ド|note 東京都|ノート 東京都",
     "個人事業", "支払手数料", "課仕10%", "SaaS・サブスク"),
    # 旅費交通費（公共交通・カーシェア・航空）
    (r"Suica|スイカ|PASMO|JR東日本|新幹線|ETC|高速道路|高速料金|タクシー|タイムズカー|タイムズ24|カーシェア|駐車場|NEXCO|旅費|全日空|ANA国内|スマートEチケット",
     "個人事業", "旅費交通費", "課仕10%", "交通費"),
    # 消耗品・備品
    (r"アスクル|ASKUL|モノタロウ|ヨドバシ|ビックカメラ|文具|事務用品|スキャナー|ブックスキャナ",
     "個人事業", "消耗品費", "課仕10%", "事務・消耗品・備品"),
    # 諸会費・租税
    (r"年会費|カード年会費", "個人事業", "諸会費", "対象外", "カード年会費"),
    (r"^消費税$|租税|収入印紙|印紙税", "個人事業", "租税公課", "対象外", "租税公課"),

    # --- 特殊処理（税務SSoT準拠）---
    (r"税理士|橋本", "個人事業", "支払報酬", "課仕10%", "税理士報酬(橋本)"),
    (r"リベ(大|シティ)|オンラインサロン", "個人事業", "研修費", "課仕10%", "サロン会費"),
]


def find_transfer(text):
    """振替（経費でない）判定。HITすれば種別を返す。"""
    t = text or ""
    for pat, kind in TRANSFER_PATTERNS:
        if re.search(pat, t):
            return kind
    return None


def classify(store, card_profile):
    """摘要(店名) → 仕訳分類 dict。未HITは区分=不明（要確認へ）。"""
    # 半角カナ/全角英数のゆれを吸収（正本CSV=半角カナ／MF=全角 の名寄せ）
    text = unicodedata.normalize("NFKC", store or "")
    # 1) 振替（経費でない）を最優先で除外
    tk = find_transfer(text)
    if tk:
        return {"区分": "投資" if tk == "投資積立" else "振替",
                "勘定科目": tk, "税区分": "対象外", "メモ": f"振替:{tk}", "is_keihi": False}
    # 2) ルール辞書
    for pat, kubun, kamoku, zei, memo in RULES:
        if re.search(pat, text):
            return {"区分": kubun, "勘定科目": kamoku, "税区分": zei,
                    "メモ": memo, "is_keihi": kubun in ("法人", "個人事業")}
    # 3) 未HIT → 要確認
    return {"区分": "不明", "勘定科目": "", "税区分": "",
            "メモ": "ルール未HIT→要確認", "is_keihi": False}
