#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
仕入スクリーニング 統合ランナー（融資資料自動化TTP・自社版 スコープA 通し）
================================================================================
物件PDF 1本 → 抽出 → 路線価等のエンリッチ → 玉川式KPI判定 → 判定シート、を1コマンドで。

  property_intake.py (PDF抽出) → エンリッチ → property_screener.py (判定)

【路線価の扱い（3段階・上から優先）】
  1) --rosenka <円/㎡>          相続税路線価を直接指定（最も正確）
  2) --kosho <円/㎡>            公示地価を指定 → 相続税路線価 ≈ 公示×0.8 で推定
  3) （将来）reinfolib公式API   REINFOLIB_API_KEY 登録後に自動取得へ差し替え予定
  いずれも無ければ None のまま（screener が❌で正直表示）。

使い方:
  python3 screen_property.py 物件.pdf
  python3 screen_property.py 物件.pdf --rosenka 95000 --uridashi 52000000 --kenchikuhi 16000000
  python3 screen_property.py 物件.pdf --kosho 120000 --uridashi 52000000 --kenchikuhi 16000000
"""
import sys
import json
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import property_intake as intake
import property_screener as screener
try:
    import reinfolib_client as reinfo
except Exception:
    reinfo = None

# 相続税路線価 ≈ 公示地価 × 0.8（相続税評価額は公示の概ね8割が目安）
KOSHO_TO_ROSENKA = 0.80


def yen_arg(v):
    """「3000万」「52000000」等を円intに。"""
    if v is None:
        return None
    s = str(v).replace(",", "")
    if "万" in s:
        return int(float(s.replace("万", "")) * 10000)
    return int(float(s))


def main():
    ap = argparse.ArgumentParser(description="仕入スクリーニング 統合ランナー（スコープA）")
    ap.add_argument("pdf", nargs="?", help="物件PDFパス（省略時は--address番地入力モード）")
    ap.add_argument("--address", help="番地入力モード：住所を1つ（PDF無しで査定）例: 岩手県盛岡市高松2丁目34-5")
    ap.add_argument("--kenbiya", help="健美家listing URL（--urlでも可）")
    ap.add_argument("--url", help="ポータル物件URL（健美家/楽待を自動判別）→詳細自動取得して本査定")
    ap.add_argument("--name", help="物件名（番地入力モードの表示名）")
    ap.add_argument("--rosenka", help="相続税路線価 円/㎡（直接指定・最優先）")
    ap.add_argument("--kosho", help="公示地価 円/㎡（×0.8で路線価推定）")
    # reinfolib自動取得（実勢取引㎡単価×0.72→路線価相当）
    ap.add_argument("--pref", help="都道府県（例: 千葉県）reinfolib自動取得用")
    ap.add_argument("--city", help="市区町村（例: 船橋市／政令市は 横浜市鶴見区）")
    ap.add_argument("--district", help="町名（例: 飯山満町）任意・精度向上用")
    ap.add_argument("--uridashi", help="想定売却額（円・「万」可）キャピタル用")
    ap.add_argument("--kenchikuhi", help="建築費（円・「万」可）キャピタル用")
    # インカム（収益物件）入力
    ap.add_argument("--nenchin", help="年間賃料（円・「万」可）→ 収益物件モード起動")
    ap.add_argument("--loan", help="融資額（円・「万」可）CF試算用")
    ap.add_argument("--rate", help="金利%（例 2.5）")
    ap.add_argument("--years", help="融資期間（年）")
    ap.add_argument("--jiko", help="自己資金（円・「万」可）CCR用")
    # 融資資料デッキ生成
    ap.add_argument("--deck", action="store_true", help="判定後に融資資料デッキ(pptx)を生成")
    ap.add_argument("--bank", help="提出先銀行名（デッキ表紙用 例: 岩手銀行）")
    ap.add_argument("--db", action="store_true", help="査定結果を物DB(SSoT)へ自動upsert")
    # 画像PDF等で抽出が❌のときの上書き（Claude推論フォールバックの受け皿）
    ap.add_argument("--price", help="価格 上書き（円・「万」可）")
    ap.add_argument("--land", help="土地面積 上書き（㎡）")
    ap.add_argument("--yoseki", help="容積率 上書き（%）")
    args = ap.parse_args()

    portal_url = args.url or args.kenbiya
    portal_ref = None
    if not args.pdf and not args.address and not portal_url:
        print("❌ 物件PDF / --address / --url のいずれかを指定してください")
        sys.exit(1)

    # ── Step0: ポータルURL → 自動取得（資料取得→本査定）──────
    if portal_url:
        import auto_pipeline
        kf = auto_pipeline.fetch_fields(portal_url)
        ref = kf["_参考情報"]
        portal = "楽待" if "rakumachi" in portal_url else ("健美家" if "kenbiya" in portal_url else "ポータル")
        print("=" * 56)
        print(f"  [0] {portal} 自動取得 — {ref.get('所在地')}")
        print(f"      土地{kf['土地面積_m2']}㎡ / 建物{ref.get('建物面積_m2')}㎡ / {ref.get('構造')} / {ref.get('築年月')}")
        # CLI未指定の項目を健美家データで補完（CLI優先）
        args.price = args.price or (str(kf["価格_円"]) if kf.get("価格_円") else None)
        args.land = args.land or (str(kf["土地面積_m2"]) if kf.get("土地面積_m2") else None)
        args.yoseki = args.yoseki or (str(kf["容積率_pct"]) if kf.get("容積率_pct") else None)
        args.nenchin = args.nenchin or (str(kf["年賃料_円"]) if kf.get("年賃料_円") else None)
        args.address = args.address or ref.get("所在地")
        args.name = args.name or kf["物件名"]
        portal_ref = ref   # 構造/築年/建物面積を積算評価へ引き継ぐ

    # ── Step1: 入力（PDF抽出 or 番地入力モード）────────────
    if not args.pdf:
        # 番地入力モード：住所→pref/city/districtを自動分解、fieldsはCLIから組む
        pref, city, district = reinfo.parse_address(args.address) if reinfo else (None, None, None)
        args.pref = args.pref or pref
        args.city = args.city or city
        args.district = args.district or district
        name = args.name or args.address
        fields = {"物件名": name, "価格_円": None, "土地面積_m2": None, "容積率_pct": None,
                  "路線価_円per_m2": None, "想定売却額_円": None, "建築費_円": None}
        meta = {"_参考情報": {"所在地": args.address, "用途地域": None, "建蔽率_pct": None, "接道メモ": None},
                "_画像PDF疑い": False}
        if portal_ref:   # ポータル取得の構造/築年/建物面積を引き継ぐ（積算評価用）
            meta["_参考情報"].update({k: v for k, v in portal_ref.items() if v is not None})
        text_len, missing = 0, []
        print("=" * 56)
        print(f"  [1/3] 番地入力モード — {name}")
        print("=" * 56)
        print(f"  住所分解: {pref} / {city} / {district}")
    else:
        if not Path(args.pdf).exists():
            print(f"❌ ファイルが見つかりません: {args.pdf}")
            sys.exit(1)
        fields, meta, text_len, missing = intake.parse(args.pdf)
    # 上書き（画像PDF/抽出ミス/番地入力モードの手入力）。指定があれば優先。
    ovr = []
    if args.price:
        fields["価格_円"] = yen_arg(args.price); ovr.append("価格")
    if args.land:
        fields["土地面積_m2"] = float(args.land); ovr.append("土地面積")
    if args.yoseki:
        fields["容積率_pct"] = float(args.yoseki); ovr.append("容積率")
    if args.pdf:
        print("=" * 56)
        print(f"  [1/3] PDF抽出 — {fields['物件名']}")
        print("=" * 56)
    print(f"  価格: {fields['価格_円'] or '❌'} / 土地: {fields['土地面積_m2'] or '❌'}㎡ / "
          f"容積率: {fields['容積率_pct'] or '❌'}%")
    if ovr:
        print(f"  ✍️ 入力: {', '.join(ovr)}")
    if meta["_画像PDF疑い"]:
        print("  ⚠️ 画像PDF疑い → ClaudeにPDFをReadさせて各項目を手入力で補完してください")
    r = meta["_参考情報"]
    print(f"  所在地: {r['所在地'] or '—'} / 用途: {r['用途地域'] or '—'}")

    # ── Step2: エンリッチ（路線価・売却額・建築費）──────────
    rosenka = None
    src = None
    if args.rosenka:
        rosenka = yen_arg(args.rosenka)
        src = "直接指定"
    elif args.kosho:
        rosenka = int(yen_arg(args.kosho) * KOSHO_TO_ROSENKA)
        src = f"公示地価{yen_arg(args.kosho):,}×0.8 推定"
    elif args.pref and args.city and reinfo:
        rl = reinfo.lookup(args.pref, args.city, args.district)
        if rl.get("error"):
            print(f"  ⚠️ reinfolib取得失敗: {rl['error']}")
        else:
            rosenka = rl["相続税路線価相当"]
            src = (f"reinfolib実勢{rl['実勢㎡単価_中央値']:,}×{rl['換算係数']} "
                   f"［{rl['市区町村']}{rl.get('町名') or ''}・{rl['件数']}件・{rl['集計範囲']}］")
    fields["路線価_円per_m2"] = rosenka
    fields["想定売却額_円"] = yen_arg(args.uridashi)
    fields["建築費_円"] = yen_arg(args.kenchikuhi)
    # インカム入力（CLI未指定なら PDF/ポータル抽出の年賃料を維持）
    fields["年賃料_円"] = yen_arg(args.nenchin) or fields.get("年賃料_円")
    fields["融資額_円"] = yen_arg(args.loan)
    fields["金利_pct"] = float(args.rate) if args.rate else None
    fields["期間_年"] = float(args.years) if args.years else None
    fields["自己資金_円"] = yen_arg(args.jiko)

    def yen_or_x(v):
        return f"{v:,}円" if v else "❌"

    print("-" * 56)
    print(f"  [2/3] エンリッチ")
    print(f"  路線価: {f'{rosenka:,} 円/㎡（{src}）' if rosenka else '❌ 未入力'}")
    print(f"  想定売却額: {yen_or_x(fields['想定売却額_円'])} / 建築費: {yen_or_x(fields['建築費_円'])}")

    # 抽出+エンリッチ済みJSONを保存（再利用可）
    out_json = intake.OUT_DIR / f"intake_{fields['物件名']}.json"
    intake.OUT_DIR.mkdir(parents=True, exist_ok=True)
    out_json.write_text(json.dumps({**fields, **meta}, ensure_ascii=False, indent=2),
                        encoding="utf-8")

    # ── Step3: 判定 ───────────────────────────────
    print("-" * 56)
    print(f"  [3/3] 玉川式KPI判定")
    print()
    fields["_参考情報"] = meta.get("_参考情報", {})   # 積算評価(構造/築年/建物)を渡す
    res = screener.calc(fields)
    console, verdict, mark, signals = screener.render_console(fields, res)
    print(console)
    xlsx = screener.write_excel(fields, res, verdict, mark, signals)
    print(f"\n📄 判定シート: {xlsx}")
    print(f"📄 抽出JSON  : {out_json}")

    # ── Step4: 融資資料デッキ（--deck）────────────────
    if args.deck:
        try:
            import loan_deck
        except Exception as e:
            print(f"  ⚠️ loan_deck読込失敗: {e}")
            return
        deck = assemble_deck(fields, res, args.bank)
        out_pptx = intake.OUT_DIR / f"融資資料_{fields['物件名']}.pptx"
        loan_deck.build(deck, str(out_pptx))
        print(f"📊 融資資料デッキ: {out_pptx}")

    # ── Step5: 物DB(SSoT)へ自動蓄積（--db）───────────────
    if args.db:
        try:
            import sheets_db
            row, mode, url = sheets_db.upsert_property(fields, res, verdict)
            print(f"🗄  物DB(SSoT) {mode}: 行{row} → {url}")
        except Exception as e:
            print(f"  ⚠️ 物DB蓄積失敗: {e}")


def _man(v):
    """円→千円表記。"""
    v = yen_arg(v) if isinstance(v, str) else v
    return f"{int(v/1000):,}千円" if v else "—"


def assemble_deck(fields, res, bank):
    """判定データ → loan_deck用dict。単一物件。"""
    name = fields.get("物件名", "物件")
    price = fields.get("価格_円")
    is_income = res.get("is_income", False)
    pc = lambda x: f"{x*100:.1f}%" if x is not None else "—"
    # KPIカードは (値, 色, sub) 形式
    from loan_deck import BRICK, GREEN, INK  # noqa
    kpi = {"取得価格": (_man(price), BRICK, "")}
    if res.get("土地値割合") is not None:
        kpi["土地値割合"] = (pc(res["土地値割合"]), GREEN, "担保")
    if is_income:
        kpi["実質利回り"] = (pc(res.get("実質利回り")), GREEN, "経費15%控除")
        if res.get("CF率") is not None:
            kpi["CF率"] = (pc(res.get("CF率")), GREEN, "税引前CF/価格")
        income = {"rent_y": _man(fields.get("年賃料_円")), "omote": pc(res.get("表面利回り")),
                  "real": pc(res.get("実質利回り")),
                  "cf": (_man(res.get("税引前CF")) if res.get("税引前CF") else "—"),
                  "rows": [("取得価格", _man(price)), ("年間賃料", _man(fields.get("年賃料_円"))),
                           ("実質利回り", pc(res.get("実質利回り"))),
                           ("税引前CF(試算)", _man(res.get("税引前CF")) if res.get("税引前CF") else "—")]}
        capital = {}
        gensen, scheme = "賃料収入", "保有（インカム）"
    else:
        if res.get("粗利率") is not None:
            kpi["粗利率"] = (pc(res.get("粗利率")), GREEN, "キャピタル")
        income = {}
        capital = {"price": _man(price), "uridashi": _man(fields.get("想定売却額_円")),
                   "arari": pc(res.get("粗利率")),
                   "rows": [("取得価格", _man(price)), ("想定売却額", _man(fields.get("想定売却額_円"))),
                            ("建築費", _man(fields.get("建築費_円")))]}
        gensen, scheme = "売却代金", "再販（キャピタル）"

    ref = (fields.get("_参考情報") or {})
    rosenka = {}
    if fields.get("路線価_円per_m2"):
        b = res.get("融資カバー率B")
        rosenka = {"souzoku": f"約{int(fields['路線価_円per_m2']):,}円/㎡",
                   "jissei": "—",
                   "tochine": pc(res.get("土地値割合")),
                   "sekisan": (f"{res['積算割合']*100:.0f}%（積算{int(res['積算評価_円']/10000):,}万円）"
                               if res.get("積算割合") is not None else None),
                   "coverB": (f"{b:.2f}（融資カバー率B・1以下で担保堅い）" if b is not None else None)}
    return {
        "title": "融資ご相談資料", "sub": name, "bank": (f"{bank} 御中" if bank else None),
        "company": "菊池ホールディングス（KHD）", "rep": "代表　菊池 研太",
        "kpi": kpi,
        "overview": [("所在地", ref.get("所在地") or "—"),
                     ("用途地域", ref.get("用途地域") or "—"),
                     ("土地面積", f"{fields.get('土地面積_m2','—')} ㎡"),
                     ("スキーム", scheme)],
        "props": [{"name": name, "price": _man(price), "addr": ref.get("所在地") or "—",
                   "land": f"{fields.get('土地面積_m2','—')} ㎡",
                   "rent_y": _man(fields.get("年賃料_円")) if is_income else "—",
                   "occ": "—"}],
        "is_income": is_income, "income": income, "capital": capital, "rosenka": rosenka,
        "loan_req": _man(fields.get("融資額_円") or price), "gensen": gensen, "scheme": scheme,
        "closing": "本案件へのご支援を賜りますようお願い申し上げます。",
    }


if __name__ == "__main__":
    main()
