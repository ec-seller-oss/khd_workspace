#!/usr/bin/env bash
# 01本部: 04_案件概要書 / 04_診療権調査(診療圏調査ベース) / 04_セミナー資料 を一括生成
set -euo pipefail

DIR="$(cd "$(dirname "$0")" && pwd)"
CONFIG="${DIR}/config.json"
PYTHON="${PYTHON:-python3}"

usage() {
  echo "Usage: $0 --name <案件名> [--date YYMMDD] [--out-dir <出力ルート>] [--no-xlsx]" >&2
  echo "  例: $0 --name \"つちはし整形外科\"" >&2
  exit 1
}

NAME=""
DATE_PREFIX="$(date +%y%m%d)"
OUT_ROOT=""
NO_XLSX=0

while [[ $# -gt 0 ]]; do
  case "$1" in
    --name) NAME="${2:-}"; shift 2 ;;
    --date) DATE_PREFIX="${2:-}"; shift 2 ;;
    --out-dir) OUT_ROOT="${2:-}"; shift 2 ;;
    --no-xlsx) NO_XLSX=1; shift ;;
    -h|--help) usage ;;
    *) echo "Unknown arg: $1" >&2; usage ;;
  esac
done

if [[ -z "$NAME" ]]; then
  usage
fi

eval "$("$PYTHON" - "$CONFIG" <<'PY'
import json, sys, shlex
c = json.load(open(sys.argv[1], encoding="utf-8"))
t = c["templates"]
out = c.get("output_root") or ""
for k, v in [
    ("T_PPTX", t["case_summary_pptx"]),
    ("T_PDF", t["case_summary_pdf"]),
    ("T_XLSX", t["medical_area_survey_xlsx"]),
    ("DEF_OUT", out),
]:
    print(f"{k}={shlex.quote(v)}")
PY
)"

if [[ -n "$OUT_ROOT" ]]; then
  DEF_OUT="$OUT_ROOT"
fi
mkdir -p "$DEF_OUT"

sanitize() {
  echo "$1" | sed 's/[\/\\:*?"<>|]/_/g' | sed 's/^[[:space:]]*//;s/[[:space:]]*$//'
}

SAFE_NAME="$(sanitize "$NAME")"
FOLDER="${DEF_OUT}/${DATE_PREFIX}_${SAFE_NAME}"

D1="${FOLDER}/04_案件概要書"
D2="${FOLDER}/04_診療権調査"
D3="${FOLDER}/04_セミナー資料"

mkdir -p "$D1" "$D2" "$D3"

for f in "$T_PPTX" "$T_PDF"; do
  if [[ ! -f "$f" ]]; then
    echo "ERROR: テンプレが見つかりません: $f" >&2
    echo "config.json の templates を環境に合わせて修正してください。" >&2
    exit 1
  fi
done

cp -p "$T_PPTX" "${D1}/04_案件概要書(${NAME}).pptx"
cp -p "$T_PDF" "${D1}/04_案件概要書(${NAME}).pdf"

if [[ "$NO_XLSX" -eq 0 ]]; then
  if [[ -f "$T_XLSX" ]]; then
    cp -p "$T_XLSX" "${D2}/04_診療圏調査_${NAME}.xlsx"
  else
    echo "WARN: 診療圏調査xlsxが見つかりません（スキップ）: $T_XLSX" >&2
  fi
fi

NOTE="${D2}/_README_診療権と診療圏.txt"
cat > "$NOTE" <<NOTE
【01本部・資料自動化】
フォルダ名は「04_診療権調査」ですが、現行テンプレは診療圏（エリア・競合）調査の Excel を流用しています。
診療権（開設許可・区域など）の法務確認は別紙・別チェックリストで管理してください。
NOTE

SEMINAR="${D3}/セミナー下書き_会社概要・ビジネスモデル・仲間と顧客募集.md"
cat > "$SEMINAR" <<MD
# セミナー資料（下書き）

案件: ${NAME}
生成日: $(date +%Y-%m-%d)

---

## 1. 会社概要

- 会社名・設立・所在地
- 代表・体制（本部／現場）
- 実績・強み（数字1〜3点）
- コンプライアンス・許認可（不動産・関連）

## 2. ビジネスモデル

- 収益の柱（例: 買取再販、仲介、管理、開発 等）
- 案件の流れ（リード →  DD →  契約 →  Exit）
- リスクと手当（法務・税務・工事・販売）
- 本案件との接点（なぜこの施設・エリアか）

## 3. 仲間募集（パートナー）

- 求める役割（資本・紹介・工事・運用 等）
- 条件の考え方（スキーム概要のみ。個別条件は別途）
- 連絡・次のアクション

## 4. 顧客・利用者募集

- ターゲット（例: 医療法人、開業医、テナント 等）
- 提供価値（立地・スキーム・スケジュール）
- 問い合わせ先

---

※ Google スライドに貼る場合は、見出しごとに1チャプターずつ展開してください。
MD

META="${FOLDER}/case_kit.meta.json"
"$PYTHON" - "$NAME" "$DATE_PREFIX" "$FOLDER" "$T_PPTX" "$T_PDF" "$T_XLSX" <<'PY' >"$META"
import json, sys, datetime
name, date_prefix, folder, pptx, pdf, xlsx = sys.argv[1:]
out = {
    "division": "01本部",
    "case_name": name,
    "date_prefix": date_prefix,
    "generated_at": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
    "output_folder": folder,
    "templates_used": {"case_summary_pptx": pptx, "case_summary_pdf": pdf, "medical_area_survey_xlsx": xlsx},
}
print(json.dumps(out, ensure_ascii=False, indent=2))
PY

echo ""
echo "作成しました: $FOLDER"
echo "  - ${D1}"
echo "  - ${D2}"
echo "  - ${D3}"
