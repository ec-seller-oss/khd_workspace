# -*- coding: utf-8 -*-
"""
開業用スプシ「04_事業計画書（オーロラ船橋店）」に
「24_公庫申込パック(依頼文・段取り)」タブを追加。
依頼文全文／申込戦略(借りれるだけ・長く)／創業計画書の空欄記入値／
法人初回の窓口・タイミング／必要書類／想定問答 を1タブに集約。
auth: scripts/sheets_token.pickle
"""
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

TOKEN="/Users/kikuchikenta/01_honbu_docs_automation/scripts/sheets_token.pickle"
SID="1ruE_rE8i_OKpO0JMV-82AJDMqLPf8S-r4HxJIEgyL8U"
TAB="24_公庫申込パック(依頼文・段取り)"

RED={"red":0.667,"green":0.180,"blue":0.149}; REDD={"red":0.549,"green":0.141,"blue":0.114}
CARD={"red":0.945,"green":0.925,"blue":0.882}; REDBG={"red":0.957,"green":0.894,"blue":0.886}
WHT={"red":1,"green":1,"blue":1}; INK={"red":0.1,"green":0.1,"blue":0.1}
INBG={"red":0.999,"green":0.984,"blue":0.882}; TIERA={"red":0.886,"green":0.949,"blue":0.905}
SEC={"red":0.357,"green":0.357,"blue":0.357}; HI={"red":1.0,"green":0.898,"blue":0.6}
BLUEBG={"red":0.812,"green":0.886,"blue":0.953}

COVER = """日本政策金融公庫 国民生活事業 ご担当者様

「新規開業資金」お借入のご相談（訪問医療マッサージ事業／株式会社KHD）

平素よりお世話になっております。不動産・医療クリニック承継コンサルティングを営む株式会社KHD 代表の菊池研太と申します。個人事業では貴庫よりお借入れをいただき、約定どおり返済を継続しております（法人としては初めてのご相談となります）。

このたび、医療保険適用の訪問マッサージ・はり灸FC「オーロラ治療院」（運営：株式会社フライハイト・全国約65店、リピート率98%）に加盟し、千葉県船橋市（本部テリトリー承認済みの独占エリア）にて開業いたします。つきましては、設備資金・運転資金として総額1,100万円のご融資をご相談申し上げます。

【事業の確度＝3つの裏付け】
1. 需要：船橋市の後期高齢者は8万人超。在宅医療マッサージの需要が構造的に拡大。
2. 実証：現役オーナー（板橋店・開業3ヶ月で黒字化）への実地ヒアリングで集客率・採用・解約率を実数検証し、保守的な前提（解約率5%等）で5ヵ年計画を策定済み。加盟前にIndeedで採用テストを行い、有資格者の応募を確認済み。
3. 体制：既存事業の医療・介護ネットワーク（医師約140件・ケアマネ網）を集客に直結。本部のレセプト代行で未収リスクを抑制。

【ご相談条件】
・設備資金 約600万円（加盟金・研修費・保証金・備品・中古車両ほか）…ご返済15年（制度上限内で長期）・据置2年を希望
・運転資金 約500万円（立上げ人件費・家賃・療養費入金サイト2〜4ヶ月の立替）…ご返済10年・据置2年を希望
※療養費（レセプト）は請求から入金まで2〜4ヶ月を要するため、運転資金を厚めに計画しております（資金繰り表添付）。
※加盟金等のお支払いは融資ご実行後を予定し、本部と支払期日を調整しております。

詳細な事業計画書（収支5ヵ年・資金繰り3パターン・営業KPI・採用計画・感度分析の全23項目）を添付いたします。ぜひご面談の機会を賜りたく、お願い申し上げます。

株式会社KHD 代表取締役 菊池研太／ec-seller@kikuchi-hd.net"""

STRAT=[
 ["申込制度","新規開業資金（旧・新創業融資を2024年に一本化）。無担保・無保証人枠あり","事業開始後おおむね7年以内まで対象＝KHDの新規事業でOK"],
 ["申込額","設備600万＋運転500万＝計1,100万（自己資金300万投入・計1,400万）","C型(パート主軸)の必要資金1,356万を根拠に。『借りれるだけ』は根拠の積み上げで実現する"],
 ["期間(長く借りる)","制度上限＝設備20年以内・運転10年以内。希望提示は設備15年・運転10年","上限ベタ希望より現実的な長期提示が通りやすい。月返済を約10万円台に圧縮"],
 ["据置","据置5年以内の制度。希望2年（利用者積み上げ期は元金据置）","レセプト入金サイト＋立上げ赤字期を凌ぐ。依頼文に明記済み"],
 ["自己資金の見せ方","投入300万＋手元温存300万＝通帳で600万を示す","『全額突っ込まず手元を残す計画性』はプラス評価。投入比率21%"],
 ["金利メニュー","基準利率ベース。創業支援・雇用拡大の特利該当を面談で確認","雇用創出（施術者採用計画）は特利の材料になる"],
 ["経営者保証","面談で『経営者保証免除特例制度』の適用可否を質問（上乗せ金利あり）","法人借入で個人保証を外せるかは交渉価値あり"],
]

FILLIN=[
 ["作成日","申込日を記入（令和8年6月）"],
 ["自己資金","600万円保有（うち投入300万・手元温存300万）→通帳エビデンス"],
 ["設備資金の内訳","加盟金150・保証金30・研修費120・備品/機材50・中古車両100・物件取得/初期システム他150 ≒600万 ※本部見積で確定（[本部確認]欄）"],
 ["運転資金","500万（立上げ人件費・家賃・広告分担金・レセプト入金サイト2〜4ヶ月の立替）"],
 ["公庫からの借入","1,100万円（現様式の『所要−600万』を修正）"],
 ["客単価","利用者1人あたり月3.8万円（保険9割・非課税）※0_前提と一致させる"],
 ["売上見通し(創業当初)","利用者10名×3.8万＝38万円/月・利益▲50万円程度（立上げ期）"],
 ["売上見通し(1年後・軌道)","利用者60名前後×3.8万＝約190〜230万円/月・利益0〜30万円（7_年度別損益C型と整合させる）"],
 ["販売先・回収条件","後期高齢者医療広域連合等（保険者）／入金まで約2〜4ヶ月（本部確認中）"],
 ["お借入の状況","代表者個人の公庫借入（既往・上野）を正直に記入＝返済実績はプラス材料"],
 ["従業員","常勤役員1名・開業時 施術者1〜2名（パート中心）＋順次増員（2_人員構成と一致）"],
 ["許認可","施術所開設届 or 出張施術業務開始届（保健所）＝『申請中/予定』と記入"],
]

MADO=[
 ["窓口の原則","法人の申込は『本店登記地 or 事業所(新店舗)所在地』の管轄支店。個人=上野の実績があっても、法人×船橋店舗なら船橋支店管轄の可能性が高い","6/12に事業資金相談ダイヤル ☎0120-154-505 で確定（上野で受けてもらえるかも同時に確認）"],
 ["申込方法","公庫HPの『インターネット申込』（24時間・最速）→管轄支店に自動ルーティング","紙より速い。書類はPDFアップロード"],
 ["★タイミング（最重要）","『本部に金を払う前』に申し込む。支払済みの資金は融資対象外になるのが原則","順番＝①6/24 GO→本部へ『融資前提・支払いは実行後』と伝え契約書(案)＋見積入手→②即オンライン申込→③面談(約1〜2週間後)→④実行・着金(申込から3週間〜1ヶ月)→⑤着金後に加盟金支払い"],
 ["法人初回の追加書類","履歴事項全部証明書(登記簿)・定款・直近決算書(令7期)・代表者の資産負債/借入明細","KHD直近は小規模・債務超過▲42万→想定問答で先回り（下記）"],
 ["面談のコツ","数字は『計画(21タブ)⇔実績管理(13タブ)』の仕組みごと見せる＝回収可能性のアピール","『どう返すか』を聞かれる前にCF表で示す"],
]

DOCS=[
 "借入申込書（オンラインなら画面入力）",
 "創業計画書（公式様式・24タブの記入値で完成させる）＋別添：本事業計画書23タブ（公式様式に『自作計画書でも可』と明記あり）",
 "FC加盟契約書(案) or 加盟内定が分かる書面＋本部見積書（設備資金の根拠）",
 "法人：履歴事項全部証明書・定款・直近決算書",
 "代表者：自己資金の通帳（600万）・個人公庫の返済予定表（返済実績の証明）・本人確認書類",
 "資金繰り表（16タブ：レセプト入金サイト反映版）",
 "（あれば）狩野オーナーヒアリング記録・Indeed採用テスト結果＝計画の実証性",
]

QA=[
 ["KHDは債務超過▲42万では？","役員借入120万との相殺予定で実質資産超過。本業（不動産・医療コンサル）は継続収益あり。新事業は本計画のとおり保守前提で2年目黒字化（C型）"],
 ["FC本部に依存しすぎでは？","本部実績（65店・リピート98%）＋現役オーナー実地検証の二重確認。さらに自社の医師140件・ケアマネ網という本部に依らない集客資産を保有"],
 ["人は採れるのか？","加盟前にIndeedテストで有資格応募を確認済み。本部の採用支援（学校紹介等）＋パート主軸設計（現役オーナー実証の定着パターン）"],
 ["返済原資は？","C型で2年目黒字化・3年目累計黒字。据置2年の間に利用者60名超へ→月返済10万円台に対し営業CFで十分カバー（7・16タブ）"],
 ["なぜ1,100万も必要？","療養費の入金が2〜4ヶ月遅れるため、売上が立っても現金が入らない期間の運転資金が必須。逆にここを軽視する計画こそ危険と考えている"],
]

def creds():
    with open(TOKEN,"rb") as f:c=pickle.load(f)
    if c and c.expired and c.refresh_token:
        c.refresh(Request())
        with open(TOKEN,"wb") as f:pickle.dump(c,f)
    return c
def mr(g,r,c0,nc):return {"mergeCells":{"range":{"sheetId":g,"startRowIndex":r,"endRowIndex":r+1,"startColumnIndex":c0,"endColumnIndex":nc},"mergeType":"MERGE_ALL"}}
def fr(g,r,bg,fg,b,sz,wrap=True,h="LEFT"):return {"repeatCell":{"range":{"sheetId":g,"startRowIndex":r,"endRowIndex":r+1},"cell":{"userEnteredFormat":{"backgroundColor":bg,"textFormat":{"foregroundColor":fg,"bold":b,"fontSize":sz},"wrapStrategy":"WRAP" if wrap else "OVERFLOW_CELL","verticalAlignment":"MIDDLE","horizontalAlignment":h}},"fields":"userEnteredFormat"}}
def rh(g,r,px,r2=None):return {"updateDimensionProperties":{"range":{"sheetId":g,"dimension":"ROWS","startIndex":r,"endIndex":(r2 or r)+1},"properties":{"pixelSize":px},"fields":"pixelSize"}}
def cw(g,ws):return [{"updateDimensionProperties":{"range":{"sheetId":g,"dimension":"COLUMNS","startIndex":i,"endIndex":i+1},"properties":{"pixelSize":w},"fields":"pixelSize"}} for i,w in enumerate(ws)]

def main():
    svc=build("sheets","v4",credentials=creds(),cache_discovery=False)
    info=svc.spreadsheets().get(spreadsheetId=SID,fields="sheets.properties").execute()
    ex={s["properties"]["title"]:s["properties"]["sheetId"] for s in info["sheets"]}
    if TAB in ex:
        svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":[{"deleteSheet":{"sheetId":ex[TAB]}}]}).execute()
    NC=3
    res=svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":[{"addSheet":{"properties":{"title":TAB,"gridProperties":{"rowCount":90,"columnCount":NC,"frozenRowCount":2}}}}]}).execute()
    gid=res["replies"][0]["addSheet"]["properties"]["sheetId"]

    R=[];A=lambda r:R.append(r+[""]*(NC-len(r)))
    A(["🏦 公庫申込パック｜『借りれるだけ・長く』の交渉設計（法人初・新規開業資金）"])
    A(["戦略＝額は『C型の谷1,356万』を根拠に1,100万＋自己300万。期間は制度上限(設備20年/運転10年/据置5年)の内側で長期希望。★本部への支払いは必ず融資着金後。"])
    A([""])
    A(["■ ① 依頼文（カバーレター・このまま印刷/メール可）"]); sC=len(R)-1
    A([COVER]); rC=len(R)-1
    A([""])
    A(["■ ② 申込戦略（借りれるだけ・長く）"]); sS=len(R)-1
    A(["項目","内容","狙い・根拠"]); hS=len(R)-1
    s1=len(R)
    for x in STRAT: A(x)
    e1=len(R)
    A([""])
    A(["■ ③ 創業計画書（公式様式）の空欄に入れる値 ※Driveの様式の[要記入][本部確認][所要-600万]をこれで置換"]); sF=len(R)-1
    A(["欄","記入値"]); hF=len(R)-1
    s2=len(R)
    for x in FILLIN: A(x)
    e2=len(R)
    A([""])
    A(["■ ④ 窓口・申込タイミング（法人は初めて＝ここを間違えない）"]); sM=len(R)-1
    A(["論点","結論","アクション"]); hM=len(R)-1
    s3=len(R)
    for x in MADO: A(x)
    e3=len(R)
    A([""])
    A(["■ ⑤ 必要書類チェックリスト"]); sD=len(R)-1
    s4=len(R)
    for d in DOCS: A(["□ "+d])
    e4=len(R)
    A([""])
    A(["■ ⑥ 面談の想定問答（先回り）"]); sQ=len(R)-1
    A(["想定質問","回答の骨子"]); hQ=len(R)-1
    s5=len(R)
    for x in QA: A(x)
    e5=len(R)

    svc.spreadsheets().values().update(spreadsheetId=SID,range=f"'{TAB}'!A1",valueInputOption="USER_ENTERED",body={"values":R}).execute()

    reqs=cw(gid,[230,440,420])
    reqs+=[mr(gid,0,0,NC),fr(gid,0,REDD,WHT,True,14),rh(gid,0,42),
           mr(gid,1,0,NC),fr(gid,1,REDBG,REDD,True,10),rh(gid,1,40)]
    for s in (sC,sS,sF,sM,sD,sQ):
        reqs+=[mr(gid,s,0,NC),fr(gid,s,SEC,WHT,True,11),rh(gid,s,26)]
    # 依頼文：全幅マージ・上寄せ
    reqs+=[mr(gid,rC,0,NC),
           {"repeatCell":{"range":{"sheetId":gid,"startRowIndex":rC,"endRowIndex":rC+1},
            "cell":{"userEnteredFormat":{"backgroundColor":BLUEBG,"wrapStrategy":"WRAP","verticalAlignment":"TOP","textFormat":{"fontSize":10,"foregroundColor":INK}}},"fields":"userEnteredFormat"}},
           rh(gid,rC,640)]
    for h in (hS,hF,hM,hQ):
        reqs.append(fr(gid,h,RED,WHT,True,10,True,"CENTER"))
    for a,b,hh in [(s1,e1,52),(s2,e2,40),(s3,e3,64),(s5,e5,52)]:
        for i in range(a,b):
            reqs.append(fr(gid,i,CARD if (i-a)%2 else WHT,INK,False,10,True,"LEFT"))
        reqs.append(rh(gid,a,hh,b-1))
        reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":a,"endRowIndex":b,"startColumnIndex":0,"endColumnIndex":1},"cell":{"userEnteredFormat":{"textFormat":{"bold":True,"fontSize":10}}},"fields":"userEnteredFormat.textFormat"}})
    # タイミング行の★を強調
    reqs.append({"repeatCell":{"range":{"sheetId":gid,"startRowIndex":s3+2,"endRowIndex":s3+3},"cell":{"userEnteredFormat":{"backgroundColor":HI}},"fields":"userEnteredFormat.backgroundColor"}})
    # 書類チェック
    for i in range(s4,e4):
        reqs+=[mr(gid,i,0,NC),fr(gid,i,TIERA if (i-s4)%2 else WHT,INK,False,10),rh(gid,i,28)]
    reqs.append({"updateBorders":{"range":{"sheetId":gid,"startRowIndex":hS,"endRowIndex":e5,"startColumnIndex":0,"endColumnIndex":NC},
        "innerHorizontal":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}},
        "innerVertical":{"style":"SOLID","color":{"red":0.85,"green":0.83,"blue":0.79}}}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SID,body={"requests":reqs}).execute()
    print("DONE tab=",TAB,"gid=",gid)
    print(f"URL: https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}")

if __name__=="__main__":
    main()
