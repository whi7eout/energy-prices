#!/usr/bin/env python3
"""資源エネルギー庁「石油製品価格調査」の週次Excelから gasoline-prices.json を生成する。

出典ページ: https://www.enecho.meti.go.jp/statistics/petroleum_and_lpgas/pl007/results.html
- 結果ページから「結果詳細版」Excel（/xlsx/YYMMDD.xlsx、都道府県別シート付き）へのリンクを辿る。
- シート「公表資料（都道府県別）」の構成:
    5行目: 油種ヘッダー（C=ハイオク, E=レギュラー, G=軽油店頭, I=灯油店頭(円/18L), K=灯油配達）
    6行目: 油種ごとに (前回調査日, 最新調査日) のペアが並ぶ。最新列は D/F/H/J/L。
    7行目〜: B列に行名（都道府県は県名のみ・スペース区切り、地域局行と全国行が混在）。
- 灯油店頭は円/18Lで公表されるため、円/Lへ換算する（小数2位）。
出力 JSON: { surveyDate, source, prices: {都道府県: {regular, highOctane, diesel, kerosene}}, national: {...} }
"""
import datetime
import json
import re
import sys
import urllib.request

RESULTS_URL = "https://www.enecho.meti.go.jp/statistics/petroleum_and_lpgas/pl007/results.html"
BASE = "https://www.enecho.meti.go.jp"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en;q=0.9",
    "Referer": "https://www.enecho.meti.go.jp/",
}

PREFECTURES = [
    "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
    "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
    "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県",
    "岐阜県", "静岡県", "愛知県", "三重県",
    "滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県",
    "鳥取県", "島根県", "岡山県", "広島県", "山口県",
    "徳島県", "香川県", "愛媛県", "高知県",
    "福岡県", "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県",
]


def fetch(url: str) -> bytes:
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read()


def canonical_prefecture(raw: str) -> str | None:
    """シートの行名（例: '栃  木', '北 海 道 局', '東  京'）を正式名へ寄せる。該当なしは None。"""
    name = re.sub(r"[\s\u3000]+", "", raw)
    if name.endswith("局"):
        name = name[:-1]
    if name == "北海道":
        return "北海道"
    if name == "東京":
        return "東京都"
    if name in ("京都", "大阪"):
        return name + "府"
    cand = name + "県"
    return cand if cand in PREFECTURES else None


def find_detail_xlsx_url(html: str) -> str:
    """結果詳細版（都道府県別シートを持つ単週ファイル）のURLを返す。"""
    hrefs = re.findall(r'href="([^"]+\.xlsx)"', html)
    detail = [h for h in hrefs if re.search(r"/xlsx/\d{6}\.xlsx$", h)]
    if not detail:
        raise RuntimeError("結果詳細版Excelへのリンクが見つかりませんでした")
    return BASE + detail[0] if detail[0].startswith("/") else detail[0]


def parse_workbook(data: bytes) -> dict:
    import io
    import openpyxl

    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    ws = wb["公表資料（都道府県別）"]

    # 6行目: 油種ペアの2列目（D/F/H/J/L）が最新調査日。全て同じ日のはずだが最大値を採用する。
    latest_cols = {"highOctane": 4, "regular": 6, "diesel": 8, "kerosene": 10}
    dates = []
    for col in latest_cols.values():
        v = ws.cell(row=6, column=col).value
        if isinstance(v, datetime.datetime):
            dates.append(v.date())
    if not dates:
        raise RuntimeError("調査日（6行目）が読み取れませんでした")
    survey_date = max(dates)

    def price(row: int, col: int, per18l: bool = False):
        v = ws.cell(row=row, column=col).value
        if not isinstance(v, (int, float)):
            return None
        return round(v / 18.0, 2) if per18l else float(v)

    prices: dict[str, dict] = {}
    national = None
    for row in range(7, ws.max_row + 1):
        label = ws.cell(row=row, column=2).value
        if not isinstance(label, str):
            continue
        entry = {
            "regular": price(row, 6),
            "highOctane": price(row, 4),
            "diesel": price(row, 8),
            "kerosene": price(row, 10, per18l=True),
        }
        if any(v is None for v in entry.values()):
            continue
        if re.sub(r"[\s\u3000]+", "", label) == "全国":
            national = entry
            continue
        pref = canonical_prefecture(label)
        if pref:
            prices[pref] = entry

    missing = [p for p in PREFECTURES if p not in prices]
    if missing:
        raise RuntimeError(f"都道府県が不足しています: {missing}")
    if national is None:
        raise RuntimeError("全国行が見つかりませんでした")

    return {
        "surveyDate": survey_date.isoformat(),
        "source": "資源エネルギー庁 石油製品価格調査",
        "prices": {p: prices[p] for p in PREFECTURES},
        "national": national,
    }


def main(out_path: str = "gasoline-prices.json") -> int:
    html = fetch(RESULTS_URL).decode("utf-8", "replace")
    xlsx_url = find_detail_xlsx_url(html)
    print(f"Excel: {xlsx_url}", file=sys.stderr)
    result = parse_workbook(fetch(xlsx_url))
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
        f.write("\n")
    print(f"surveyDate={result['surveyDate']} prefectures={len(result['prices'])} "
          f"栃木 regular={result['prices']['栃木県']['regular']}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "gasoline-prices.json"))
