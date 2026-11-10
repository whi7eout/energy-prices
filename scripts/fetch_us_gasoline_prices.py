#!/usr/bin/env python3
"""EIA（米エネルギー情報局）の週次小売ガソリン価格から us-gasoline-prices.json を生成する。

出典: https://www.eia.gov/petroleum/gasdiesel/
APIキー不要の公式ワークブックを使う:
- pswrgvwall.xls: 週次ガソリン価格（全米・PADD地域・一部州・都市）。
  レギュラー = ソースキーが EMM_EPMR_PTE_ のシート、プレミアム = EMM_EPMP_PTE_ のシート。
- psw18vwall.xls: 週次ディーゼル価格（超低硫黄）。ソースキーは EMD_EPD2DXL0_PTE_。
  対象地域は全米・PADD・カリフォルニア・西海岸（CA除く）のみ。
各系列の最新値（空欄でない最後の行）と、その週の日付（月曜）を採用する。
都市別系列（キー Y....）は収録しない。

出力 JSON: { surveyDate, source, sourceUrl, unit, fetchedAt,
             prices: {地域コード: {regular, premium, diesel?}}, national: {...} }
地域コード: US / CA,CO,FL,MA,MN,NY,OH,TX,WA / PADD1,PADD1A,PADD1B,PADD1C,PADD2,PADD3,PADD4,PADD5 / PADD5XCA
"""
import datetime
import io
import json
import sys
import urllib.request

import xlrd

GASOLINE_URL = "https://www.eia.gov/petroleum/gasdiesel/xls/pswrgvwall.xls"
DIESEL_URL = "https://www.eia.gov/petroleum/gasdiesel/xls/psw18vwall.xls"
HEADERS = {"User-Agent": "Mozilla/5.0 (energy-prices updater)"}

# ソースキーの地域サフィックス → 出力コード
AREA_CODES = {
    "NUS": "US",
    "SCA": "CA", "SCO": "CO", "SFL": "FL", "SMA": "MA", "SMN": "MN",
    "SNY": "NY", "SOH": "OH", "STX": "TX", "SWA": "WA",
    "R10": "PADD1", "R1X": "PADD1A", "R1Y": "PADD1B", "R1Z": "PADD1C",
    "R20": "PADD2", "R30": "PADD3", "R40": "PADD4", "R50": "PADD5",
    "R5XCA": "PADD5XCA",
}
AREA_ORDER = ["US", "CA", "CO", "FL", "MA", "MN", "NY", "OH", "TX", "WA",
              "PADD1", "PADD1A", "PADD1B", "PADD1C", "PADD2", "PADD3",
              "PADD4", "PADD5", "PADD5XCA"]


def fetch(url: str) -> bytes:
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=120) as resp:
        return resp.read()


def latest_series(book: xlrd.book.Book, key_prefix: str) -> dict:
    """ソースキーの接頭辞が一致するシートから {地域コード: (日付, 値)} を返す。"""
    result = {}
    for sheet in book.sheets():
        if sheet.nrows < 3 or sheet.ncols < 2:
            continue
        keys = [str(sheet.cell_value(1, c)) for c in range(sheet.ncols)]
        if not any(k.startswith(key_prefix) for k in keys):
            continue
        for c, key in enumerate(keys):
            if not key.startswith(key_prefix):
                continue
            suffix = key.split("_PTE_")[-1].replace("_DPG", "")
            code = AREA_CODES.get(suffix)
            if code is None:
                continue  # 都市別などは対象外
            for r in range(sheet.nrows - 1, 2, -1):
                value = sheet.cell_value(r, c)
                if isinstance(value, (int, float)) and value:
                    date = xlrd.xldate.xldate_as_datetime(sheet.cell_value(r, 0), book.datemode).date()
                    result[code] = (date, round(float(value), 3))
                    break
    return result


def main() -> None:
    out_path = sys.argv[1] if len(sys.argv) > 1 else "us-gasoline-prices.json"

    gasoline_book = xlrd.open_workbook(file_contents=fetch(GASOLINE_URL))
    regular = latest_series(gasoline_book, "EMM_EPMR_PTE_")
    premium = latest_series(gasoline_book, "EMM_EPMP_PTE_")
    diesel_book = xlrd.open_workbook(file_contents=fetch(DIESEL_URL))
    diesel = latest_series(diesel_book, "EMD_EPD2DXL0_PTE_")

    if "US" not in regular:
        raise SystemExit("U.S. regular gasoline series not found")

    survey_date = regular["US"][0]
    prices = {}
    for code in AREA_ORDER:
        entry = {}
        if code in regular:
            entry["regular"] = regular[code][1]
        if code in premium:
            entry["premium"] = premium[code][1]
        if code in diesel:
            entry["diesel"] = diesel[code][1]
        if entry:
            prices[code] = entry

    output = {
        "surveyDate": survey_date.isoformat(),
        "source": "U.S. Energy Information Administration (EIA), Weekly Retail Gasoline and Diesel Prices",
        "sourceUrl": "https://www.eia.gov/petroleum/gasdiesel/",
        "unit": "USD per gallon",
        "fetchedAt": datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds"),
        "prices": prices,
        "national": prices["US"],
    }
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
        f.write("\n")
    print(f"Wrote {out_path}: surveyDate={survey_date}, areas={len(prices)}, US regular={prices['US'].get('regular')}")


if __name__ == "__main__":
    main()
