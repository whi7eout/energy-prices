#!/usr/bin/env python3
"""EIA「Electric Power Monthly」Table 5.6.A から us-electricity-prices.json を生成する。

出典: https://www.eia.gov/electricity/monthly/ （APIキー不要の公式 xlsx）
Table 5.6.A は州別・部門別の平均単価（セント/kWh）。A列=州名、B列=家庭用の最新月。
地方（Census Division）行と注記行は州名マップに無いため自動的に除外される。

出力 JSON: { period, source, sourceUrl, unit, fetchedAt, nationalAverageCents,
             states: {州コード: セント/kWh} }
"""
import datetime
import json
import re
import sys
import urllib.request

import openpyxl

XLSX_URL = "https://www.eia.gov/electricity/monthly/xls/table_5_06_a.xlsx"
HEADERS = {"User-Agent": "Mozilla/5.0 (energy-prices updater)"}

STATE_CODES = {
    "Alabama": "AL", "Alaska": "AK", "Arizona": "AZ", "Arkansas": "AR",
    "California": "CA", "Colorado": "CO", "Connecticut": "CT", "Delaware": "DE",
    "District of Columbia": "DC", "Florida": "FL", "Georgia": "GA",
    "Hawaii": "HI", "Idaho": "ID", "Illinois": "IL", "Indiana": "IN",
    "Iowa": "IA", "Kansas": "KS", "Kentucky": "KY", "Louisiana": "LA",
    "Maine": "ME", "Maryland": "MD", "Massachusetts": "MA", "Michigan": "MI",
    "Minnesota": "MN", "Mississippi": "MS", "Missouri": "MO", "Montana": "MT",
    "Nebraska": "NE", "Nevada": "NV", "New Hampshire": "NH", "New Jersey": "NJ",
    "New Mexico": "NM", "New York": "NY", "North Carolina": "NC",
    "North Dakota": "ND", "Ohio": "OH", "Oklahoma": "OK", "Oregon": "OR",
    "Pennsylvania": "PA", "Rhode Island": "RI", "South Carolina": "SC",
    "South Dakota": "SD", "Tennessee": "TN", "Texas": "TX", "Utah": "UT",
    "Vermont": "VT", "Virginia": "VA", "Washington": "WA",
    "West Virginia": "WV", "Wisconsin": "WI", "Wyoming": "WY",
}


def main() -> None:
    out_path = sys.argv[1] if len(sys.argv) > 1 else "us-electricity-prices.json"

    req = urllib.request.Request(XLSX_URL, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=120) as resp:
        raw = resp.read()
    import io
    book = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
    sheet = book.active

    # タイトル行（2行目）から対象月を取り出す: "by State, August 2026 and 2025 (...)"
    title = str(sheet.cell(row=2, column=1).value or "")
    match = re.search(r"([A-Z][a-z]+ \d{4}) and \d{4}", title)
    period = match.group(1) if match else str(sheet.cell(row=4, column=2).value or "")

    states = {}
    national = None
    for row in sheet.iter_rows(min_row=5, max_col=2):
        name = row[0].value
        value = row[1].value
        if not isinstance(name, str) or not isinstance(value, (int, float)):
            continue
        name = name.strip()
        if name == "U.S. Total":
            national = round(float(value), 2)
        elif name in STATE_CODES:
            states[STATE_CODES[name]] = round(float(value), 2)

    if national is None or len(states) < 50:
        raise SystemExit(f"Parse failed: national={national}, states={len(states)}")

    output = {
        "period": period,
        "source": "U.S. Energy Information Administration (EIA), Electric Power Monthly, Table 5.6.A (Residential)",
        "sourceUrl": "https://www.eia.gov/electricity/monthly/",
        "unit": "cents per kWh",
        "fetchedAt": datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds"),
        "nationalAverageCents": national,
        "states": dict(sorted(states.items())),
    }
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
        f.write("\n")
    print(f"Wrote {out_path}: period={period}, states={len(states)}, national={national}")


if __name__ == "__main__":
    main()
