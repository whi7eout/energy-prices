# energy-prices

PHEVCalculator アプリ向けの価格マスターデータ置き場（公開情報のみ）。

## gasoline-prices.json

- 出典：資源エネルギー庁「石油製品価格調査」（給油所小売価格調査、週次）
  - 結果ページ: https://www.enecho.meti.go.jp/statistics/petroleum_and_lpgas/pl007/results.html
  - 毎週月曜日調査・水曜日公表。価格は現金価格（消費税込み）、単位は円/L。
  - 灯油（店頭）は公表値が円/18Lのため、円/Lへ換算している。
- 生成：`scripts/fetch_gasoline_prices.py` が結果詳細版 Excel を取得し、
  シート「公表資料（都道府県別）」の最新調査日分を JSON 化する。
- 更新：GitHub Actions（`.github/workflows/update-gasoline-prices.yml`）が
  毎週水曜 15:00 (JST) 目安に自動実行し、差分があればコミットする。

形式:

```json
{
  "surveyDate": "2026-11-02",
  "source": "資源エネルギー庁 石油製品価格調査",
  "prices": {
    "栃木県": { "regular": 169.5, "highOctane": 180.6, "diesel": 160.5, "kerosene": 145.33 }
  },
  "national": { "regular": 170.0, "highOctane": 180.9, "diesel": 159.5, "kerosene": 143.67 }
}
```
