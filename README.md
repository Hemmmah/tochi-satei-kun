# 土地価格査定クン (Tochi-Kasei-Kun) — White-Box Land Appraisal Skill for Claude Cowork

**土地価格査定クン** は、不動産仲介担当者・媒介査定担当者向けの **白箱AVM（Automated Valuation Model）** スキルです。MLIT（国土交通省）取引価格情報CSVと地価公示GeoJSONを入力に、**ヘドニック回帰係数を都度算定して個別格差補正を全開示** し、業者用シート（係数開示）と顧客用シート（流推方式準拠）の2シート構成xlsxを出力します。

媒介査定担当者が売主から物件を預かる際の **一次査定** を、説明可能な形で自動化します。

---

## What Makes It Different — 白箱AVMの核心

既存のAVMやネットの査定サービスは、**時点修正のみ係数を開示し、個別格差補正は黒箱**で出力するのが業界標準です。

土地価格査定クンは違います：

| 比較項目 | 既存AVM | 土地価格査定クン |
|---|---|---|
| 時点修正 | 係数開示 | 係数開示 |
| **個別格差補正** | **黒箱** | **全開示**（β、p値、標準誤差） |
| 統計手法 | 不明 | ヘドニック対数線形OLS（statsmodels） |
| 係数の更新 | 固定値 | **都度回帰で算定** |
| 補正項目 | 不明 | 8特徴量（面積・駅徒歩・道路幅員・形状指数・南向き・私道・袋地・不整形） |
| 適用範囲 | 全国一律ロジック | 市区町村単位、直近1年、直近事例で都度学習 |

「鑑定士の経験則を、ヘドニック係数として統計的に裏付ける」が本スキルの差別化メッセージです。

---

## What It Does

| Phase | Function |
|---|---|
| ① | MLIT取引価格情報CSVと地価公示GeoJSONを読み込み、市区町村単位・直近1年に絞り込み |
| ② | IQR法で外れ値除外、件数判定（15件未満なら降格） |
| ③ | 公示地価から直近1年の地価変動率を算定し、各事例単価を査定時点に補正 |
| ④ | ヘドニック対数線形OLSで都度回帰、8特徴量の係数を算定 |
| ⑤ | 重み付き類似度スコアで上位3事例を選定、ヘドニック係数で個別格差補正 |
| ⑥ | top3集約の中央値で査定価格、Q1/Q3で価格レンジを生成 |
| ⑦ | 公示地価との比較（地域標準価格チェック） |
| ⑧ | 業者用シート（係数全開示）＋顧客用シート（流推方式準拠）の2シート構成xlsx出力 |

---

## Who It's For

- 宅地建物取引士 / Licensed Real Estate Agents (*takken-shi*)
- 不動産鑑定士 / Certified Real Estate Appraisers
- 媒介査定担当者 / Brokerage Valuation Coordinators
- 地方銀行・信託銀行の不動産担当 / Bank Real Estate Officers

> **注：** 本スキルは『不動産の鑑定評価に関する法律』に基づく **鑑定評価書ではなく、宅建業者の媒介査定書** です。鑑定士本人が鑑定評価書を作成する用途では別の鑑定評価ソフトを使用してください（Phase 2 公開時の方針）。

---

## How to Install

```bash
# Cowork Marketplace に登録
/plugin marketplace add signal-yield/tochi-satei-kun
/plugin install tochi-satei-kun

# Python依存関係のインストール
pip install pandas statsmodels scikit-learn openpyxl
```

---

## How to Use

Claude Cowork でこんな風に発話するだけ：

```
土地価格査定クンを使って、港区麻布十番の120㎡の土地を査定して
```

スキルが起動し、必要なCSV（MLIT取引価格情報・地価公示GeoJSON）の取得手順を案内します。

物件JSONの最低限の必須項目：

```json
{
  "物件略号": "MIN001",
  "都道府県名": "東京都",
  "市区町村名": "港区",
  "地区名": "麻布十番",
  "面積(㎡)": 120,
  "最寄駅:名称": "麻布十番",
  "最寄駅:距離(分)": 7,
  "土地の形状": "整形",
  "間口": 8.0,
  "前面道路:方位": "南",
  "前面道路:種類": "公道",
  "前面道路:幅員(m)": 6.0,
  "都市計画": "第一種中高層住居専用地域",
  "建ぺい率(%)": 60,
  "容積率(%)": 200,
  "査定時点": "2026-05-19"
}
```

---

## Required Data

| データ | 形式 | 取得元 | 必須/任意 |
|---|---|---|---|
| MLIT 取引価格情報 | CSV（cp932） | https://www.reinfolib.mlit.go.jp/ | **必須** |
| 地価公示 | GeoJSON（国土数値情報 L01） | https://nlftp.mlit.go.jp/ksj/gml/datalist/KsjTmplt-L01.html | **必須** |
| 都道府県地価調査（基準地価） | GeoJSON（国土数値情報 L02） | 同上 L02 | 任意（補完用） |

スキル本体には実データを同梱していません（再配布リスク回避＋常に最新版を使うため、ユーザー自身でダウンロード）。

---

## Example Output

**港区麻布十番 120㎡ 整形地（2026-05-19 時点）の査定例：**

業者用シート：
- 査定価格：**1.2億円**（㎡単価 100万円／坪単価 330万円）
- 信頼度：高（n=42, 自由度調整済 R²=0.78, 期待符号と全整合）
- ヘドニック回帰サマリ：8特徴量すべての β、p値、標準誤差を全開示
- 比準表（鑑定書様式の2行式）：top3事例の取引価格・補正係数・試算値
- 2価格サマリ：採用査定価格 vs ヘドニック母集団予測値の乖離率

顧客用シート（売主提示用）：
- ですます調、係数・統計用語は完全非表示
- 主比準取引事例1件と価格レンジ
- 「面積はやや広め」「駅距離は同水準」など定性的記述

---

## Important Disclaimer

> **本スキルは媒介査定における一次スクリーニング支援ツールです。**
> 出力は『不動産の鑑定評価に関する法律』に基づく鑑定評価ではなく、宅建業者の媒介査定書として参考利用されるものです。最終的な売出価格・成約価格は媒介担当者の判断と現地確認に基づいて決定されるべきものです。

> **対象範囲：** 日本の宅地（更地・所有権のみ）。マンション、戸建（建物含む）、借地権・地上権・賃借権等の権利調整評価は対象外です。

---

## License

MIT License. Feel free to fork, adapt, and improve.
Pull requests and issue reports are welcome.

---

## Author

**松田幸一 / Koichi Matsuda**
不動産鑑定士（Certified Real Estate Appraiser, registered 2002）
Signal Yield Advisory
〒106-0032 東京都港区六本木３－１６－１２ 六本木KSビル５F

- 📧 [signalYield@gmail.com](mailto:signalYield@gmail.com)
- 📝 [note.com/matsudansyaku](https://note.com/matsudansyaku)
- 💼 [VisasQ プロフィール](https://expert.visasq.com/profile/#/)

---

## Related Skills

- **[重調クン (Jucho-kun)](https://github.com/pinotan2024-coder/jucho-kun)** — 重要事項説明書（35条書面）の事前調査自動化スキル。土地価格査定クンと併用することで、媒介取得から重説作成まで一気通貫の業務効率化が可能。
