"""データ範囲スコープ規則。
市区町村絞り込み → 直近1年フィルタ → IQR外れ値除外。

**確定規則（2026-05-10 松田レビュー反映）**
- 地区＝市区町村単位（隣接市区町村への自動拡張は行わない・全面禁止）
- 期間＝直近1年（取引事例比較法の業界慣行）
- 件数15件未満ならヘドニック回帰スキップ（呼び出し側が判断）
- 件数が極端に少ない場合は無理に査定価格を出さず、取れた事例を類似性順に提示する運用
"""
from datetime import date
import pandas as pd

DEFAULT_PERIOD_YEARS = 1  # 取引事例比較法の業界慣行：取引事例の収集は直近1年以内を原則
MIN_COUNT = 15  # この件数を下回ると hedonic.py 側でスキップ判定


def filter_period(df: pd.DataFrame, asof: date, years: int = DEFAULT_PERIOD_YEARS) -> pd.DataFrame:
    """asof から years 年前までの取引に絞る。"""
    cutoff = date(asof.year - years, asof.month, 1)
    return df[df["transaction_date"] >= cutoff].copy()


def filter_iqr(df: pd.DataFrame, col: str = "unit_price", k: float = 1.5) -> pd.DataFrame:
    """IQR 法で外れ値を除外。"""
    if len(df) < 4:
        return df
    q1 = df[col].quantile(0.25)
    q3 = df[col].quantile(0.75)
    iqr = q3 - q1
    lo = q1 - k * iqr
    hi = q3 + k * iqr
    return df[(df[col] >= lo) & (df[col] <= hi)].copy()


def scope_dataframe(df: pd.DataFrame, target: dict, asof: date) -> tuple:
    """物件 target に対してスコープ規則を適用し、(scoped_df, scope_log) を返す。

    Returns:
        scoped_df: 絞り込み後 DataFrame
        scope_log: dict — 件数・警告フラグ
    """
    log = {
        "target_city": target["市区町村名"],
        "expanded_to": [],  # 後方互換のため空配列を保持（隣接拡張は廃止）
        "period_years": DEFAULT_PERIOD_YEARS,
        "iqr_removed": 0,
        "final_count": 0,
        "warnings": [],
    }
    target_city = target["市区町村名"]

    # ① 市区町村絞り込み（隣接拡張なし）
    sub = df[df["city"] == target_city].copy()
    sub = filter_period(sub, asof)

    # ② IQR 外れ値除外
    before_iqr = len(sub)
    sub = filter_iqr(sub, "unit_price")
    log["iqr_removed"] = before_iqr - len(sub)

    log["final_count"] = len(sub)
    if len(sub) < MIN_COUNT:
        log["warnings"].append(
            f"件数 {len(sub)} 件 < 最低 {MIN_COUNT} 件: ヘドニック回帰スキップ・類似度ベース集約に降格"
        )

    return sub.reset_index(drop=True), log


if __name__ == "__main__":
    from pathlib import Path
    import json
    from load_mlit import load_mlit_csv

    here = Path(__file__).parent.parent / "samples"
    df = load_mlit_csv(here / "sample_mlit.csv")
    with open(here / "sample_property.json", encoding="utf-8") as f:
        target = json.load(f)
    asof = date(2025, 12, 1)
    scoped, log = scope_dataframe(df, target, asof)
    print(f"scoped: {len(scoped)} rows")
    print(f"log: {log}")
