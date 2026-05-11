"""ヘドニック回帰（対数線形 OLS）。
被説明変数: ln(adjusted_unit_price)  ※時点修正後の単価
特徴量: ln(面積), 駅徒歩分, D_私道, D_袋地, D_不整形

確定方針（プラン §2-1）：都度回帰、固定値ではない。
件数判定は scope.py 側で警告ログを付与するが、最終判定は呼び出し側。
"""
import math
import pandas as pd
import statsmodels.api as sm

MIN_SAMPLES_FOR_REGRESSION = 15

# 南向き判定：南、南東、南西を「南向き道路」とみなす（業界慣行）
SOUTH_FACING = {"南", "南東", "南西"}

# 特徴量 → SKILL.md/業者用シートで使う日本語名のマッピング
FEATURE_LABELS = {
    "ln_area": "面積（対数）",
    "ln_area_sq": "面積²（対数の二乗、規模逓減項）",
    "walk_min": "駅徒歩分",
    "ln_shape": "形状指数 ln(間口²/面積)",
    "ln_road_w": "道路幅員（対数）",
    "D_south": "南向き道路ダミー",
    "D_shidou": "私道ダミー",
    "D_fukuro": "袋地ダミー",
    "D_fuseikei": "不整形ダミー",
    "ln_district_mean": "地区平均単価（対数・ターゲット符号化）",
    "ln_station_mean": "最寄駅平均単価（対数・ターゲット符号化）",
    "const": "定数項",
}

# 地区／駅平均単価特徴量を有効にする最低サンプル数（地区／駅内）
DISTRICT_MEAN_MIN_SAMPLES = 3
STATION_MEAN_MIN_SAMPLES = 3


def _build_features(df: pd.DataFrame) -> pd.DataFrame:
    """特徴量行列を構築。MLIT既存項目を可能な限り活用。

    注：間口は単独では誤解を招く（広くても面積小さければ帯地）ため、
    形状指数 ln(間口²/面積) = 2*ln(間口) - ln(面積) として扱う。
    値0付近で正方形、正で横長（帯）、負で縦長（旗竿）。
    """
    X = pd.DataFrame(index=df.index)
    X["ln_area"] = df["area"].apply(math.log)
    # 面積²（規模逓減項）：大規模化で単価が下がる傾向を捕捉
    X["ln_area_sq"] = X["ln_area"] ** 2
    X["walk_min"] = df["walk_min"].fillna(
        df["walk_min"].median() if "walk_min" in df else 10
    )
    # 形状指数 = ln(間口²/面積) = 2·ln(間口) - ln(面積)
    if "kanguchi" in df.columns:
        kang = pd.to_numeric(df["kanguchi"], errors="coerce")
        kang_med = kang.median() if kang.notna().any() else 6.0
        kang = kang.fillna(kang_med).clip(lower=0.5)
        X["ln_shape"] = 2 * kang.apply(math.log) - X["ln_area"]
    else:
        X["ln_shape"] = 0.0
    # 道路幅員（対数、欠損は中央値で補完）
    if "road_width" in df.columns:
        rw = pd.to_numeric(df["road_width"], errors="coerce")
        rw_med = rw.median() if rw.notna().any() else 5.0
        rw = rw.fillna(rw_med).clip(lower=1.0)
        X["ln_road_w"] = rw.apply(math.log)
    else:
        X["ln_road_w"] = math.log(5.0)
    # 南向きダミー（南・南東・南西）
    X["D_south"] = df.get("road_dir", pd.Series([""] * len(df))).apply(
        lambda v: 1 if str(v) in SOUTH_FACING else 0
    ).astype(int)
    # 既存ダミー
    X["D_shidou"] = (df.get("road_type", "") == "私道").astype(int)
    X["D_fukuro"] = (df.get("shape", "") == "袋地").astype(int)
    X["D_fuseikei"] = (df.get("shape", "") == "不整形").astype(int)
    # 地区／駅平均単価（ターゲット符号化）：annotate_*_mean で事前に df に
    # ln_district_mean / ln_station_mean 列を付与しておく前提
    if "ln_district_mean" in df.columns:
        X["ln_district_mean"] = df["ln_district_mean"]
    if "ln_station_mean" in df.columns:
        X["ln_station_mean"] = df["ln_station_mean"]
    return X


def annotate_district_mean(df: pd.DataFrame) -> pd.DataFrame:
    """df に ln_district_mean 列を追加。

    地区別の単価平均を ln 取って格納。地区内 n >= DISTRICT_MEAN_MIN_SAMPLES なら
    同地区平均、それ未満は市区町村全体平均で代用。
    """
    out = df.copy()
    if "district" not in out.columns or "unit_price" not in out.columns or len(out) == 0:
        return out
    overall_mean = out["unit_price"].mean()
    counts = out.groupby("district")["unit_price"].transform("count")
    means = out.groupby("district")["unit_price"].transform("mean")
    district_price = means.where(counts >= DISTRICT_MEAN_MIN_SAMPLES, overall_mean)
    out["ln_district_mean"] = district_price.clip(lower=1.0).apply(math.log)
    return out


def annotate_station_mean(df: pd.DataFrame) -> pd.DataFrame:
    """df に ln_station_mean 列を追加。

    最寄駅別の単価平均を ln 取って格納。駅内 n >= STATION_MEAN_MIN_SAMPLES なら
    同駅平均、それ未満は市区町村全体平均で代用。
    """
    out = df.copy()
    if "station" not in out.columns or "unit_price" not in out.columns or len(out) == 0:
        return out
    overall_mean = out["unit_price"].mean()
    counts = out.groupby("station")["unit_price"].transform("count")
    means = out.groupby("station")["unit_price"].transform("mean")
    station_price = means.where(counts >= STATION_MEAN_MIN_SAMPLES, overall_mean)
    out["ln_station_mean"] = station_price.clip(lower=1.0).apply(math.log)
    return out


def fit_hedonic(df: pd.DataFrame) -> dict:
    """対数線形回帰を実行し、結果辞書を返す。

    Returns:
        {
          "ok": bool,
          "n": int,
          "r2": float, "adj_r2": float,
          "coefficients": {feature_name: {"beta": float, "se": float, "p": float, "label": str}},
          "skip_reason": str (if not ok),
        }
    """
    n = len(df)
    if n < MIN_SAMPLES_FOR_REGRESSION:
        return {
            "ok": False, "n": n, "r2": None, "adj_r2": None,
            "coefficients": {},
            "skip_reason": f"件数 {n} < {MIN_SAMPLES_FOR_REGRESSION}: 回帰スキップ",
        }
    if "ln_adjusted_unit_price" not in df.columns:
        y = df["unit_price"].apply(math.log)
    else:
        y = df["ln_adjusted_unit_price"]
    X = _build_features(df)
    X = sm.add_constant(X)
    try:
        model = sm.OLS(y, X).fit()
    except Exception as e:
        return {
            "ok": False, "n": n, "r2": None, "adj_r2": None,
            "coefficients": {},
            "skip_reason": f"OLS失敗: {e}",
        }

    coef = {}
    for name in X.columns:
        coef[name] = {
            "beta": float(model.params[name]),
            "se": float(model.bse[name]),
            "p": float(model.pvalues[name]),
            "label": FEATURE_LABELS.get(name, name),
        }
    return {
        "ok": True, "n": n,
        "r2": float(model.rsquared),
        "adj_r2": float(model.rsquared_adj),
        "coefficients": coef,
        "skip_reason": None,
    }


if __name__ == "__main__":
    from pathlib import Path
    from datetime import date
    import json
    from load_mlit import load_mlit_csv, load_koji_csv, load_kijun_csv
    from scope import scope_dataframe
    from time_adjust import annual_rate_for_city, apply_time_adjustment

    here = Path(__file__).parent.parent / "samples"
    df = load_mlit_csv(here / "sample_mlit.csv")
    koji = load_koji_csv(here / "sample_koji.csv")
    kijun = load_kijun_csv(here / "sample_kijun.csv")
    with open(here / "sample_property.json", encoding="utf-8") as f:
        target = json.load(f)
    asof = date(2025, 12, 1)
    scoped, _ = scope_dataframe(df, target, asof)
    rate = annual_rate_for_city(koji, kijun, target["市区町村名"])["rate"]
    adjusted = apply_time_adjustment(scoped, asof, rate)
    result = fit_hedonic(adjusted)
    print(f"ok={result['ok']}, n={result['n']}, R²={result['r2']:.3f}")
    for name, c in result["coefficients"].items():
        print(f"  {name:15s}: β={c['beta']:+.4f} (se={c['se']:.4f}, p={c['p']:.3f})")
