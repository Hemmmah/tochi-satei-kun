"""個別格差補正の適用。
比準法の核：各事例の単価に、ヘドニックβを使って「事例→査定対象」の特徴差分の補正を施す。

ln(adjusted) = ln(case_unit_price) + Σ β_i × (target_x_i - case_x_i)

確定方針（プラン §6.4）：被説明変数は ln(単価/㎡)、対数線形。
件数不足でβ辞書が空の場合は補正なしで返す（類似度ベース集約に降格）。
"""
import math
import pandas as pd

# 南向き判定（hedonic.py と同期）
SOUTH_FACING = {"南", "南東", "南西"}

# 補正対象の特徴量（hedonic.py の FEATURE_LABELS と整合）
CORRECTION_FEATURES = [
    "ln_area", "walk_min",
    "ln_shape", "ln_road_w", "D_south",
    "D_shidou", "D_fukuro", "D_fuseikei",
]

# 比準表での補正種別の分類（鑑定書様式に準拠）
# 標準化補正 = 画地・形状系（事例の個別性を標準画地に揃える）
# 地域格差 = 地域・街路・交通系（事例の地域性を査定地域に揃える）
HIJUN_GROUP = {
    "ln_shape":   "標準化補正",
    "D_fukuro":   "標準化補正",
    "D_fuseikei": "標準化補正",
    "ln_area":    "地域格差",
    "walk_min":   "地域格差",
    "ln_road_w":  "地域格差",
    "D_south":    "地域格差",
    "D_shidou":   "地域格差",
}


def _shape_index(area: float, kanguchi: float) -> float:
    """形状指数 = ln(間口²/面積)。
    値0付近で正方形、正で横長（帯）、負で縦長（旗竿）。
    """
    a = max(float(area), 1.0)
    k = max(float(kanguchi), 0.5)
    return 2 * math.log(k) - math.log(a)


def _target_feature_value(target: dict, feature: str) -> float:
    if feature == "ln_area":
        return math.log(target["面積(㎡)"])
    if feature == "walk_min":
        return float(target.get("最寄駅:距離(分)", 10))
    if feature == "ln_shape":
        kang = target.get("間口", 6.0)
        try:
            kang = float(kang)
        except (TypeError, ValueError):
            kang = 6.0
        return _shape_index(target["面積(㎡)"], kang)
    if feature == "ln_road_w":
        v = target.get("前面道路:幅員(m)", 5.0)
        try:
            v = float(v)
        except (TypeError, ValueError):
            v = 5.0
        return math.log(max(v, 1.0))
    if feature == "D_south":
        return 1.0 if str(target.get("前面道路:方位", "")) in SOUTH_FACING else 0.0
    if feature == "D_shidou":
        return 1.0 if target.get("前面道路:種類") == "私道" else 0.0
    if feature == "D_fukuro":
        return 1.0 if target.get("土地の形状") == "袋地" else 0.0
    if feature == "D_fuseikei":
        return 1.0 if target.get("土地の形状") == "不整形" else 0.0
    return 0.0


def _case_feature_value(row: pd.Series, feature: str) -> float:
    if feature == "ln_area":
        return math.log(row["area"])
    if feature == "walk_min":
        v = row.get("walk_min")
        return float(v) if pd.notna(v) else 10.0
    if feature == "ln_shape":
        kang = row.get("kanguchi")
        if pd.isna(kang) or kang is None:
            return 0.0
        try:
            return _shape_index(row["area"], float(kang))
        except (TypeError, ValueError):
            return 0.0
    if feature == "ln_road_w":
        v = row.get("road_width")
        if pd.isna(v) or v is None:
            return math.log(5.0)
        try:
            return math.log(max(float(v), 1.0))
        except (TypeError, ValueError):
            return math.log(5.0)
    if feature == "D_south":
        return 1.0 if str(row.get("road_dir", "")) in SOUTH_FACING else 0.0
    if feature == "D_shidou":
        return 1.0 if row.get("road_type") == "私道" else 0.0
    if feature == "D_fukuro":
        return 1.0 if row.get("shape") == "袋地" else 0.0
    if feature == "D_fuseikei":
        return 1.0 if row.get("shape") == "不整形" else 0.0
    return 0.0


def apply_correction(cases_df: pd.DataFrame, hedonic_result: dict, target: dict) -> pd.DataFrame:
    """各事例に個別格差補正を適用し、'corrected_unit_price' 列と各補正項目の寄与列を付与。

    補正なし（hedonic.ok=False）の場合は adjusted_unit_price をそのまま使う。
    """
    out = cases_df.copy()
    if not hedonic_result["ok"]:
        # 件数不足：補正なしで時点修正後単価をそのまま使う（類似度ベース集約に降格）
        if "adjusted_unit_price" in out.columns:
            out["corrected_unit_price"] = out["adjusted_unit_price"]
        else:
            out["corrected_unit_price"] = out["unit_price"]
        for feat in CORRECTION_FEATURES:
            out[f"correction_{feat}"] = 0.0
        out["correction_log_total"] = 0.0
        return out

    coef = hedonic_result["coefficients"]
    base_col = "adjusted_unit_price" if "adjusted_unit_price" in out.columns else "unit_price"
    log_corrections = []
    per_feature = {f: [] for f in CORRECTION_FEATURES}

    for _, row in out.iterrows():
        log_total = 0.0
        for feat in CORRECTION_FEATURES:
            if feat not in coef:
                per_feature[feat].append(0.0)
                continue
            beta = coef[feat]["beta"]
            tx = _target_feature_value(target, feat)
            cx = _case_feature_value(row, feat)
            contrib = beta * (tx - cx)
            per_feature[feat].append(contrib)
            log_total += contrib
        log_corrections.append(log_total)

    out["correction_log_total"] = log_corrections
    for feat, vals in per_feature.items():
        out[f"correction_{feat}"] = vals
    out["corrected_unit_price"] = out.apply(
        lambda r: r[base_col] * math.exp(r["correction_log_total"]), axis=1
    )
    return out


def hijun_correction_for_case(row: pd.Series, hedonic_result: dict, target: dict) -> dict:
    """事例1件について比準表用の補正係数を計算。

    Returns:
        {
          "事情補正": float,    # 通常1.0（取引事情なし）。"取引の事情等"記載があれば調整
          "事情補正_適用": bool, # False なら「100/-」表示
          "時点修正": float,    # (1+r)^年
          "建付減価": float,    # 土地のみ→1.0、適用なし
          "建付減価_適用": bool, # False なら「100/-」
          "標準化補正": float,  # exp(Σ β_標準化 × (target_x - case_x))
          "地域格差": float,    # exp(Σ β_地域 × (target_x - case_x))
          "試算値": float,      # 取引価格 × 全補正の積
        }
    """
    # 事情補正：「取引の事情等」が空なら通常取引（補正なし）
    jijo_raw = row.get("取引の事情等") if "取引の事情等" in row else None
    jijo_apply = bool(jijo_raw and not pd.isna(jijo_raw) and str(jijo_raw).strip())
    jijo_mult = 1.0  # 個別判断、現状は通常取引扱い

    # 時点修正：apply_time_adjustment 後の adjusted_unit_price から逆算
    base_price = float(row["unit_price"])
    if "adjusted_unit_price" in row and pd.notna(row["adjusted_unit_price"]):
        time_mult = float(row["adjusted_unit_price"]) / base_price if base_price > 0 else 1.0
    else:
        time_mult = 1.0

    # 建付減価：MLITは土地のみ抽出済 → 適用なし
    kentsuke_apply = False
    kentsuke_mult = 1.0

    # ヘドニック係数による補正を「標準化補正」「地域格差」に分類
    hyojunka_mult = 1.0
    chiiki_mult = 1.0
    if hedonic_result.get("ok"):
        coef = hedonic_result["coefficients"]
        for feat in CORRECTION_FEATURES:
            if feat not in coef:
                continue
            beta = coef[feat]["beta"]
            tx = _target_feature_value(target, feat)
            cx = _case_feature_value(row, feat)
            contrib = beta * (tx - cx)  # 対数空間での補正項
            mult = math.exp(contrib)
            group = HIJUN_GROUP.get(feat)
            if group == "標準化補正":
                hyojunka_mult *= mult
            elif group == "地域格差":
                chiiki_mult *= mult

    # 試算値 = 取引価格 × 全補正
    shisan = base_price * jijo_mult * time_mult * kentsuke_mult * hyojunka_mult * chiiki_mult

    return {
        "事情補正": jijo_mult,
        "事情補正_適用": jijo_apply,
        "時点修正": time_mult,
        "建付減価": kentsuke_mult,
        "建付減価_適用": kentsuke_apply,
        "標準化補正": hyojunka_mult,
        "地域格差": chiiki_mult,
        "試算値": shisan,
    }


def correction_breakdown(cases_df: pd.DataFrame, hedonic_result: dict) -> list:
    """業者用シート向けの「補正の内訳」を辞書のリストで返す。
    事例識別子は MLIT原本の case_no（行番号）を使う。
    """
    if not hedonic_result["ok"]:
        return []
    coef = hedonic_result["coefficients"]
    rows = []
    for _, r in cases_df.iterrows():
        case_no = r.get("case_no")
        entry = {
            "事例番号": int(case_no) if pd.notna(case_no) else 0,
            "district": r.get("district"),
            "area": r.get("area"),
        }
        for feat in CORRECTION_FEATURES:
            label = coef.get(feat, {}).get("label", feat)
            ratio = (math.exp(r.get(f"correction_{feat}", 0.0)) - 1.0) * 100
            entry[label] = round(ratio, 2)
        entry["合計補正率(%)"] = round(
            (math.exp(r.get("correction_log_total", 0.0)) - 1.0) * 100, 2
        )
        rows.append(entry)
    return rows


if __name__ == "__main__":
    from pathlib import Path
    from datetime import date
    import json
    from load_mlit import load_mlit_csv, load_koji_csv, load_kijun_csv
    from scope import scope_dataframe
    from similarity import compute_similarity, top_k
    from time_adjust import annual_rate_for_city, apply_time_adjustment
    from hedonic import fit_hedonic

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
    hed = fit_hedonic(adjusted)
    sim = compute_similarity(adjusted, target)
    top = top_k(sim, 5)
    corrected = apply_correction(top, hed, target)
    print(corrected[["district", "area", "adjusted_unit_price",
                     "correction_log_total", "corrected_unit_price"]])
