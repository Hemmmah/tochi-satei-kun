"""パイプライン全体のオーケストレータ。Claude が SKILL.md の指示でこれを呼ぶ。

使い方:
    python main.py <property.json> <mlit.csv> <koji.csv> <kijun.csv> [--out <dir>] [--asof YYYY-MM-DD]
"""
import argparse
import json
import math
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from load_mlit import load_mlit_csv, load_koji_auto, load_kijun_auto
from scope import scope_dataframe, filter_recent_for_comparison, DEFAULT_COMPARISON_MONTHS
from similarity import compute_similarity, top_k
from time_adjust import annual_rate_for_city, apply_time_adjustment
from hedonic import fit_hedonic, annotate_district_mean, annotate_station_mean
from correction import (apply_correction, correction_breakdown, hijun_correction_for_case,
                        compute_target_district_mean, compute_target_station_mean)
from aggregation import assess
from xlsx_writer import write_xlsx


SOUTH_FACING = {"南", "南東", "南西"}


def _hedonic_population_predict(hed: dict, target: dict) -> float:
    """係数辞書から対象物件の母集団予測値を算出。"""
    if not hed["ok"]:
        return None
    coef = hed["coefficients"]
    ln_pred = 0.0
    for name, c in coef.items():
        if name == "const":
            ln_pred += c["beta"]
            continue
        if name == "ln_area":
            x = math.log(target["面積(㎡)"])
        elif name == "ln_area_sq":
            x = math.log(target["面積(㎡)"]) ** 2
        elif name == "walk_min":
            x = float(target.get("最寄駅:距離(分)", 10))
        elif name == "ln_shape":
            kang = target.get("間口", 6.0)
            try:
                kang = float(kang)
            except (TypeError, ValueError):
                kang = 6.0
            area = float(target["面積(㎡)"])
            x = 2 * math.log(max(kang, 0.5)) - math.log(max(area, 1.0))
        elif name == "ln_road_w":
            v = target.get("前面道路:幅員(m)", 5.0)
            try:
                v = float(v)
            except (TypeError, ValueError):
                v = 5.0
            x = math.log(max(v, 1.0))
        elif name == "D_south":
            x = 1.0 if str(target.get("前面道路:方位", "")) in SOUTH_FACING else 0.0
        elif name == "D_shidou":
            x = 1.0 if target.get("前面道路:種類") == "私道" else 0.0
        elif name == "D_fukuro":
            x = 1.0 if target.get("土地の形状") == "袋地" else 0.0
        elif name == "D_fuseikei":
            x = 1.0 if target.get("土地の形状") == "不整形" else 0.0
        elif name == "ln_district_mean":
            v = target.get("_target_district_mean", 0.0)
            x = math.log(v) if v and v > 0 else 0.0
        elif name == "ln_station_mean":
            v = target.get("_target_station_mean", 0.0)
            x = math.log(v) if v and v > 0 else 0.0
        else:
            x = 0.0
        ln_pred += c["beta"] * x
    return math.exp(ln_pred)


def _interpolate_price_at_asof(rows, asof: date):
    """同一標準地の年次価格行から asof 時点の補間価格を返す（線形補間）。
    L01 GeoJSON は過去5年分（例：2022-01-01〜2026-01-01）の年次価格を保持。
    asof を挟む2点間で線形補間、範囲外は最も近い点の値を採用。
    """
    rows_sorted = sorted(rows, key=lambda r: r["price_date"])
    if len(rows_sorted) == 0:
        return None
    if len(rows_sorted) == 1:
        return float(rows_sorted[0]["price_per_sqm"])
    for i in range(len(rows_sorted) - 1):
        d1 = rows_sorted[i]["price_date"]
        d2 = rows_sorted[i + 1]["price_date"]
        if d1 <= asof <= d2:
            p1 = float(rows_sorted[i]["price_per_sqm"])
            p2 = float(rows_sorted[i + 1]["price_per_sqm"])
            span = (d2 - d1).days
            if span <= 0:
                return p1
            t = (asof - d1).days / span
            return p1 + (p2 - p1) * t
    # 外挿：最も近い点の価格
    if asof < rows_sorted[0]["price_date"]:
        return float(rows_sorted[0]["price_per_sqm"])
    return float(rows_sorted[-1]["price_per_sqm"])


def _standard_price_for_city(koji, kijun, city: str, asof: date,
                             target_district: str = None) -> dict:
    """標準地を地区一致優先で選定し、asof時点に時点補間して返す。

    Returns:
        {
            "standard_price_per_sqm": 補間後平均（円/㎡）,
            "source": "公示" or "公示+基準地",
            "selected_points": [{id, source, district, address, use, price_at_asof}],
            "selection_method": "district_match" | "city_average",
            "n_points": int,
        }
    """
    all_data = []
    for src_df, src_name in [(koji, "公示"), (kijun, "基準地")]:
        if src_df is None or src_df.empty:
            continue
        sub = src_df[src_df["city"] == city]
        if not sub.empty:
            all_data.append((src_name, sub))

    if not all_data:
        return {"standard_price_per_sqm": None, "source": None,
                "selected_points": [], "selection_method": "none", "n_points": 0}

    selected_points = []
    selection_method = "city_average"

    # ① 地区一致優先
    if target_district:
        for src_name, sub in all_data:
            district_match = sub[sub["district"] == target_district]
            if not district_match.empty:
                for std_id, group in district_match.groupby("標準地番号"):
                    rows = group.to_dict("records")
                    interp = _interpolate_price_at_asof(rows, asof)
                    if interp:
                        selected_points.append({
                            "id": std_id,
                            "source": src_name,
                            "district": rows[0]["district"],
                            "use": rows[0].get("use", ""),
                            "price_at_asof": interp,
                        })
                if selected_points:
                    selection_method = "district_match"

    # ② 地区一致がなければ市区町村全体
    if not selected_points:
        for src_name, sub in all_data:
            for std_id, group in sub.groupby("標準地番号"):
                rows = group.to_dict("records")
                interp = _interpolate_price_at_asof(rows, asof)
                if interp:
                    selected_points.append({
                        "id": std_id,
                        "source": src_name,
                        "district": rows[0]["district"],
                        "use": rows[0].get("use", ""),
                        "price_at_asof": interp,
                    })

    if not selected_points:
        return {"standard_price_per_sqm": None, "source": None,
                "selected_points": [], "selection_method": "none", "n_points": 0}

    avg = sum(p["price_at_asof"] for p in selected_points) / len(selected_points)
    sources = sorted(set(p["source"] for p in selected_points))
    return {
        "standard_price_per_sqm": avg,
        "source": "+".join(sources),
        "selected_points": selected_points,
        "selection_method": selection_method,
        "n_points": len(selected_points),
    }


def run_pipeline(property_path: str, mlit_path: str, koji_path: str, kijun_path: str,
                 out_dir: str = None, asof: date = None) -> Path:
    # 1. 入力読込
    with open(property_path, encoding="utf-8") as f:
        target = json.load(f)
    if asof is None:
        asof_str = target.get("査定時点")
        if asof_str:
            asof = datetime.strptime(asof_str, "%Y-%m-%d").date()
        else:
            asof = date.today()

    df = load_mlit_csv(mlit_path)
    koji = load_koji_auto(koji_path)
    kijun = load_kijun_auto(kijun_path)

    # 2. スコープ
    scoped, scope_log = scope_dataframe(df, target, asof)

    # 3. 時点修正（地区一致優先＋直近1年変動率）
    rate_info = annual_rate_for_city(
        koji, kijun, target["市区町村名"],
        target_district=target.get("地区名"),
        asof=asof,
    )
    adjusted = apply_time_adjustment(scoped, asof, rate_info["rate"])

    # 3b. 地区／最寄駅 平均単価をターゲット符号化として df に annotate（事例側・target 側共通）
    adjusted = annotate_district_mean(adjusted)
    adjusted = annotate_station_mean(adjusted)
    target["_target_district_mean"] = compute_target_district_mean(adjusted, target)
    target["_target_station_mean"] = compute_target_station_mean(adjusted, target)

    # 4. ヘドニック回帰（MLIT全期間で係数推定、n 最大化）
    hed = fit_hedonic(adjusted)

    # 5. 類似度 → top 3（取引事例比準は直近1.5年に絞って最新市場感を反映）
    recent = filter_recent_for_comparison(adjusted, asof, months=DEFAULT_COMPARISON_MONTHS)
    if len(recent) < 3:
        # 直近1.5年で3件未満なら全期間で代用（警告付き）
        recent = adjusted
        scope_log["warnings"].append(
            f"直近{DEFAULT_COMPARISON_MONTHS}ヶ月の事例が3件未満のため、比準事例選定も全期間から実施"
        )
    scope_log["comparison_recent_count"] = len(recent)
    sim = compute_similarity(recent, target)
    top_cases = top_k(sim, k=3)

    # 6. 個別格差補正
    corrected = apply_correction(top_cases, hed, target)
    breakdown = correction_breakdown(corrected, hed)

    # 6b. 比準表データ生成（鑑定書様式）
    hijun_rows = []
    for idx, (_, row) in enumerate(corrected.iterrows()):
        h = hijun_correction_for_case(row, hed, target)
        # 事例番号 = MLIT CSV原本の行番号（透明性のため）
        case_no = row.get("case_no")
        h["事例番号"] = int(case_no) if pd.notna(case_no) else (idx + 1)
        h["順位"] = "主比準" if idx == 0 else f"検証{idx}"  # 内部用ラベル
        h["取引価格"] = float(row["unit_price"])
        h["地区"] = row.get("district", "")
        h["取引時点"] = str(row.get("transaction_date", ""))
        h["面積"] = int(row["area"])
        hijun_rows.append(h)

    # 7. ヘドニック母集団予測
    hed_pred = _hedonic_population_predict(hed, target)

    # 9. 集約：top3 の中央値で査定価格、top3 の Q1/中央/Q3 でレンジ生成（案②）
    assessment = assess(corrected, target["面積(㎡)"])

    # 10. 地域標準価格チェック（地区一致優先＋時点補間）
    standard_check = _standard_price_for_city(
        koji, kijun, target["市区町村名"], asof,
        target_district=target.get("地区名"),
    )

    # 11. xlsx 出力
    out_dir = Path(out_dir) if out_dir else Path(property_path).parent / "output"
    fname = f"土地査定_{target.get('物件略号', 'NONAME')}_{asof.strftime('%Y%m%d')}.xlsx"
    out_path = out_dir / fname

    ctx = {
        "target": target,
        "asof": asof,
        "scope_log": scope_log,
        "rate_info": rate_info,
        "hedonic": hed,
        "cases": corrected,
        "breakdown": breakdown,
        "assess": assessment,
        "refs": {"hedonic_pred": hed_pred},
        "standard_check": standard_check,
        "hijun_rows": hijun_rows,
    }
    write_xlsx(ctx, out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("property", help="査定対象物件JSON")
    ap.add_argument("mlit", help="MLIT 取引価格情報CSV")
    ap.add_argument("koji", help="公示地価CSV")
    ap.add_argument("kijun", help="基準地価CSV")
    ap.add_argument("--out", default=None, help="出力ディレクトリ")
    ap.add_argument("--asof", default=None, help="査定時点 YYYY-MM-DD")
    args = ap.parse_args()
    asof = datetime.strptime(args.asof, "%Y-%m-%d").date() if args.asof else None
    out_path = run_pipeline(args.property, args.mlit, args.koji, args.kijun,
                            out_dir=args.out, asof=asof)
    print(f"[OK] 生成完了: {out_path}")


if __name__ == "__main__":
    main()
