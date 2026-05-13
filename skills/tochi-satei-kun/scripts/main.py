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
                        hijun_breakdown_detail,
                        compute_target_district_mean, compute_target_station_mean)
from aggregation import assess
from xlsx_writer import write_xlsx


from hedonic import DIR_SCORE
SOUTH_FACING = {"南", "南東", "南西"}

# 公示番号の市区町村コード → 短縮名
_CITY_CODE_TO_SHORT = {
    "13101": "千代田", "13102": "中央", "13103": "港", "13104": "新宿",
    "13105": "文京", "13106": "台東", "13107": "墨田", "13108": "江東",
    "13109": "品川", "13110": "目黒", "13111": "大田", "13112": "世田谷",
    "13113": "渋谷", "13114": "中野", "13115": "杉並", "13116": "豊島",
    "13117": "北", "13118": "荒川", "13119": "板橋", "13120": "練馬",
    "13121": "足立", "13122": "葛飾", "13123": "江戸川",
}


def _short_koji_id(std_id: str) -> str:
    """公示番号を「市区町村-連番」形式に短縮。例: '13112-000-050' → '世田谷-50'"""
    if not std_id:
        return ""
    parts = str(std_id).split("-")
    if len(parts) < 3:
        return str(std_id)
    city_short = _CITY_CODE_TO_SHORT.get(parts[0], parts[0])
    try:
        n = int(parts[2])
    except (TypeError, ValueError):
        n = parts[2]
    return f"{city_short}-{n}"


# 用途地域カテゴリ判定（公示地点と対象物件のマッチング用）
_ZONING_CATEGORIES = {
    "低専": ["低専", "低住", "1種低層", "2種低層", "１低", "２低", "1低", "2低"],
    "中高": ["中高", "1種中高", "2種中高", "1中", "2中", "１中", "２中"],
    "住居": ["1住居", "2住居", "１住居", "２住居", "準住居"],
    "近商": ["近隣商業", "近商"],
    "商業": ["商業"],
    "準工": ["準工業", "準工"],
    "工業": ["工業", "工専"],
}


def _zoning_category(z: str) -> str:
    """用途地域文字列をカテゴリに正規化（マッチング比較用）。"""
    if not z:
        return ""
    z = str(z).strip()
    for cat, names in _ZONING_CATEGORIES.items():
        if any(n in z for n in names):
            return cat
    return ""


def _normalize_chome(s: str) -> str:
    """丁目数字（全角・漢数字を含む）を半角数字に正規化。"""
    if not s:
        return ""
    s = str(s).replace("丁目", "")
    zen2han = str.maketrans("０１２３４５６７８９", "0123456789")
    s = s.translate(zen2han)
    kanji = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5",
             "六": "6", "七": "7", "八": "8", "九": "9", "十": "10"}
    for k, v in kanji.items():
        s = s.replace(k, v)
    # 末尾数字のみ残す
    import re
    m = re.search(r"(\d+)", s)
    return m.group(1) if m else ""


def _score_koji_point(pt: dict, target: dict) -> float:
    """公示地点と対象物件の類似度スコア（0〜1、高いほど類似）。
    重み: 用途地域カテゴリ一致 0.4 / 容積率近さ 0.3 / 丁目一致 0.3。
    """
    score = 0.0
    # ① 用途地域カテゴリ一致
    t_cat = _zoning_category(target.get("都市計画", ""))
    p_cat = _zoning_category(pt.get("zoning", ""))
    if t_cat and p_cat and t_cat == p_cat:
        score += 0.4

    # ② 容積率の近さ
    try:
        t_far = float(target.get("容積率(%)", 200) or 200)
        p_far_raw = pt.get("floor_area_ratio")
        if p_far_raw not in (None, "", "_"):
            p_far = float(p_far_raw)
            if t_far > 0 and p_far > 0:
                diff_ratio = abs(t_far - p_far) / max(t_far, p_far)
                score += 0.3 * (1.0 - min(diff_ratio, 1.0))
    except (TypeError, ValueError):
        pass

    # ③ 丁目一致
    t_chome = _normalize_chome(target.get("丁目", ""))
    p_chome = ""
    addr = str(pt.get("address", ""))
    if "丁目" in addr:
        before = addr.split("丁目")[0]
        # 「赤堤５」「赤堤五」「赤堤5」の末尾数字を取得
        p_chome = _normalize_chome(before[-3:])
    if t_chome and p_chome and t_chome == p_chome:
        score += 0.3

    return score


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
        elif name == "ln_far":
            v = target.get("容積率(%)", 200)
            try:
                v = float(v)
            except (TypeError, ValueError):
                v = 200.0
            x = math.log(max(v, 1.0))
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
        elif name == "dir_score":
            x = float(DIR_SCORE.get(str(target.get("前面道路:方位", "")).strip(), 0))
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
                             target_district: str = None,
                             target: dict = None) -> dict:
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
                            "address": rows[0].get("address", ""),
                            "area_sqm": rows[0].get("area_sqm"),
                            "use_detail": rows[0].get("use_detail", ""),
                            "road_type": rows[0].get("road_type", ""),
                            "road_dir": rows[0].get("road_dir", ""),
                            "road_width": rows[0].get("road_width"),
                            "station": rows[0].get("station", ""),
                            "station_dist_m": rows[0].get("station_dist_m"),
                            "zoning": rows[0].get("zoning", ""),
                            "building_coverage": rows[0].get("building_coverage"),
                            "floor_area_ratio": rows[0].get("floor_area_ratio"),
                            "frontage_ratio": rows[0].get("frontage_ratio"),
                            "depth_ratio": rows[0].get("depth_ratio"),
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
                        "address": rows[0].get("address", ""),
                        "area_sqm": rows[0].get("area_sqm"),
                        "use_detail": rows[0].get("use_detail", ""),
                        "road_type": rows[0].get("road_type", ""),
                        "road_dir": rows[0].get("road_dir", ""),
                        "road_width": rows[0].get("road_width"),
                        "station": rows[0].get("station", ""),
                        "station_dist_m": rows[0].get("station_dist_m"),
                        "zoning": rows[0].get("zoning", ""),
                        "building_coverage": rows[0].get("building_coverage"),
                        "floor_area_ratio": rows[0].get("floor_area_ratio"),
                    })

    if not selected_points:
        return {"standard_price_per_sqm": None, "source": None,
                "selected_points": [], "selection_method": "none", "n_points": 0,
                "label": ""}

    # 1地点に絞る：対象物件と最も類似する地点を選定（用途地域+容積率+丁目スコアリング）
    if target is not None and len(selected_points) > 1:
        scored = [(p, _score_koji_point(p, target)) for p in selected_points]
        # スコア降順、同点なら price 中央値に近い順
        prices = sorted(p["price_at_asof"] for p in selected_points)
        median_p = prices[len(prices) // 2]
        scored.sort(key=lambda x: (-x[1], abs(x[0]["price_at_asof"] - median_p)))
        best = scored[0][0]
        best_score = scored[0][1]
        # スコア記録（業者用シートでの透明性のため）
        best = dict(best)
        best["similarity_score"] = best_score
        selected_points = [best]

    avg = selected_points[0]["price_at_asof"]
    sources = sorted(set(p["source"] for p in selected_points))
    return {
        "standard_price_per_sqm": avg,
        "source": "+".join(sources),
        "selected_points": selected_points,
        "selection_method": selection_method,
        "n_points": len(selected_points),
        "label": _label_for_standard_points(selected_points),
    }


def _label_for_standard_points(points: list) -> str:
    """選定された公示標準地のラベル生成（1地点なら番号、複数なら地点数）。
    例：「赤堤（13112-000-015）」「赤堤3地点平均」「3地点平均」
    """
    if not points:
        return ""
    used_ids = sorted(set(p.get("id", "") for p in points if p.get("id")))
    used_districts = sorted(set(p.get("district", "") for p in points if p.get("district")))
    if len(used_ids) == 1:
        short_id = _short_koji_id(used_ids[0])
        if used_districts:
            return f"{used_districts[0]}（{short_id}）"
        return f"標準地 {short_id}"
    if len(used_districts) == 1:
        return f"{used_districts[0]}{len(used_ids)}地点平均"
    if used_districts:
        return f"{len(used_ids)}地点平均（{'、'.join(used_districts)}）"
    return f"{len(used_ids)}地点平均"


def _compute_koji_timeseries(koji, city: str, district: str = None,
                              selected_ids: list = None) -> dict:
    """時点修正に使用した公示標準地の年次価格推移を返す（時点修正と整合）。

    selected_ids が指定されればその標準地のみで集計、なければ地区一致 → 市区町村平均で集計。

    Returns:
        {
          "data": [{"year": int, "price": float}, ...],
          "selected_ids": [str, ...],
          "label": "赤堤3地点平均" or "13112-000-015 単独" など
        }
    """
    empty = {"data": [], "selected_ids": [], "label": ""}
    if koji is None or koji.empty:
        return empty
    matched = koji[koji["city"] == city]
    if selected_ids and "標準地番号" in matched.columns:
        # 時点修正で使った標準地に絞る
        matched = matched[matched["標準地番号"].isin(selected_ids)]
    elif district:
        d_match = matched[matched["district"] == district]
        if len(d_match) > 0:
            matched = d_match
    if "price_date" not in matched.columns or "price_per_sqm" not in matched.columns:
        return empty
    m = matched.dropna(subset=["price_date", "price_per_sqm"]).copy()
    m["year"] = m["price_date"].apply(lambda d: d.year if d else None)
    yearly = m.dropna(subset=["year"]).groupby("year")["price_per_sqm"].mean().reset_index()
    yearly = yearly.sort_values("year")
    data = [{"year": int(row["year"]), "price": float(row["price_per_sqm"])}
            for _, row in yearly.iterrows()]
    used_ids = sorted(matched["標準地番号"].dropna().unique().tolist()) if "標準地番号" in matched.columns else []
    used_districts = sorted(matched["district"].dropna().unique().tolist()) if "district" in matched.columns else []
    # ラベル生成（短縮 ID を使用：13112-000-050 → 世田谷-50）
    if len(used_ids) == 1:
        label = _short_koji_id(used_ids[0])
    elif used_districts:
        label = f"{used_districts[0]}{len(used_ids)}地点平均" if len(used_districts) == 1 else f"{len(used_ids)}地点平均"
    else:
        label = f"{len(used_ids)}地点平均"
    return {"data": data, "selected_ids": used_ids, "label": label}


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
    # 個別格差（角地・方位・不整形）の係数効果を target vs each case で計算
    import math as _m
    def _kobetsu_pct(beta, tx_val, cx_val):
        if beta is None:
            return 0.0
        contrib = float(beta) * (float(tx_val) - float(cx_val))
        return (_m.exp(contrib) - 1.0) * 100

    coef = hed.get("coefficients", {}) if hed.get("ok") else {}
    target_dir_score = float(DIR_SCORE.get(str(target.get("前面道路:方位", "")).strip(), 0))
    target_fusei = 1.0 if target.get("土地の形状") == "不整形" else 0.0
    # 角地補正率：MLIT データに角地情報が無いためヘドニックで推定不能。
    # 白箱ポリシーに従い、業者が明示的に入力した値のみ採用（自動デフォルト無し）。
    kado_explicit = target.get("角地補正率(%)")
    if kado_explicit is None:
        target_kado = 0.0
    else:
        try:
            target_kado = float(kado_explicit)
        except (TypeError, ValueError):
            target_kado = 0.0

    hijun_rows = []
    hijun_detail_rows = []
    for idx, (_, row) in enumerate(corrected.iterrows()):
        h = hijun_correction_for_case(row, hed, target)
        # 個別格差
        case_dir_score = float(DIR_SCORE.get(str(row.get("road_dir", "")).strip(), 0))
        case_fusei = 1.0 if row.get("shape") == "不整形" else 0.0
        h["個別格差_角地"] = target_kado
        h["個別格差_方位"] = _kobetsu_pct(
            coef.get("dir_score", {}).get("beta") if "dir_score" in coef else None,
            target_dir_score, case_dir_score)
        h["個別格差_不整形"] = _kobetsu_pct(
            coef.get("D_fuseikei", {}).get("beta") if "D_fuseikei" in coef else None,
            target_fusei, case_fusei)
        # 総和（積×100、Excel 関数と整合）
        f = (1 + h["個別格差_角地"]/100) * (1 + h["個別格差_方位"]/100) * (1 + h["個別格差_不整形"]/100)
        h["個別格差_総和_pct"] = (f - 1.0) * 100  # 例: +5.0% なら 5.0
        h["個別格差_総和_factor"] = f * 100        # Excel 表示用: 例 105.0
        h["案件査定価格"] = float(h["試算値"]) * f  # 試算値 × 総和/100
        # 事例番号 = MLIT CSV原本の行番号（透明性のため）
        case_no = row.get("case_no")
        h["事例番号"] = int(case_no) if pd.notna(case_no) else (idx + 1)
        h["順位"] = "規範性の高い事例" if idx == 0 else f"類似事例{['②','③','④','⑤'][idx-1] if idx-1 < 4 else idx+1}"  # 内部用ラベル
        h["取引価格"] = float(row["unit_price"])
        h["地区"] = row.get("district", "")
        h["取引時点"] = str(row.get("transaction_date", ""))
        h["取引四半期"] = row.get("transaction_quarter_str", "") or str(row.get("transaction_date", ""))
        h["面積"] = int(row["area"])
        # 取引事例の概要表用の追加属性
        h["最寄駅"] = row.get("station", "")
        h["駅距離"] = row.get("walk_min", "")
        h["道路種別"] = row.get("road_type", "")
        h["道路幅員"] = row.get("road_width", "")
        h["方位"] = row.get("road_dir", "")
        h["形状"] = row.get("shape", "")
        h["用途地域"] = row.get("zoning", row.get("city_planning", ""))
        h["容積率_pct"] = row.get("floor_area_ratio", "")
        hijun_rows.append(h)
        # 詳細内訳（補修正率と地域格差率 表用）
        detail = hijun_breakdown_detail(row, hed, target)
        detail["記号"] = chr(ord("A") + idx)  # A, B, C
        detail["事例番号"] = h["事例番号"]
        hijun_detail_rows.append(detail)

    # 7. ヘドニック母集団予測
    hed_pred = _hedonic_population_predict(hed, target)

    # 9. 集約：top3 の中央値で査定価格、top3 の Q1/中央/Q3 でレンジ生成（案②）
    assessment = assess(corrected, target["面積(㎡)"])

    # 10. 地域標準価格チェック（地区一致優先＋時点補間、1地点に絞る）
    standard_check = _standard_price_for_city(
        koji, kijun, target["市区町村名"], asof,
        target_district=target.get("地区名"),
        target=target,
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
        "hijun_detail_rows": hijun_detail_rows,
        "koji_timeseries": _compute_koji_timeseries(
            koji, target["市区町村名"], target.get("地区名"),
            selected_ids=[pt["id"] for pt in standard_check.get("selected_points", [])]
        ),
        "adjusted_full": adjusted,  # 散布図用：時点修正後の全事例
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
