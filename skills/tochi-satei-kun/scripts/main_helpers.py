"""main.py から切り出したヘルパー関数群（v1.2.9、Cowork 配布層の truncate 回避）。

main.py 本体を 20KB 未満に圧縮するため、koji/kijun の標準価格計算・座標補正・
公示番号変換・時系列補間等の補助関数を本モジュールに移動した。

依存：load_mlit, scope, time_adjust, hedonic, pandas, math
"""
import math
from datetime import date

import pandas as pd

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


