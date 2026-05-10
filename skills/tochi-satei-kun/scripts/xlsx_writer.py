"""2シート構成の xlsx 生成。
業者用シート（係数全開示）+ 顧客用シート（流推方式準拠、禁止語チェック）。

確定方針（プラン §7）：
- 業者用：n, R², 全β、補正内訳、参考値、警告
- 顧客用：ですます調、禁止語ブロック
"""
import math
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from forbidden_words import assert_clean

# ===== スタイル =====
TITLE_FONT = Font(name="游ゴシック", size=14, bold=True, color="FFFFFF")
TITLE_FILL = PatternFill("solid", fgColor="2F5496")
SECTION_FONT = Font(name="游ゴシック", size=11, bold=True, color="FFFFFF")
SECTION_FILL = PatternFill("solid", fgColor="4472C4")
LABEL_FONT = Font(name="游ゴシック", size=10, bold=True)
VALUE_FONT = Font(name="游ゴシック", size=10)
BIG_VALUE_FONT = Font(name="游ゴシック", size=18, bold=True, color="C00000")
WARN_FILL = PatternFill("solid", fgColor="FFE699")
MISSING_FILL = PatternFill("solid", fgColor="F4B084")
P_LOW_FILL = PatternFill("solid", fgColor="C6E0B4")  # p<0.05
P_MID_FILL = PatternFill("solid", fgColor="FFE699")  # 0.05<=p<0.1
P_HIGH_FILL = PatternFill("solid", fgColor="F4B084")  # p>=0.1
PRIMARY_FILL = PatternFill("solid", fgColor="E2EFDA")  # 主比準事例の薄緑（モジュール共通）
PRIMARY_FONT = Font(name="游ゴシック", size=10, bold=True)
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _format_jpy(v):
    if v is None:
        return ""
    return f"{int(round(v)):,}円"


def _format_pct(v):
    if v is None:
        return ""
    return f"{v*100:+.2f}%"


def _format_hijun_corr(multiplier, applies=True, mode="auto"):
    """比準表用の補正値表記（顧客用シート、1セル文字列）。
    mode:
      "top"    = 常に分子側（時点修正：査定時点 / 事例時点）
      "bottom" = 常に分母側（標準化補正・地域格差：100 / 事例評点）
      "auto"   = 補正方向で自動切替
    """
    if not applies:
        return "100/-"
    if abs(multiplier - 1.0) < 0.0005:
        return "100/100"
    if mode == "top":
        return f"{multiplier*100:.1f}/100"
    if mode == "bottom":
        return f"100/{100/multiplier:.1f}"
    if multiplier > 1.0:
        return f"{multiplier*100:.1f}/100"
    return f"100/{100/multiplier:.1f}"


def _hijun_top_bottom(multiplier, applies=True, mode="auto"):
    """比準表用の分子/分母を別々に返す（業者用シートの2行式表示）。
    mode は _format_hijun_corr と同じ。
    Returns: (top, bottom) — 数値または "―"
    """
    if not applies:
        return (100, "―")
    if abs(multiplier - 1.0) < 0.0005:
        return (100, 100)
    if mode == "top":
        return (round(100 * multiplier, 1), 100)
    if mode == "bottom":
        return (100, round(100 / multiplier, 1))
    if multiplier > 1.0:
        return (round(100 * multiplier, 1), 100)
    return (100, round(100 / multiplier, 1))


def _round_3sig(n):
    """上位3桁に四捨五入。例: 424,674,476 → 425,000,000"""
    if n is None or n == 0:
        return n
    import math
    sign = -1 if n < 0 else 1
    n = abs(n)
    digits = int(math.log10(n)) + 1
    if digits <= 3:
        return sign * int(round(n))
    factor = 10 ** (digits - 3)
    return sign * int(round(n / factor) * factor)


def _format_price_full(total_price, area):
    """査定価格表記: 「総額（㎡単価／坪単価）」を上位3桁四捨五入で。
    坪単価 = ㎡単価 ÷ 0.3025
    """
    if total_price is None or area is None or area <= 0:
        return ""
    total_r = _round_3sig(total_price)
    unit_per_sqm = total_r / area
    unit_per_sqm_r = _round_3sig(unit_per_sqm)
    unit_per_tsubo_r = _round_3sig(unit_per_sqm / 0.3025)
    return f"{total_r:,}円（㎡単価 {unit_per_sqm_r:,}円／坪単価 {unit_per_tsubo_r:,}円）"


def _set(ws, row, col, value, font=None, fill=None, align=None, border=False, number_format=None):
    c = ws.cell(row=row, column=col, value=value)
    if font: c.font = font
    if fill: c.fill = fill
    if align: c.alignment = align
    if border: c.border = BORDER
    if number_format: c.number_format = number_format
    return c


def _section_header(ws, row, text, end_col=8):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_col)
    _set(ws, row, 1, text, font=SECTION_FONT, fill=SECTION_FILL,
         align=Alignment(horizontal="left", vertical="center"))


def _adjust_col_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ===== 業者用シート =====
def _write_gyosha_sheet(wb: Workbook, ctx: dict):
    ws = wb.create_sheet("業者用")
    target = ctx["target"]
    asof = ctx["asof"]
    scope_log = ctx["scope_log"]
    rate_info = ctx["rate_info"]
    hed = ctx["hedonic"]
    cases = ctx["cases"]
    breakdown = ctx["breakdown"]
    assess = ctx["assess"]
    refs = ctx["refs"]
    standard_check = ctx["standard_check"]
    hijun_rows = ctx.get("hijun_rows", [])

    r = 1
    # タイトル
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    _set(ws, r, 1, f"土地価格査定 業者用シート — {target.get('物件略号', '')} ({target['市区町村名']} {target.get('地区名', '')})",
         font=TITLE_FONT, fill=TITLE_FILL,
         align=Alignment(horizontal="left", vertical="center"))
    ws.row_dimensions[r].height = 28
    r += 2

    # ヘッダ：物件概要
    _section_header(ws, r, "■ 物件概要・スコープ")
    r += 1
    info = [
        ("査定時点", asof.isoformat()),
        ("所在", f"{target['都道府県名']} {target['市区町村名']} {target.get('地区名', '')}"),
        ("面積", f"{target['面積(㎡)']} ㎡"),
        ("最寄駅", f"{target.get('最寄駅:名称', '')} 徒歩{target.get('最寄駅:距離(分)', '')}分"),
        ("形状", target.get("土地の形状", "")),
        ("接道", f"{target.get('前面道路:種類', '')} 幅員{target.get('前面道路:幅員(m)', '')}m {target.get('前面道路:方位', '')}向"),
        ("使用事例件数", f"{scope_log['final_count']} 件 (IQR除外: {scope_log['iqr_removed']}件、市区町村単位・隣接拡張なし)"),
    ]
    for label, value in info:
        _set(ws, r, 1, label, font=LABEL_FONT, border=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        _set(ws, r, 2, value, font=VALUE_FONT, border=True)
        r += 1
    r += 1

    # 査定価格
    _section_header(ws, r, "■ 査定価格")
    r += 1
    target_area = target["面積(㎡)"]
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    _set(ws, r, 1,
         _format_price_full(assess["central_total_price"], target_area),
         font=BIG_VALUE_FONT,
         align=Alignment(horizontal="left", vertical="center"))
    ws.row_dimensions[r].height = 36
    r += 1

    # 信頼度ラベル（高/中/中-低/低）：n × 自由度調整済 R² × 期待符号整合性ベース
    # 中-低 は「構造問題」（符号反転2件以上 or adj_R² が極端に低い）に限定。
    # 単に件数が少なめなだけのケースは中扱い。
    if hed["ok"]:
        n = hed["n"]
        adj_r2 = hed["adj_r2"]
        EXPECTED_NEG = ("ln_area", "walk_min", "D_shidou", "D_fukuro", "D_fuseikei")
        coef = hed["coefficients"]
        sign_inconsistent = sum(
            1 for name in EXPECTED_NEG if name in coef and coef[name]["beta"] > 0
        )
        sign_checked = sum(1 for name in EXPECTED_NEG if name in coef)
        if sign_inconsistent >= 2 or adj_r2 < 0.3:
            reasons = []
            if sign_inconsistent >= 2:
                reasons.append(f"符号反転 {sign_inconsistent}/{sign_checked} 件")
            if adj_r2 < 0.3:
                reasons.append(f"adj R² = {adj_r2:.2f}（低水準）")
            conf_label = (f"信頼度：中-低（n = {n}, "
                          + ", ".join(reasons)
                          + " — 構造問題の可能性、要再確認）")
            conf_fill = P_HIGH_FILL
        elif n >= 20 and adj_r2 >= 0.5 and sign_inconsistent == 0:
            conf_label = (f"信頼度：高（n = {n}, 自由度調整済 R² = {adj_r2:.2f}, "
                          f"期待符号と全整合）")
            conf_fill = P_LOW_FILL
        else:
            reasons = []
            if n < 20:
                reasons.append(f"事例件数 n = {n} と少なめ")
            if adj_r2 < 0.5:
                reasons.append(f"adj R² = {adj_r2:.2f}（中程度）")
            if sign_inconsistent == 1:
                reasons.append(f"符号反転 1/{sign_checked} 件")
            if not reasons:
                reasons.append(f"n = {n}, adj R² = {adj_r2:.2f}")
            conf_label = "信頼度：中（" + " / ".join(reasons) + "）"
            conf_fill = P_MID_FILL
    else:
        conf_label = "信頼度：低（件数不足のため係数推定不能。顧客用シートは『参考情報』として出力）"
        conf_fill = P_HIGH_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    _set(ws, r, 1, conf_label,
         font=Font(name="游ゴシック", size=11, bold=True),
         fill=conf_fill, border=True,
         align=Alignment(horizontal="left", vertical="center"))
    ws.row_dimensions[r].height = 24
    r += 2

    # 価格レンジ
    _section_header(ws, r, "■ 価格レンジ（類似上位3事例の最大／中央／最小）")
    r += 1
    rng = assess["range"]
    headers = ["区分", "総額", "㎡単価", "坪単価"]
    for j, h in enumerate(headers):
        _set(ws, r, j+1, h, font=LABEL_FONT, fill=PatternFill("solid", fgColor="D9E1F2"), border=True)
    r += 1
    for label, total, unit in [
        ("上限", rng["high_total"], rng["high_unit"]),
        ("中央", rng["central_total"], rng["central_unit"]),
        ("下限", rng["low_total"], rng["low_unit"]),
    ]:
        total_r = _round_3sig(total) if total else None
        unit_sqm_r = _round_3sig(unit) if unit else None
        unit_tsubo_r = _round_3sig(unit / 0.3025) if unit else None
        _set(ws, r, 1, label, font=LABEL_FONT, border=True)
        _set(ws, r, 2, f"{total_r:,}円" if total_r else "", font=VALUE_FONT, border=True)
        _set(ws, r, 3, f"{unit_sqm_r:,}円" if unit_sqm_r else "", font=VALUE_FONT, border=True)
        _set(ws, r, 4, f"{unit_tsubo_r:,}円" if unit_tsubo_r else "", font=VALUE_FONT, border=True)
        r += 1
    r += 1

    # ヘドニック回帰サマリ
    _section_header(ws, r, "■ ヘドニック回帰サマリ（係数全開示）")
    r += 1
    if hed["ok"]:
        _set(ws, r, 1, f"サンプル数 n = {hed['n']}", font=VALUE_FONT)
        _set(ws, r, 3, f"R² = {hed['r2']:.3f}", font=VALUE_FONT)
        _set(ws, r, 5, f"自由度調整済 R² = {hed['adj_r2']:.3f}", font=VALUE_FONT)
        r += 1
        for j, h in enumerate(["特徴量", "推定値 β", "標準誤差", "p値", "有意性"]):
            _set(ws, r, j+1, h, font=LABEL_FONT, fill=PatternFill("solid", fgColor="D9E1F2"), border=True)
        r += 1
        for name, c in hed["coefficients"].items():
            p = c["p"]
            if p < 0.05: fill = P_LOW_FILL; sig = "** (p<0.05)"
            elif p < 0.10: fill = P_MID_FILL; sig = "*  (p<0.10)"
            else: fill = P_HIGH_FILL; sig = "ns"
            _set(ws, r, 1, c["label"], font=VALUE_FONT, border=True)
            _set(ws, r, 2, f"{c['beta']:+.4f}", font=VALUE_FONT, border=True)
            _set(ws, r, 3, f"{c['se']:.4f}", font=VALUE_FONT, border=True)
            _set(ws, r, 4, f"{p:.4f}", font=VALUE_FONT, border=True, fill=fill)
            _set(ws, r, 5, sig, font=VALUE_FONT, border=True, fill=fill)
            r += 1
        r += 1

        # β符号チェック（期待符号 vs 実際の符号）
        EXPECTED_SIGNS = {
            "ln_area": ("負", "面積大→単価下落"),
            "walk_min": ("負", "駅遠→単価下落"),
            "D_shidou": ("負", "私道→減価"),
            "D_fukuro": ("負", "袋地→減価"),
            "D_fuseikei": ("負", "不整形→減価"),
        }
        coef = hed["coefficients"]
        sign_check_label = Font(name="游ゴシック", size=10, bold=True, color="595959")
        _set(ws, r, 1, "▼ β符号チェック（期待符号 vs 実際）", font=sign_check_label)
        r += 1
        for j, h in enumerate(["特徴量", "期待符号", "実際 β", "整合", "経済的解釈"]):
            _set(ws, r, j+1, h, font=LABEL_FONT,
                 fill=PatternFill("solid", fgColor="D9E1F2"), border=True)
        r += 1
        any_inconsistent = False
        for name, (expected, interpretation) in EXPECTED_SIGNS.items():
            if name not in coef:
                continue
            beta = coef[name]["beta"]
            is_neg_expected = (expected == "負")
            is_consistent = (is_neg_expected and beta < 0) or (not is_neg_expected and beta > 0)
            if not is_consistent:
                any_inconsistent = True
            mark = "○ 整合" if is_consistent else "× 反転（要確認）"
            ok_fill = P_LOW_FILL if is_consistent else P_HIGH_FILL
            _set(ws, r, 1, coef[name]["label"], font=VALUE_FONT, border=True)
            _set(ws, r, 2, expected, font=VALUE_FONT, border=True)
            _set(ws, r, 3, f"{beta:+.4f}", font=VALUE_FONT, border=True)
            _set(ws, r, 4, mark, font=VALUE_FONT, border=True, fill=ok_fill)
            _set(ws, r, 5, interpretation, font=VALUE_FONT, border=True)
            r += 1
        if any_inconsistent:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
            _set(ws, r, 1,
                 "※ 符号反転がある場合は外れ値・特徴量不足・地区特性などの構造問題の可能性。事例を再確認してください。",
                 font=VALUE_FONT, fill=WARN_FILL)
            r += 1
    else:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        _set(ws, r, 1, f"※ {hed['skip_reason']}（類似度ベース集約に降格）",
             font=VALUE_FONT, fill=WARN_FILL)
        r += 1
    r += 1

    # （旧「使用事例テーブル」セクションは比準表と重複するため削除）
    # 比準表（後段）に事例番号・取引価格・補正値・試算値が集約されている。

    # 比準表（サンプル準拠の2行式、×列を削除して9列構成）
    if hijun_rows:
        _section_header(ws, r, "■ 比準表（標準画地の比準価格）")
        r += 1
        # 列構成（9列）：
        # 1=事例番号, 2=取引価格, 3=事情補正, 4=時点修正, 5=建付減価,
        # 6=標準化補正, 7=地域格差, 8=試算値, 9=比準値
        header_fill = PatternFill("solid", fgColor="D9E1F2")
        for j, h in enumerate(["事例番号", "取引価格(円/㎡)", "事情補正", "時点修正",
                               "建付減価", "標準化補正", "地域格差",
                               "試算値(円/㎡)", "比準値(円/㎡)"]):
            _set(ws, r, j+1, h, font=LABEL_FONT, fill=header_fill, border=True,
                 align=Alignment(horizontal="center", vertical="center", wrap_text=True))
        ws.row_dimensions[r].height = 32
        r += 1
        # 試算値の中央値を比準値に
        n_rows = len(hijun_rows)
        shisan_list = sorted(h["試算値"] for h in hijun_rows)
        if n_rows % 2 == 1:
            hijun_central = shisan_list[n_rows // 2]
        else:
            hijun_central = (shisan_list[n_rows // 2 - 1] + shisan_list[n_rows // 2]) / 2

        block_start_row = r
        center_align = Alignment(horizontal="center", vertical="center")
        for idx, h in enumerate(hijun_rows):
            is_primary = (h.get("順位") == "主比準")
            fill = PRIMARY_FILL if is_primary else None
            font_top = PRIMARY_FONT if is_primary else VALUE_FONT
            label_font = PRIMARY_FONT if is_primary else LABEL_FONT
            top_row = r
            bot_row = r + 1
            # 補正項目の分子/分母（鑑定書様式）
            # 時点修正：分子側（査定時点 / 事例時点）
            # 標準化補正・地域格差：分母側（100 / 事例評点 = 事例側を分母に置く慣習）
            jijo_top, jijo_bot = _hijun_top_bottom(h["事情補正"], h.get("事情補正_適用", False))
            time_top, time_bot = _hijun_top_bottom(h["時点修正"], mode="top")
            kent_top, kent_bot = _hijun_top_bottom(h["建付減価"], h.get("建付減価_適用", False))
            hyo_top, hyo_bot = _hijun_top_bottom(h["標準化補正"], mode="bottom")
            chi_top, chi_bot = _hijun_top_bottom(h["地域格差"], mode="bottom")
            # 上行（分子）
            _set(ws, top_row, 3, jijo_top, font=font_top, fill=fill, border=True, align=center_align)
            _set(ws, top_row, 4, time_top, font=font_top, fill=fill, border=True, align=center_align)
            _set(ws, top_row, 5, kent_top, font=font_top, fill=fill, border=True, align=center_align)
            _set(ws, top_row, 6, hyo_top, font=font_top, fill=fill, border=True, align=center_align)
            _set(ws, top_row, 7, chi_top, font=font_top, fill=fill, border=True, align=center_align)
            # 下行（分母）
            _set(ws, bot_row, 3, jijo_bot, font=font_top, fill=fill, border=True, align=center_align)
            _set(ws, bot_row, 4, time_bot, font=font_top, fill=fill, border=True, align=center_align)
            _set(ws, bot_row, 5, kent_bot, font=font_top, fill=fill, border=True, align=center_align)
            _set(ws, bot_row, 6, hyo_bot, font=font_top, fill=fill, border=True, align=center_align)
            _set(ws, bot_row, 7, chi_bot, font=font_top, fill=fill, border=True, align=center_align)
            # 2行マージ：事例番号(1), 取引価格(2), 試算値(8)
            for col in [1, 2, 8]:
                ws.merge_cells(start_row=top_row, start_column=col,
                               end_row=bot_row, end_column=col)
            # 事例番号 = MLITデータ番号（透明性のため、人為的ラベルではない）
            case_no_str = str(h.get("事例番号", "?"))
            _set(ws, top_row, 1, case_no_str, font=label_font, fill=fill, border=True,
                 align=center_align)
            _set(ws, top_row, 2, f"{int(h['取引価格']):,}",
                 font=font_top, fill=fill, border=True, align=center_align)
            _set(ws, top_row, 8, f"{int(round(h['試算値'])):,}",
                 font=font_top, fill=fill, border=True, align=center_align)
            r += 2
        block_end_row = r - 1
        # 比準値列を全事例マージ
        ws.merge_cells(start_row=block_start_row, start_column=9,
                       end_row=block_end_row, end_column=9)
        _set(ws, block_start_row, 9, f"{int(round(hijun_central)):,}",
             font=Font(name="游ゴシック", size=12, bold=True, color="C00000"),
             fill=PRIMARY_FILL, border=True,
             align=Alignment(horizontal="center", vertical="center"))
        # 注釈
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
        _set(ws, r, 1,
             "※ 事例番号 = MLITデータ原本の行番号。比準値 = 3事例の試算値の中央値。"
             "各補正は「分子/分母」形式（上段=分子、下段=分母）。「100/-」は補正非該当。"
             "標準化補正＝画地・形状（ln_shape, 袋地, 不整形）、"
             "地域格差＝地域・街路・交通（面積, 駅徒歩, 道路幅員, 南向き, 私道）のヘドニック係数積。"
             "緑色行＝類似度top1の事例。",
             font=Font(name="游ゴシック", size=9, italic=True, color="595959"),
             align=Alignment(wrap_text=True, vertical="top"))
        ws.row_dimensions[r].height = 50
        r += 2

    # 個別格差補正の内訳
    if breakdown:
        _section_header(ws, r, "■ 個別格差補正の内訳（事例別 × 補正項目別、%）")
        r += 1
        # キー順序：事例番号 を先頭に固定し、それ以外をその後に
        all_keys = list(breakdown[0].keys())
        ordered_keys = ["事例番号"] + [k for k in all_keys if k != "事例番号"]
        meta_keys = {"事例番号", "district", "area"}
        header_labels = {"事例番号": "事例番号", "district": "地区", "area": "面積㎡"}
        for j, h in enumerate(ordered_keys):
            label = header_labels.get(h, h)
            _set(ws, r, j+1, label, font=LABEL_FONT,
                 fill=PatternFill("solid", fgColor="D9E1F2"), border=True)
        r += 1
        for idx, entry in enumerate(breakdown):
            is_primary = (idx == 0)
            fill = PRIMARY_FILL if is_primary else None
            font = PRIMARY_FONT if is_primary else VALUE_FONT
            for j, k in enumerate(ordered_keys):
                v = entry.get(k)
                if k in meta_keys:
                    if isinstance(v, float):
                        v = int(v) if k == "area" else v
                    _set(ws, r, j+1, v,
                         font=PRIMARY_FONT if (is_primary and k == "事例番号") else font,
                         border=True, fill=fill)
                else:
                    _set(ws, r, j+1,
                         f"{v:+.2f}%" if isinstance(v, (int, float)) else v,
                         font=font, border=True, fill=fill)
            r += 1
        r += 1

    # 2価格サマリ：採用査定価格 vs ヘドニック母集団予測
    target_area_local = target["面積(㎡)"]
    central_unit = assess.get("central_unit_price")
    hed_pred = refs.get("hedonic_pred")

    _section_header(ws, r, "■ 2価格サマリ（採用査定価格 vs ヘドニック母集団予測の乖離）")
    r += 1
    for j, h in enumerate(["区分", "㎡単価", "総額", "採用査定との乖離率"]):
        _set(ws, r, j+1, h, font=LABEL_FONT,
             fill=PatternFill("solid", fgColor="D9E1F2"), border=True)
    r += 1

    # ① 土地比準（採用）
    _set(ws, r, 1, "① 土地比準（採用）",
         font=Font(name="游ゴシック", size=10, bold=True),
         fill=PRIMARY_FILL, border=True)
    if central_unit:
        _set(ws, r, 2, f"{int(central_unit):,}円",
             font=Font(name="游ゴシック", size=10, bold=True), fill=PRIMARY_FILL, border=True)
        _set(ws, r, 3, _format_jpy(central_unit * target_area_local),
             font=Font(name="游ゴシック", size=10, bold=True), fill=PRIMARY_FILL, border=True)
    _set(ws, r, 4, "—", font=VALUE_FONT, fill=PRIMARY_FILL, border=True)
    r += 1

    # ② ヘドニック母集団予測
    if hed_pred and central_unit:
        dev = (hed_pred - central_unit) / central_unit * 100
        abs_dev = abs(dev)
        if abs_dev <= 15:
            dev_fill = P_LOW_FILL
            dev_guide = "（15%以内：採用査定とヘドニック予測が概ね整合）"
        elif abs_dev <= 30:
            dev_fill = P_MID_FILL
            dev_guide = "（15〜30%：地域特性または個別事例の特殊性を確認すると良い）"
        else:
            dev_fill = P_HIGH_FILL
            dev_guide = ("※ 30%超：主比準事例が母集団から外れている可能性。"
                         "事例選定と特徴量を再確認してください。")
        _set(ws, r, 1, "② ヘドニック母集団予測", font=VALUE_FONT, border=True)
        _set(ws, r, 2, f"{int(hed_pred):,}円", font=VALUE_FONT, border=True)
        _set(ws, r, 3, _format_jpy(hed_pred * target_area_local), font=VALUE_FONT, border=True)
        _set(ws, r, 4, f"{dev:+.1f}%", font=VALUE_FONT, fill=dev_fill, border=True)
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        _set(ws, r, 1, dev_guide,
             font=Font(name="游ゴシック", size=9, italic=True, color="595959"))
        r += 1
    elif central_unit:
        _set(ws, r, 1, "② ヘドニック母集団予測", font=VALUE_FONT, border=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        _set(ws, r, 2, "（件数不足のため算出不能）", font=VALUE_FONT, fill=MISSING_FILL, border=True)
        r += 1
    r += 1

    # 地域標準価格チェック（地区一致優先＋時点補間）
    _section_header(ws, r, "■ 地域標準価格チェック（公示・基準地価との比較）")
    r += 1
    if standard_check.get("standard_price_per_sqm"):
        # 選定方法ラベル
        method = standard_check.get("selection_method", "city_average")
        method_label = "地区一致" if method == "district_match" else "市区町村平均"
        method_fill = P_LOW_FILL if method == "district_match" else P_MID_FILL
        _set(ws, r, 1, "標準地選定方法", font=VALUE_FONT, border=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        _set(ws, r, 2,
             f"{method_label}（n = {standard_check.get('n_points', 0)} 地点）",
             font=VALUE_FONT, fill=method_fill, border=True)
        _set(ws, r, 5, f"asof = {asof.isoformat()} へ時点補間済み",
             font=Font(name="游ゴシック", size=9, italic=True, color="595959"))
        r += 1
        _set(ws, r, 1, f"{standard_check['source']} 標準価格（補間後平均）",
             font=LABEL_FONT, border=True)
        _set(ws, r, 3, f"{int(standard_check['standard_price_per_sqm']):,} 円/㎡",
             font=Font(name="游ゴシック", size=10, bold=True), border=True)
        r += 1
        ratio = assess["central_unit_price"] / standard_check["standard_price_per_sqm"] if standard_check["standard_price_per_sqm"] else None
        if ratio:
            ratio_fill = (P_LOW_FILL if 0.85 <= ratio <= 1.15 else
                          P_MID_FILL if 0.7 <= ratio <= 1.5 else P_HIGH_FILL)
            _set(ws, r, 1, "査定単価／標準価格", font=LABEL_FONT, border=True)
            _set(ws, r, 3, f"{ratio:.2f} 倍",
                 font=Font(name="游ゴシック", size=10, bold=True),
                 fill=ratio_fill, border=True)
            r += 1
        # 使用した標準地の番号と所在（最大3件）
        points = standard_check.get("selected_points", [])
        if points:
            _set(ws, r, 1, "使用した標準地",
                 font=LABEL_FONT, fill=PatternFill("solid", fgColor="D9E1F2"), border=True)
            for j, h in enumerate(["番号", "所在（地区）", "用途", "asof補間価格"]):
                _set(ws, r, j + 2, h, font=LABEL_FONT,
                     fill=PatternFill("solid", fgColor="D9E1F2"), border=True)
            r += 1
            for pt in points[:5]:
                _set(ws, r, 1, "", border=True)
                _set(ws, r, 2, str(pt.get("id", "")), font=VALUE_FONT, border=True)
                _set(ws, r, 3, str(pt.get("district", "")), font=VALUE_FONT, border=True)
                _set(ws, r, 4, str(pt.get("use", "")), font=VALUE_FONT, border=True)
                _set(ws, r, 5, f"{int(pt.get('price_at_asof', 0)):,} 円/㎡",
                     font=VALUE_FONT, border=True)
                r += 1
            if len(points) > 5:
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
                _set(ws, r, 1,
                     f"（他 {len(points) - 5} 地点も平均算出に含めています）",
                     font=Font(name="游ゴシック", size=9, italic=True, color="595959"))
                r += 1
    else:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        _set(ws, r, 1, "標準価格情報なし（公示・基準地価データ欠損）",
             font=VALUE_FONT, fill=MISSING_FILL)
        r += 1
    if rate_info.get("rate") is not None:
        method = rate_info.get("method", "")
        method_label = (
            "地区一致" if method == "district_match"
            else "市区町村平均" if method == "city_average"
            else "隣接拡張" if method == "neighbor"
            else method
        )
        method_fill = P_LOW_FILL if method == "district_match" else P_MID_FILL
        _set(ws, r, 1, "時点修正年率", font=LABEL_FONT, border=True)
        _set(ws, r, 3, f"{rate_info['rate']*100:+.2f}% / 年",
             font=Font(name="游ゴシック", size=10, bold=True),
             fill=method_fill, border=True)
        _set(ws, r, 5,
             f"（{method_label}, n = {rate_info['n_points']} 地点 / 出典 {rate_info['source']}）",
             font=Font(name="游ゴシック", size=9, italic=True, color="595959"))
        r += 1

        # 変動率の期間
        rate_pts = rate_info.get("selected_points", [])
        if rate_pts:
            d_prev = rate_pts[0].get("date_prev")
            d_curr = rate_pts[0].get("date_curr")
            if d_prev and d_curr:
                _set(ws, r, 1, "変動率の期間", font=VALUE_FONT, border=True)
                _set(ws, r, 3,
                     f"{d_prev} → {d_curr}（直近1年の変動率を採用）",
                     font=VALUE_FONT, border=True)
                r += 1

        # 年率算出に使用した標準地（最大5件）
        if rate_pts:
            _set(ws, r, 1, "年率算出に使用した標準地",
                 font=LABEL_FONT, fill=PatternFill("solid", fgColor="D9E1F2"), border=True)
            for j, h in enumerate(["番号", "地区", "出典", "前年", "当年", "変動率"]):
                _set(ws, r, j + 2, h, font=LABEL_FONT,
                     fill=PatternFill("solid", fgColor="D9E1F2"), border=True)
            r += 1
            for pt in rate_pts[:5]:
                _set(ws, r, 1, "", border=True)
                _set(ws, r, 2, str(pt.get("id", "")), font=VALUE_FONT, border=True)
                _set(ws, r, 3, str(pt.get("district", "")), font=VALUE_FONT, border=True)
                _set(ws, r, 4, str(pt.get("source", "")), font=VALUE_FONT, border=True)
                _set(ws, r, 5, f"{int(pt.get('p_prev', 0)):,}",
                     font=VALUE_FONT, border=True)
                _set(ws, r, 6, f"{int(pt.get('p_curr', 0)):,}",
                     font=VALUE_FONT, border=True)
                _set(ws, r, 7, f"{pt.get('rate', 0)*100:+.2f}%",
                     font=VALUE_FONT, border=True)
                r += 1
            if len(rate_pts) > 5:
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
                _set(ws, r, 1,
                     f"（他 {len(rate_pts) - 5} 地点も平均算出に含めています）",
                     font=Font(name="游ゴシック", size=9, italic=True, color="595959"))
                r += 1
    r += 1

    # 警告欄
    _section_header(ws, r, "■ 警告・注記")
    r += 1
    warnings = list(scope_log.get("warnings", []))
    if not hed["ok"]:
        warnings.append(hed["skip_reason"])
    if assess.get("warning"):
        warnings.append(assess["warning"])
    if not warnings:
        warnings = ["特になし"]
    for w in warnings:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        _set(ws, r, 1, f"・{w}", font=VALUE_FONT, fill=WARN_FILL if w != "特になし" else None)
        r += 1

    _adjust_col_widths(ws, [22, 14, 18, 14, 18, 18, 18, 14])


# ===== 顧客用シート =====
def _qualitative(beta_label, target_v, mean_v):
    """個別要因の定性表現（係数を開示せず方向と程度のみ）。"""
    if mean_v == 0 or mean_v is None or target_v is None:
        return None
    diff = target_v - mean_v
    if abs(diff) < 0.1 * abs(mean_v):
        return "同水準"
    if diff > 0:
        return "やや広め" if "面積" in beta_label else "やや遠め" if "駅" in beta_label else "ややプラス要因"
    return "やや狭め" if "面積" in beta_label else "やや近め" if "駅" in beta_label else "ややマイナス要因"


def _write_kokyaku_sheet(wb: Workbook, ctx: dict):
    ws = wb.create_sheet("顧客用")
    target = ctx["target"]
    asof = ctx["asof"]
    rate_info = ctx["rate_info"]
    cases = ctx["cases"]
    assess = ctx["assess"]
    standard_check = ctx["standard_check"]
    is_degraded = not ctx["hedonic"]["ok"]
    hijun_rows = ctx.get("hijun_rows", [])

    r = 1
    # タイトル（降格時はラベル変更）
    title_text = (
        f"土地価格 参考情報 — {target['市区町村名']} {target.get('地区名', '')}"
        if is_degraded
        else f"土地査定報告書 — {target['市区町村名']} {target.get('地区名', '')}"
    )
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    assert_clean(title_text, "title")
    _set(ws, r, 1, title_text, font=TITLE_FONT, fill=TITLE_FILL,
         align=Alignment(horizontal="left", vertical="center"))
    ws.row_dimensions[r].height = 28
    r += 1

    # 降格時の参考情報バナー（赤色）
    if is_degraded:
        warn_red_fill = PatternFill("solid", fgColor="C00000")
        warn_red_font = Font(name="游ゴシック", size=11, bold=True, color="FFFFFF")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        banner_text = (
            "※ 取引事例件数が不足しているため、本資料は「参考情報」としてご覧ください（査定書ではありません）。"
            " 正式な査定価格は、ご担当者による現地確認・追加調査を経て決定する必要があります。"
        )
        assert_clean(banner_text, "degraded banner")
        _set(ws, r, 1, banner_text, font=warn_red_font, fill=warn_red_fill,
             align=Alignment(wrap_text=True, vertical="center"))
        ws.row_dimensions[r].height = 38
        r += 2
    else:
        r += 1

    # ■ 査定結果サマリ
    _section_header(ws, r, "■ 査定結果サマリ", end_col=6)
    r += 1
    rng = assess["range"]
    target_area = target["面積(㎡)"]

    # 査定価格（総額＋㎡単価＋坪単価併記、降格時はラベル変更）
    price_label = "参考価格" if is_degraded else "査定価格"
    summary_rows = [
        (price_label, _format_price_full(rng["central_total"], target_area)),
        ("面積", f"{target_area} ㎡"),
        ("査定時点", asof.isoformat()),
    ]
    for label, value in summary_rows:
        assert_clean(label, "summary label")
        assert_clean(value, "summary value")
        _set(ws, r, 1, label, font=LABEL_FONT, border=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        _set(ws, r, 2, value, font=VALUE_FONT, border=True)
        r += 1
    r += 1

    # 価格レンジ（上限／中央／下限）
    _section_header(ws, r, "■ 価格レンジ", end_col=6)
    r += 1
    for j, h in enumerate(["区分", "総額", "㎡単価", "坪単価"]):
        _set(ws, r, j+1, h, font=LABEL_FONT,
             fill=PatternFill("solid", fgColor="D9E1F2"), border=True)
    r += 1
    for label, total, unit in [
        ("上限", rng["high_total"], rng["high_unit"]),
        ("中央", rng["central_total"], rng["central_unit"]),
        ("下限", rng["low_total"], rng["low_unit"]),
    ]:
        total_r = _round_3sig(total) if total else None
        unit_sqm_r = _round_3sig(unit) if unit else None
        unit_tsubo_r = _round_3sig(unit / 0.3025) if unit else None
        _set(ws, r, 1, label, font=LABEL_FONT, border=True)
        _set(ws, r, 2, f"{total_r:,}円" if total_r else "", font=VALUE_FONT, border=True)
        _set(ws, r, 3, f"{unit_sqm_r:,}円" if unit_sqm_r else "", font=VALUE_FONT, border=True)
        _set(ws, r, 4, f"{unit_tsubo_r:,}円" if unit_tsubo_r else "", font=VALUE_FONT, border=True)
        r += 1
    r += 1

    # ■ 標準価格
    _section_header(ws, r, "■ 標準価格", end_col=6)
    r += 1
    if standard_check.get("standard_price_per_sqm"):
        text = f"本地区の{standard_check['source']}による標準価格は {int(standard_check['standard_price_per_sqm']):,} 円/㎡ です。"
    else:
        text = "本地区の公的な標準価格データが取得できなかったため、参考情報なしで査定しています。"
    assert_clean(text, "standard")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    _set(ws, r, 1, text, font=VALUE_FONT, align=Alignment(wrap_text=True, vertical="top"))
    ws.row_dimensions[r].height = 30
    r += 2

    # ■ 時点修正
    _section_header(ws, r, "■ 時点修正", end_col=6)
    r += 1
    if rate_info.get("rate") is not None:
        rate_pct = rate_info["rate"] * 100
        direction = "上昇" if rate_pct > 0 else "下落" if rate_pct < 0 else "横ばい"
        text = (f"直近の本地区の地価は年率 {rate_pct:+.2f}% で{direction}しています。"
                f"この動きを踏まえて、過去の取引事例を査定時点に補正しています。")
    else:
        text = "時点修正の根拠となる地価データが取得できなかったため、補正なしで査定しています。"
    assert_clean(text, "time_adjust")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    _set(ws, r, 1, text, font=VALUE_FONT, align=Alignment(wrap_text=True, vertical="top"))
    ws.row_dimensions[r].height = 36
    r += 2

    # ■ 個別要因補正（主比準事例との比較ベース）
    _section_header(ws, r, "■ 個別要因補正", end_col=6)
    r += 1
    qual_lines = []
    if len(cases) > 0:
        primary = cases.iloc[0]
        primary_area = float(primary["area"])
        primary_walk = float(primary["walk_min"]) if "walk_min" in cases.columns and pd.notna(primary["walk_min"]) else None
        primary_shape = str(primary.get("shape", ""))
        primary_road_type = str(primary.get("road_type", ""))
        # 面積
        if target["面積(㎡)"] > primary_area * 1.1:
            qual_lines.append("・面積は主比準事例よりやや広めの水準です。")
        elif target["面積(㎡)"] < primary_area * 0.9:
            qual_lines.append("・面積は主比準事例よりやや狭めの水準です。")
        else:
            qual_lines.append("・面積は主比準事例と同水準です。")
        # 駅距離
        if primary_walk and target.get("最寄駅:距離(分)"):
            tw = target["最寄駅:距離(分)"]
            if tw < primary_walk - 1:
                qual_lines.append("・最寄駅までの距離は主比準事例より近めです。")
            elif tw > primary_walk + 1:
                qual_lines.append("・最寄駅までの距離は主比準事例よりやや遠めです。")
            else:
                qual_lines.append("・最寄駅までの距離は主比準事例と同水準です。")
        # 形状
        shape = target.get("土地の形状", "")
        if shape == primary_shape:
            if shape == "整形":
                qual_lines.append("・形状は整っており、主比準事例と同水準です。")
            else:
                qual_lines.append(f"・形状は主比準事例と同水準（{shape}）です。")
        else:
            if shape == "整形":
                qual_lines.append("・形状は主比準事例より整っているため、その分プラスに調整しています。")
            elif shape == "不整形":
                qual_lines.append("・形状に不整形な部分があり、その分マイナスに調整しています。")
            elif shape == "袋地":
                qual_lines.append("・袋地（接道条件が限定的）のため、相応のマイナス調整を行っています。")
        # 接道
        target_road = target.get("前面道路:種類", "")
        if target_road == primary_road_type:
            qual_lines.append(f"・接道は主比準事例と同条件（{target_road}）です。")
        elif target_road == "私道":
            qual_lines.append("・接道は私道のため、その分マイナスに調整しています。")
        else:
            qual_lines.append("・接道は公道のため、主比準事例（私道）と比べてプラスに調整しています。")
    text = "\n".join(qual_lines) if qual_lines else "（事例情報が不足しているため、定量的な補正は行っていません）"
    assert_clean(text, "qualitative")
    ws.merge_cells(start_row=r, start_column=1, end_row=r+len(qual_lines), end_column=6)
    _set(ws, r, 1, text, font=VALUE_FONT, align=Alignment(wrap_text=True, vertical="top"))
    ws.row_dimensions[r].height = max(20, 18 * len(qual_lines))
    r += len(qual_lines) + 2

    # ■ 査定価格 / 参考価格
    section_label = "■ 参考価格" if is_degraded else "■ 査定価格"
    _section_header(ws, r, section_label, end_col=6)
    r += 1
    if is_degraded:
        text = (f"上記を踏まえた本物件の参考価格は {_format_price_full(rng['central_total'], target_area)} です。"
                f"（取引事例件数が不足しているため、正式な査定価格ではありません）")
    else:
        text = f"上記を踏まえた本物件の査定価格は {_format_price_full(rng['central_total'], target_area)} となります。"
    assert_clean(text, "final")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    _set(ws, r, 1, text, font=BIG_VALUE_FONT, align=Alignment(wrap_text=True, vertical="center"))
    ws.row_dimensions[r].height = 50
    r += 2

    # ■ 比準表（主比準1事例の縦転置形式）— 業者用比準表を縦に
    if hijun_rows:
        primary_h = next((h for h in hijun_rows if h.get("順位") == "主比準"), hijun_rows[0])
        _section_header(ws, r, "■ 比準表（主比準取引事例による試算）", end_col=6)
        r += 1
        # 縦2列構成（項目 / 値）
        header_fill = PatternFill("solid", fgColor="D9E1F2")
        _set(ws, r, 1, "項目", font=LABEL_FONT, fill=header_fill, border=True,
             align=Alignment(horizontal="center"))
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        _set(ws, r, 2, "値", font=LABEL_FONT, fill=header_fill, border=True,
             align=Alignment(horizontal="center"))
        r += 1
        case_summary = (
            f"事例 {primary_h.get('事例番号', '?')}（{primary_h.get('地区', '')} "
            f"{primary_h.get('面積', 0)}㎡, 取引時期 {primary_h.get('取引時点', '')}）"
        )
        assert_clean(case_summary, "case summary")
        items = [
            ("事例番号", case_summary),
            ("取引価格", f"{int(primary_h['取引価格']):,} 円/㎡"),
            ("事情補正", _format_hijun_corr(
                primary_h["事情補正"], primary_h.get("事情補正_適用", False))),
            ("時点修正", _format_hijun_corr(primary_h["時点修正"], mode="top")),
            ("建付減価", _format_hijun_corr(
                primary_h["建付減価"], primary_h.get("建付減価_適用", False))),
            ("形状補正", _format_hijun_corr(primary_h["標準化補正"], mode="bottom")),
            ("地域格差", _format_hijun_corr(primary_h["地域格差"], mode="bottom")),
            ("試算値", f"{int(round(primary_h['試算値'])):,} 円/㎡"),
        ]
        for label, value in items:
            assert_clean(label, "hijun row label")
            assert_clean(str(value), "hijun row value")
            is_shisan = (label == "試算値")
            row_fill = PRIMARY_FILL if is_shisan else None
            value_font = (
                Font(name="游ゴシック", size=12, bold=True, color="C00000")
                if is_shisan
                else Font(name="游ゴシック", size=11, bold=True)
            )
            _set(ws, r, 1, label, font=LABEL_FONT, fill=row_fill, border=True,
                 align=Alignment(horizontal="center", vertical="center"))
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            _set(ws, r, 2, value, font=value_font, fill=row_fill, border=True,
                 align=Alignment(horizontal="left", vertical="center"))
            ws.row_dimensions[r].height = 22
            r += 1
        # 注釈
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        _set(ws, r, 1,
             "※ 補正値の表記：「100/100」は補正なし、「●/100」は事例より本物件が良い項目"
             "（プラス補正）、「100/●」は事例の方が良い項目（マイナス補正）、"
             "「100/-」は該当なし。",
             font=Font(name="游ゴシック", size=9, italic=True, color="595959"),
             align=Alignment(wrap_text=True, vertical="top"))
        ws.row_dimensions[r].height = 30
        r += 2

    # ■ 査定の考え方
    _section_header(ws, r, "■ 査定の考え方", end_col=6)
    r += 1
    sentences = [
        "本査定は、本地区とその周辺で本物件と類似性の高い取引事例を複数選定し、",
        "それぞれの単価を本物件の特徴（面積・最寄駅までの距離・形状・接道など）に合わせて調整したうえで、査定価格を算出しています。",
        "なお、最も類似性の高い1件を「主比準取引事例」として表示しています。",
        "標準価格や地価の動きとの整合も確認しています。",
        "あくまで一次査定であり、現地確認や市場動向によって最終価格は変動し得ます。",
    ]
    text = "".join(sentences)
    assert_clean(text, "story")
    ws.merge_cells(start_row=r, start_column=1, end_row=r+2, end_column=6)
    _set(ws, r, 1, text, font=VALUE_FONT, align=Alignment(wrap_text=True, vertical="top"))
    ws.row_dimensions[r].height = 80
    r += 4

    # ■ 留意点
    _section_header(ws, r, "■ 留意点", end_col=6)
    r += 1
    notes = [
        "本資料は媒介査定における一次査定の参考資料です。",
        "実勢価格との乖離が生じる可能性があります。",
        "最終的な売出価格・成約価格は、ご担当者と相談のうえ判断してください。",
    ]
    for note in notes:
        assert_clean(note, "note")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        _set(ws, r, 1, f"・{note}", font=VALUE_FONT)
        r += 1

    _adjust_col_widths(ws, [18, 16, 18, 16, 16, 16])


def write_xlsx(ctx: dict, output_path: Path) -> Path:
    wb = Workbook()
    # デフォルトシートを削除
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    _write_gyosha_sheet(wb, ctx)
    _write_kokyaku_sheet(wb, ctx)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
