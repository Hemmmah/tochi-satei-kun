"""顧客用シートの禁止語チェック。
顧客に出す文言にAI/モデル/統計用語が混入することを機械的に防ぐ。
xlsx_writer.py が顧客用シート書き込み前に呼び出す。
"""

FORBIDDEN_WORDS = [
    "AI", "A.I.",
    "モデル",
    "機械学習",
    "ヘドニック",
    "β", "ベータ",
    "回帰", "OLS",
    "統計",
    "予測", "推定",
    "アルゴリズム",
    "学習",
    "係数",
    "R²", "R2", "R^2",
    "p値", "p-value",
]


def check_text(text: str):
    if not isinstance(text, str):
        return True, []
    detected = [w for w in FORBIDDEN_WORDS if w in text]
    return (len(detected) == 0, detected)


def assert_clean(text: str, context: str = ""):
    ok, detected = check_text(text)
    if not ok:
        raise ValueError(
            f"顧客用シートに禁止語が混入: {detected} / context={context} / text={text!r}"
        )
