"""Cowork outputs フォルダを監視し、新しい xlsx をデスクトップへ自動コピー。

Claude Cowork のサンドボックスは、ホスト OS のデスクトップに直接ファイルを
書き込めません。tochi-satei-kun が生成した査定書 (土地査定_*.xlsx) を
ユーザーの手元へ届けるため、本スクリプトを Windows 側で常駐させて
outputs フォルダを監視します。

【動作】
  5 秒ごとに Cowork の local-agent-mode-sessions フォルダを再帰的に走査し、
  「土地査定_*.xlsx」が新規生成されていればデスクトップへコピーします。
  起動時点で既に存在するファイルは無視し（過去ファイルを再コピーしない）、
  以降に生成された新規ファイルだけが対象です。

【手動起動】
  python C:\\Users\\<あなた>\\watch_cowork_outputs.py

【Windows 起動時の自動常駐】
  1. Win+R → shell:startup
  2. 開いたフォルダに watch_cowork_outputs.bat を配置
     （.bat の中身は付属のテンプレートを参照）
  3. PC 再起動後、pythonw でバックグラウンド常駐

【停止方法】
  - フォアグラウンド実行中: Ctrl+C
  - バックグラウンド常駐中: タスクマネージャで pythonw.exe を終了

【ログ】
  %USERPROFILE%\\watch_cowork_outputs.log
"""
from pathlib import Path
import shutil
import sys
import time

# ===== 設定 =====
# Cowork (Claude Desktop アプリ) の outputs ルート。
# `Claude_pzs8sxrjxfjjc` は Claude Desktop アプリの Package Family Name で
# Windows 環境共通。ユーザー名部分は Path.home() で自動解決します。
WATCH_ROOT = (
    Path.home()
    / "AppData" / "Local" / "Packages"
    / "Claude_pzs8sxrjxfjjc"
    / "LocalCache" / "Roaming" / "Claude" / "local-agent-mode-sessions"
)
DESKTOP_CANDIDATES = [
    Path.home() / "OneDrive" / "デスクトップ",
    Path.home() / "OneDrive" / "Desktop",
    Path.home() / "Desktop",
]
PATTERN = "土地査定_*.xlsx"   # tochi-satei-kun の出力ファイル名規則
POLL_INTERVAL = 5             # 秒
LOG_FILE = Path.home() / "watch_cowork_outputs.log"


def _resolve_desktop() -> Path:
    """デスクトップパスを順に試行し、最初に見つかったものを返す。"""
    for cand in DESKTOP_CANDIDATES:
        if cand.exists() and cand.is_dir():
            return cand
    raise RuntimeError(
        "デスクトップフォルダが見つかりません。候補:\n  "
        + "\n  ".join(str(c) for c in DESKTOP_CANDIDATES)
    )


def _log(msg: str):
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}"
    try:
        print(line, flush=True)
    except UnicodeEncodeError:
        print(line.encode("ascii", "replace").decode("ascii"), flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass  # ログ書込失敗は常駐継続を優先して無視


def _wait_for_watch_root() -> None:
    """WATCH_ROOT が現れるまでポーリング。

    Cowork を一度も起動していない環境では local-agent-mode-sessions
    フォルダ自体が存在しない。Watcher を先に常駐させても落ちないよう、
    フォルダが現れるまで POLL_INTERVAL 秒ごとに静かに待機する。
    """
    if WATCH_ROOT.exists():
        return
    _log(f"WATCH_ROOT がまだ存在しません: {WATCH_ROOT}")
    _log("Cowork (Claude Desktop アプリ) の初回起動を待機中…")
    while not WATCH_ROOT.exists():
        time.sleep(POLL_INTERVAL)
    _log("WATCH_ROOT を検出。監視を開始します。")


def main():
    try:
        desktop = _resolve_desktop()
    except RuntimeError as e:
        _log(f"[ERROR] {e}")
        sys.exit(1)

    _wait_for_watch_root()

    _log("=" * 60)
    _log("Watcher 起動")
    _log(f"監視先: {WATCH_ROOT}")
    _log(f"配置先: {desktop}")
    _log(f"対象  : {PATTERN}")
    _log(f"間隔  : {POLL_INTERVAL} 秒")

    # 起動時点で既に存在する xlsx は対象外（過去ファイルを再コピーしない）
    seen = set()
    for f in WATCH_ROOT.rglob(PATTERN):
        seen.add(f.resolve())
    _log(f"初期スキャン: 既存 xlsx {len(seen)} 件を seen に登録")

    while True:
        try:
            for src in WATCH_ROOT.rglob(PATTERN):
                key = src.resolve()
                if key in seen:
                    continue
                try:
                    # ファイル書き込み完了まで待つ（小さい遅延）
                    size1 = src.stat().st_size
                    time.sleep(1)
                    size2 = src.stat().st_size
                    if size1 != size2:
                        # まだ書き込み中。次回ポーリングに回す
                        continue
                    dest = desktop / src.name
                    shutil.copy2(src, dest)
                    _log(f"[COPY OK] {src.name} ({size2:,} bytes) -> {dest}")
                    seen.add(key)
                except (OSError, PermissionError) as e:
                    _log(f"[COPY ERR] {src.name}: {e}")
            time.sleep(POLL_INTERVAL)
        except KeyboardInterrupt:
            _log("Ctrl+C 検出、watcher 終了")
            break
        except Exception as e:
            _log(f"[FATAL] {e}; {POLL_INTERVAL} 秒後にリトライ")
            time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()
