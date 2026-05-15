# tools/

Claude Cowork のサンドボックスはホスト OS のデスクトップに直接書き込めないため、本ディレクトリには **プラグイン外部で動かすホスト側ユーティリティ** をまとめています。エンドユーザー向けの導入手順は [../INSTALL.md](../INSTALL.md) を参照してください。

## ファイル一覧

| ファイル | 役割 |
|---|---|
| `watch_cowork_outputs.py` | Cowork outputs フォルダを監視し、`土地査定_*.xlsx` が生成されたらデスクトップへ自動コピーする Windows 常駐スクリプト |
| `watch_cowork_outputs.bat` | Windows スタートアップ登録用テンプレート。`pythonw` でバックグラウンド起動 |

## 開発者向けノート

- 監視対象パターン `土地査定_*.xlsx` は `skills/tochi-satei-kun/scripts/xlsx_writer.py` の出力ファイル名規則と一致させています。出力命名規則を変更する際は Watcher 側の `PATTERN` も同期してください。
- `WATCH_ROOT` 内の `Claude_pzs8sxrjxfjjc` は Claude Desktop アプリの Package Family Name で、Windows 環境では全ユーザー共通です。アプリのリブランディングや配布形態変更があれば差し替えが必要です。
- 現状は Windows 専用です。macOS 版（LaunchAgent 実装）は Phase 4 で検討予定。
