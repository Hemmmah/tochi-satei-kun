@echo off
REM Cowork outputs Watcher 起動 (tochi-satei-kun)
REM
REM 使い方:
REM   1. watch_cowork_outputs.py を %USERPROFILE% (例: C:\Users\<あなた>) にコピー
REM   2. Win+R → "shell:startup" で開いたフォルダに本 .bat をコピー
REM   3. ダブルクリック or PC 再起動で常駐開始
REM
REM 別フォルダに置きたい場合は下行の "%USERPROFILE%\watch_cowork_outputs.py" を
REM 実際のフルパスに書き換えてください。

start "" pythonw "%USERPROFILE%\watch_cowork_outputs.py"
