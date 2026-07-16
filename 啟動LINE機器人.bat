@echo off
chcp 65001 >nul
REM ============================================================
REM  鮮奶物流 LINE 機器人 一鍵啟動
REM  雙擊本檔：自動起隧道 + line_bot，並抓新網址寫入 .env
REM  關閉本視窗 = 全部停止（LINE 就連不上）
REM ============================================================
set HERE=%~dp0
cd /d "%HERE%"

echo 正在啟動 LINE 鮮奶物流機器人...
echo (首次會下載/啟動隧道，約 10~15 秒後出現網址)
echo.
python start_line_bot.py
pause
