@echo off
chcp 65001 >nul
REM ============================================================
REM  本機派車 一鍵執行
REM  雙擊本檔：用本機正常網路跑派車路線（繞開 Render 雲端卡網逾時），
REM  報表/地圖/派車單產到 桌面/當日車輛報表/今天/，
REM  若 .env 已設 LINE_USER_ID 則自動把摘要 push 回你的 LINE。
REM
REM  資料來源：桌面『路線規劃/每日配送.xlsx』
REM  （要強制分車請改用 plan_and_push_line.py --auto / --vehicles N）
REM ============================================================
set HERE=%~dp0
cd /d "%HERE%"

echo 正在用本機網路跑派車路線...
echo (若首次會呼叫 Google 地理編碼/距離，約數十秒)
echo.

python plan_and_push_line.py

echo.
echo ============================================================
echo  跑完了。報表在：桌面\當日車輛報表\今天\
echo  若 .env 有 LINE_USER_ID，摘要已 push 到你的 LINE。
echo ============================================================
pause
