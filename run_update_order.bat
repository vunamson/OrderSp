@echo off
REM Chuyển code page sang UTF-8
chcp 65001 > nul

cd /d "C:\17track"
py main_update_order.py >> "C:\17track\update_order_log.txt" 2>&1