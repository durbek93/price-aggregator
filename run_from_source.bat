@echo off
REM Только для разработки (нужен Python в PATH).
REM Для пользователей: dist\PriceAggregator.exe или установщик из installer\BUILD.txt
cd /d "%~dp0"
python app_main.py
if errorlevel 1 pause
