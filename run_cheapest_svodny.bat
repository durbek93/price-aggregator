@echo off
setlocal EnableExtensions
chcp 65001 >nul 2>&1
pushd "%~dp0"

echo Running min-price highlight for summary files...
".\.venv\Scripts\python.exe" "%~dp0cheapest_svodny_prices.py" --inputs "%~dp0output"
set "ERR=%ERRORLEVEL%"

if not "%ERR%"=="0" goto :err

echo.
echo Done. Min-price cells were highlighted in summary files.
goto :end

:err
echo.
echo Failed. Close Excel if output file is open, then run again.

:end
popd
pause
endlocal
