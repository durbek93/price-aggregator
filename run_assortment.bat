@echo off
setlocal EnableExtensions
chcp 65001 >nul 2>&1
pushd "%~dp0"

echo Running assortment count for prices...
".\.venv\Scripts\python.exe" "%~dp0assortment_prices.py" --config "%~dp0config.example.yml" --inputs "%~dp0prices"
set "ERR=%ERRORLEVEL%"

if not "%ERR%"=="0" goto :err

echo.
echo Done. Output file was saved to the output folder.
goto :end

:err
echo.
echo Failed. Close Excel if output file is open, then run again.

:end
popd
pause
endlocal

