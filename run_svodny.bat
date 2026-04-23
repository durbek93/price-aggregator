@echo off
setlocal EnableExtensions
chcp 65001 >nul 2>&1
pushd "%~dp0"

echo Running merged price generation...
".\.venv\Scripts\python.exe" "%~dp0svodny_price.py" --config "%~dp0config.example.yml" --inputs "%~dp0prices" --producer-matrix
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
