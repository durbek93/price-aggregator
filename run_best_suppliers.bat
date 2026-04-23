@echo off
setlocal EnableExtensions
chcp 65001 >nul 2>&1
pushd "%~dp0"

if exist "%~dp0.venv\Scripts\python.exe" (
  "%~dp0.venv\Scripts\python.exe" "%~dp0best_suppliers_analysis.py" --inputs-dir "%~dp0output" %*
) else (
  python "%~dp0best_suppliers_analysis.py" --inputs-dir "%~dp0output" %*
)
set "ERR=%ERRORLEVEL%"

if not "%ERR%"=="0" goto :err
goto :end

:err
echo.
if "%ERR%"=="1" (
  echo Ошибка 1: папка output не найдена или в ней нет подходящего XLSX. Укажите --file «путь» или --inputs-dir «папка».
) else if "%ERR%"=="3" (
  echo Ошибка 3: файл открыт, но не найдены целевые препараты с ценами. Нужен сводный прайс ^(матрица^), см. TARGET_PREP_NAMES в best_suppliers_analysis.py.
) else (
  echo Ошибка %ERR%. Укажите --file «путь к сводному прайсу» или положите XLSX в папку output.
)

:end
popd
pause
endlocal
