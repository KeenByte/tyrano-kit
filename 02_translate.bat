@echo off
chcp 65001 >nul
call "%~dp0config.bat"

echo === Machine translation ===
echo.
echo File: %XLSX%
echo.
echo Engines: yandex, google
echo Language codes: en, ru, de, fr, ja, ko, zh, es, it, pt ...
echo Full list: python translate_xlsx.py --list
echo.

set /p ENGINE="Engine [yandex]: "
if "%ENGINE%"=="" set "ENGINE=yandex"

set /p FROM_LANG="Source language [en]: "
if "%FROM_LANG%"=="" set "FROM_LANG=en"

set /p TO_LANG="Target language [ru]: "
if "%TO_LANG%"=="" set "TO_LANG=ru"

echo.
echo Engine: %ENGINE%  Translation: %FROM_LANG% -^> %TO_LANG%
echo Result: %XLSX_TRANSLATED%
echo.

python "%~dp0translate_xlsx.py" "%XLSX%" --engine %ENGINE% --from %FROM_LANG% --to %TO_LANG%

echo.
pause
