@echo off
chcp 65001 >nul
call "%~dp0config.bat"

echo === Apply translations to .ks files ===
echo.
echo Originals: %KS_DIR%
echo XLSX:      %XLSX_TRANSLATED%
echo Output:    %OUT_DIR%
echo.

python "%~dp0tyrano_l10n.py" apply "%KS_DIR%" "%XLSX_TRANSLATED%" --output "%OUT_DIR%"

echo.
echo Done! Translated files in: %OUT_DIR%
echo Make a backup of originals before copying!
pause
