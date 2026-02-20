@echo off
chcp 65001 >nul
call "%~dp0config.bat"

echo === Extract strings from .ks files ===
echo.
echo KS folder: %KS_DIR%
echo Output:    %XLSX%
echo.

python "%~dp0tyrano_l10n.py" extract "%KS_DIR%" --output "%XLSX%"

echo.
echo Done!
pause
