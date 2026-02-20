@echo off
REM Try pythonw first (no console window), fallback to python
where pythonw >nul 2>&1 && (
    start "" pythonw "%~dp0launcher.pyw"
    exit /b
)
where python >nul 2>&1 && (
    start "" python "%~dp0launcher.pyw"
    exit /b
)
where py >nul 2>&1 && (
    start "" py "%~dp0launcher.pyw"
    exit /b
)
echo Python ne najden! Ustanovi Python i dobav v PATH.
pause
