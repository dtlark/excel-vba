@echo off
rem Double-click this file to run raw_invoices.py.
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 raw_invoices.py
) else (
    python raw_invoices.py
)
echo.
pause
