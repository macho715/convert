@echo off
chcp 65001 > nul
echo ================================================================
echo       PST Scanner v6.0 - Quick Run (2025 June)
echo ================================================================
echo.

rem ========== Outlook 강제 종료 (1차) ==========
echo Closing Outlook...
taskkill /F /IM outlook.exe /T 2>nul
if %ERRORLEVEL% EQU 0 (
    echo Outlook closed successfully.
) else (
    echo Outlook is not running.
)
timeout /t 2 /nobreak >nul
echo.
rem ===========================================

rem ========== PST 경로 하드코딩 ==========
set PST_PATH=C:\Users\SAMSUNG\Documents\Outlook 파일\minkyu.cha@samsung.comswe - outlook.RECOVERED.20251002-092839.pst
rem ========================================

echo PST File: %PST_PATH%
echo.

echo.
echo Settings:
echo    - Start date: 2025-06-01
echo    - End date: 2025-06-30
echo    - Batch size: 5,000 (Standard)
echo    - Folders: Select in program
echo.

set /p CONFIRM="Continue with these settings? (y/n): "
if /i not "%CONFIRM%"=="y" (
    echo Cancelled.
    pause
    exit /b 0
)

echo.
echo [Running PST Scanner...]
echo.

rem Run Python script with PST path
python outlook_pst_scanner.py --pst "%PST_PATH%" --start 2025-06-01 --end 2025-06-30 --folders all --batch-size 5000

echo.
echo ================================================================
echo       Completed
echo ================================================================
pause
