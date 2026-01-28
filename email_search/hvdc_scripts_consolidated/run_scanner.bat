@echo off
cls
chcp 65001 > nul
echo ================================================================
echo       PST Scanner v6.0 - Simple Run
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
echo Starting PST Scanner...
echo.
python outlook_pst_scanner.py --pst "%PST_PATH%" --auto
echo.
echo ================================================================
pause
