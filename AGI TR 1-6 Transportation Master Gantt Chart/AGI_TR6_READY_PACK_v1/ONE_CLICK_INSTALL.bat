@echo off
setlocal
echo [AGI_TR6] One-click install: create READY XLSM with VBA embedded
echo.
echo If it fails, enable:
echo   Excel > File > Options > Trust Center > Trust Center Settings...
echo   > Macro Settings > [x] Trust access to the VBA project object model
echo.
cscript //nologo "%~dp0ONE_CLICK_INSTALL.vbs"
echo.
pause