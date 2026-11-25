@echo off
cd /d "%~dp0"
if "%~1"=="RUN" goto MAIN
start cmd /k "%~nx0" RUN
exit /b
:MAIN
"D:\Program Files\Python\Python314\python.exe" "%~dp0cmw500_lte_tx_tool.py"
echo.
echo Tool exited. Type exit to close this window.
