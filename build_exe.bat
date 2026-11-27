@echo off
REM Build CMW500_LTE_Tool.exe directly into this folder (no dist subfolder).
REM Usage: double-click, or run "build_exe.bat RUN" from a console.
cd /d "%~dp0"
if "%~1"=="RUN" goto MAIN
start cmd /k "%~nx0" RUN
exit /b

:MAIN
tasklist /FI "IMAGENAME eq CMW500_LTE_Tool.exe" 2>nul | find /I "CMW500_LTE_Tool.exe" >nul
if not errorlevel 1 goto BUSY

echo [1/3] Building exe into project root ...
".venv38\Scripts\pyinstaller.exe" --noconfirm --clean --distpath . CMW500_LTE_Tool.spec
if errorlevel 1 goto FAIL

echo.
echo [2/3] Running frozen selftest ...
CMW500_LTE_Tool.exe --selftest
if errorlevel 1 goto FAIL

echo.
echo [3/3] Removing build cache ...
if exist build rmdir /s /q build
if exist selftest_result.txt del /q selftest_result.txt

echo.
echo Build OK: CMW500_LTE_Tool.exe in this folder.
exit /b 0

:BUSY
echo.
echo [ERROR] CMW500_LTE_Tool.exe is running. Close it, then build again.
exit /b 1

:FAIL
echo.
echo [ERROR] Build or selftest failed. See messages above.
exit /b 1
