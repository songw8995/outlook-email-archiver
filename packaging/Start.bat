@echo off
setlocal
cd /d "%~dp0"
set "PYTHONHOME="
set "PYTHONPATH="
set "PYTHONNOUSERSITE=1"
set "PYTHONUTF8=1"
if not exist "runtime\pythonw.exe" goto missing
if /I "%~1"=="--self-test" goto test
"runtime\pythonw.exe" "portable_launch.pyw"
if errorlevel 1 goto failed
exit /b 0
:test
"runtime\python.exe" -c "from outlook_archiver.gui import App; a=App(); a.update_idletasks(); a.destroy(); print('PORTABLE_START_OK')"
if errorlevel 1 goto failed
exit /b 0
:missing
echo Runtime missing. Extract the COMPLETE ZIP before starting.
:failed
echo Startup failed. See the Chinese error dialog and startup log.
pause
exit /b 1
