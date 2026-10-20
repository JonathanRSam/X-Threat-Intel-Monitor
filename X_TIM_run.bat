@echo off
setlocal
cd /d "%~dp0"
echo Starting X TIM daily processor...
echo Folder: %CD%
echo.
py -u X_TIM_script.py %*
set EXITCODE=%ERRORLEVEL%
echo.
if %EXITCODE% NEQ 0 (
  echo X TIM finished with ERRORLEVEL %EXITCODE%.
  echo Check X_TIM_logs\X_TIM_latest_run_notes.txt and the console message above.
) else (
  echo X TIM finished successfully.
)
echo.
pause
exit /b %EXITCODE%
