@echo off
for %%I in ("%~dp0.") do set "APPDIR=%%~fI"
net share DeepSeekExcel /delete /y >nul 2>&1
if /i "%~1"=="remove" exit /b 0
net share DeepSeekExcel="%APPDIR%" /grant:everyone,READ >nul 2>&1
if errorlevel 1 exit /b 1
exit /b 0
