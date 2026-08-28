@echo off
title Switch back to local mode
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  set "NODE=C:\Users\81019\AppData\Local\Microsoft\WinGet\Packages\OpenJS.NodeJS.LTS_Microsoft.Winget.Source_8wekyb3d8bbwe\node-v24.14.0-win-x64\node.exe"
) else (
  set "NODE=node"
)
"%NODE%" deploy.mjs --local
if exist "%LOCALAPPDATA%\DeepSeekExcelAssistant\manifest.xml" (
  copy /Y "%~dp0manifest.xml" "%LOCALAPPDATA%\DeepSeekExcelAssistant\manifest.xml" >nul
  echo Synced to installed folder.
)
echo Switched to local mode. Double-click the desktop shortcut to start server + Excel.
pause
