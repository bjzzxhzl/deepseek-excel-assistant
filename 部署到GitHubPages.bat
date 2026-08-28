@echo off
title Deploy ExcelAI to GitHub Pages
cd /d "%~dp0"
set /p URL=Enter your GitHub Pages URL (e.g. https://username.github.io/deepseek-excel-assistant, no trailing slash): 
where node >nul 2>nul
if errorlevel 1 (
  set "NODE=C:\Users\81019\AppData\Local\Microsoft\WinGet\Packages\OpenJS.NodeJS.LTS_Microsoft.Winget.Source_8wekyb3d8bbwe\node-v24.14.0-win-x64\node.exe"
) else (
  set "NODE=node"
)
echo.
"%NODE%" deploy.mjs "%URL%"
if not exist "%~dp0pages" mkdir "%~dp0pages"
for %%f in (taskpane.html app.js app.css icon16.png icon32.png icon64.png icon80.png manifest.xml) do (
  copy /Y "%~dp0%%f" "%~dp0pages\%%f" >nul
)
if exist "%LOCALAPPDATA%\DeepSeekExcelAssistant" (
  for %%f in (taskpane.html app.js app.css icon16.png icon32.png icon64.png icon80.png manifest.xml manifest-standalone.xml) do (
    copy /Y "%~dp0%%f" "%LOCALAPPDATA%\DeepSeekExcelAssistant\%%f" >nul
  )
  echo Synced runtime and both manifests to the installed folder.
)
echo.
echo Next steps:
echo   1) Upload the 8 files in the pages folder to the ROOT of your GitHub repo
echo   2) Repo Settings ^> Pages ^> main branch ^> Save, wait 1 minute
echo   3) Fully restart Excel, then click the ExcelAI button
echo.
echo No local server needed anymore.
pause
