@echo off
chcp 65001 >nul
title ExcelAI - 安装
set "DST=%LOCALAPPDATA%\DeepSeekExcelAssistant"
echo 正在安装到 %DST% ...
if not exist "%DST%" mkdir "%DST%"
robocopy "%~dp0" "%DST%" /E /XD test dist .codex-tmp /XF install.bat setup.sed DeepSeekExcelAssistantSetup.exe >nul
if not exist "%DST%\carrier" mkdir "%DST%\carrier"
copy /y "%DST%\DeepSeekExcelAssistant-Standalone-Autoload.xlsm" "%DST%\carrier\DeepSeekExcelAssistant-Standalone-Autoload.xlsm" >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%DST%\configure-autoload.ps1" -InstallRoot "%DST%" -CarrierPath "%DST%\carrier\DeepSeekExcelAssistant-Standalone-Autoload.xlsm"
if errorlevel 1 exit /b 1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v 7c1e9a24-5d3f-4b8e-9c2a-0f6d1b4e8a5c /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v 1537f254-10aa-41d5-aed2-0a00b89da104 /t REG_SZ /d "%DST%\manifest-standalone.xml" /f >nul
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{c5e8d2f1-3b4a-4c5d-9e1f-8a2b7c6d5e4f}" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache" /v 1537f254-10aa-41d5-aed2-0a00b89da104_developer.Microsoft.Excel.Workbook /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache" /v 7c1e9a24-5d3f-4b8e-9c2a-0f6d1b4e8a5c_developer.Microsoft.Excel.Workbook /f >nul 2>&1
powershell -NoProfile -Command "$desktop = [Environment]::GetFolderPath('Desktop'); $programs = [Environment]::GetFolderPath('Programs'); foreach ($legacy in @(($desktop + '\DeepSeek Excel 助手.lnk'), ($desktop + '\DeepSeek Excel Assistant.lnk'), ($programs + '\DeepSeek Excel 助手.lnk'), ($programs + '\DeepSeek Excel Assistant.lnk'))) { Remove-Item -LiteralPath $legacy -Force -ErrorAction SilentlyContinue }; $ws = New-Object -ComObject WScript.Shell; $target = $env:WINDIR + '\System32\wscript.exe'; $args = '""%DST%\start-all.vbs""'; foreach ($path in @(($desktop + '\ExcelAI.lnk'), ($programs + '\ExcelAI.lnk'))) { $sc = $ws.CreateShortcut($path); $sc.TargetPath = $target; $sc.Arguments = $args; $sc.WorkingDirectory = '%DST%'; $sc.IconLocation = '%DST%\app.ico,0'; $sc.Save() }"
echo.
echo 安装完成。请完全退出并重新打开 Excel。
echo 本机兼容版按钮会尝试出现在“开始”选项卡；管理员集中部署版使用独立 ID，不会被本安装覆盖。
pause
