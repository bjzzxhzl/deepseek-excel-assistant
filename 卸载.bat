@echo off
chcp 65001 >nul
title ExcelAI - 卸载
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v 1537f254-10aa-41d5-aed2-0a00b89da104 /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{c5e8d2f1-3b4a-4c5d-9e1f-8a2b7c6d5e4f}" /f >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0configure-autoload.ps1" -InstallRoot "%~dp0" -Uninstall >nul 2>&1
del /f /q "%APPDATA%\Microsoft\Excel\XLSTART\DeepSeekExcelAssistant-Standalone-Autoload.xlsm" >nul 2>&1
del /f /q "%APPDATA%\Microsoft\Excel\XLSTART\DeepSeekExcelAssistant-Autoload.xlsx" >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache" /v 1537f254-10aa-41d5-aed2-0a00b89da104_developer.Microsoft.Excel.Workbook /f >nul 2>&1
powershell -NoProfile -Command "$desktop = [Environment]::GetFolderPath('Desktop'); $programs = [Environment]::GetFolderPath('Programs'); foreach ($link in @(($desktop + '\ExcelAI.lnk'), ($programs + '\ExcelAI.lnk'))) { Remove-Item -LiteralPath $link -Force -ErrorAction SilentlyContinue }" >nul 2>&1
echo 已移除清单注册和隐藏启动载体。重新启动 Excel 后生效。
pause
