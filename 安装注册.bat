@echo off
chcp 65001 >nul
title ExcelAI - 配置安装目录
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0configure-autoload.ps1" -InstallRoot "%~dp0" -CarrierPath "%~dp0DeepSeekExcelAssistant-Standalone-Autoload.xlsm"
if errorlevel 1 exit /b 1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v 7c1e9a24-5d3f-4b8e-9c2a-0f6d1b4e8a5c /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v 1537f254-10aa-41d5-aed2-0a00b89da104 /t REG_SZ /d "%~dp0manifest-standalone.xml" /f >nul
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{c5e8d2f1-3b4a-4c5d-9e1f-8a2b7c6d5e4f}" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache" /v 7c1e9a24-5d3f-4b8e-9c2a-0f6d1b4e8a5c_developer.Microsoft.Excel.Workbook /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache" /v 1537f254-10aa-41d5-aed2-0a00b89da104_developer.Microsoft.Excel.Workbook /f >nul 2>&1
echo 配置完成。完全退出并重新打开 Excel 后，ExcelAI 按钮会自动出现。
pause
