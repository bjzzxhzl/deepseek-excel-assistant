@echo off
chcp 65001 >nul
title 生成 ExcelAI 安装包
cd /d "%~dp0"
set "ISCC=C:\Program Files (x86)\Inno Setup 7\ISCC.exe"
if not exist "%ISCC%" set "ISCC=C:\Program Files\Inno Setup 7\ISCC.exe"
if not exist "%ISCC%" set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" (
  echo 未找到 Inno Setup 编译器 ISCC.exe，请安装 Inno Setup 6/7 后重试。
  pause
  exit /b 1
)
echo 正在生成图标...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0generate-icon-assets.ps1" || exit /b 1
echo 正在生成隐藏启动载体...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0generate-autoload-workbook.ps1" || exit /b 1
echo 正在编译 Inno Setup 安装包...
"%ISCC%" "%~dp0setup.iss"
if exist "%~dp0dist\ExcelAI-Standalone-Setup-v0.40.exe" (
  echo.
  echo 完成！安装包位置：%~dp0dist\ExcelAI-Standalone-Setup-v0.40.exe
) else (
  echo 编译未成功，请检查上方日志。
  exit /b 1
)
pause
