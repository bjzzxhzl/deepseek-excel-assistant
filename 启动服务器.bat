@echo off
chcp 65001 >nul
title ExcelAI - 本地服务器
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  set "NODE=C:\Users\81019\AppData\Local\Microsoft\WinGet\Packages\OpenJS.NodeJS.LTS_Microsoft.Winget.Source_8wekyb3d8bbwe\node-v24.14.0-win-x64\node.exe"
) else (
  set "NODE=node"
)
echo 正在启动 ExcelAI 服务器 (http://127.0.0.1:8090) ...
echo 请保持本窗口开启，使用期间不要关闭。
echo 关闭本窗口即停止服务器。
echo.
"%NODE%" server.mjs
pause
