@echo off
chcp 65001 >nul
title ExcelAI - Local Server
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  if exist "%ProgramFiles%\nodejs\node.exe" (
    set "NODE=%ProgramFiles%\nodejs\node.exe"
  ) else (
    echo [X] Node.js not found. Please install Node.js LTS from https://nodejs.org and run again.
    pause
    exit /b 1
  )
) else (
  set "NODE=node"
)
echo Starting ExcelAI server (http://127.0.0.1:8090) ...
echo Keep this window open while using the add-in.
echo.
"%NODE%" server.mjs
pause
