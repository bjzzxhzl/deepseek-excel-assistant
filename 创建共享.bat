@echo off
title Create share for ExcelAI
net share DeepSeekExcel="%LOCALAPPDATA%\DeepSeekExcelAssistant" /grant:everyone,READ
if errorlevel 1 (
  echo [X] Failed to create share. Please run this file as administrator.
) else (
  echo [OK] Share created: \\localhost\DeepSeekExcel
)
pause
