@echo off
title BOT TONG QUAN MAU - BO CUU HO
cd /d "%~dp0"
if not exist "%~dp0BOT_TONG_QUAN.vbs" (
  echo LOI: thieu file BOT_TONG_QUAN.vbs trong thu muc nay.
  echo Can du 4 file: CHAY.bat  APP_TONG_QUAN.ps1  BOT_TONG_QUAN.vbs  DUNG.bat
  pause
  exit /b 1
)
if not exist "%~dp0APP_TONG_QUAN.ps1" (
  echo LOI: thieu file APP_TONG_QUAN.ps1 trong thu muc nay.
  pause
  exit /b 1
)
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0APP_TONG_QUAN.ps1"
if errorlevel 1 (
  echo.
  echo LOI: khong mo duoc app. Chup man hinh dong chu do phia tren gui giang vien.
  pause
)
