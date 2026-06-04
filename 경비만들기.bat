@echo off
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Run-ExpenseApp.ps1"
if errorlevel 1 pause