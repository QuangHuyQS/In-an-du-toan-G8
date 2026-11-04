@echo off
chcp 65001 >nul
rem Mo thang form In an du toan G8 (khong can mo file .xlsm bang tay)
rem Go dau "tai tu Internet" (Zalo, email, zip) de Excel khong chan macro - chi trong thu muc nay
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath '%~dp0.' -File | Unblock-File" >nul 2>&1
start "" wscript.exe //nologo "%~dp0MO_FORM_IN_AN.vbs"
