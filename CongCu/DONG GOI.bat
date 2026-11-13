@echo off
chcp 65001 >nul
title Dong goi cong cu de gui nguoi khac
echo ==============================================
echo    DONG GOI CONG CU DE GUI NGUOI KHAC
echo ==============================================
echo  Goi khong chua giao an, file mon hoc, ten truong, ten giao vien.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0dong_goi.ps1"
echo.
pause
