@echo off
chcp 65001 >nul
title Kiem tra cong cu soan giao an
echo ==============================================
echo    KIEM TRA CONG CU SAU KHI SUA
echo ==============================================
echo  Chay lai tuan 10 va 11 roi so voi ban moc da luu.
echo  Mat khoang 1-2 phut. Nho dong Word truoc khi chay.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kiem_tra.ps1"
echo.
echo  Neu cac cho khac la dung y co muon, go lenh sau de lay lam ban moc moi:
echo     powershell -ExecutionPolicy Bypass -File "%~dp0kiem_tra.ps1" -LuuMoc
echo.
pause
