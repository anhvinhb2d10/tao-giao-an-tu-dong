@echo off
chcp 65001 >nul
title Soan giao an tu dong - Lop 2C
echo ==============================================
echo    SOAN GIAO AN TUAN TU DONG - LOP 2C
echo ==============================================
echo  Nho dong file Word ket qua cu neu dang mo.
echo.
set "CAUHINH="
set DEM=0
for %%f in ("%~dp0cauhinh*.json") do set /a DEM+=1
if %DEM% GTR 1 (
  echo Co nhieu file cau hinh:
  for %%f in ("%~dp0cauhinh*.json") do echo    %%~nxf
  set /p CAUHINH=Nhap ten file cau hinh can dung, hoac Enter de dung cauhinh.json:
  echo.
)
if "%CAUHINH%"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tao_giao_an.ps1"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tao_giao_an.ps1" -CauHinh "%CAUHINH%"
)
echo.
pause
