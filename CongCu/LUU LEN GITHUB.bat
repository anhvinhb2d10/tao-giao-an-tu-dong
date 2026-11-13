@echo off
chcp 65001 >nul
title Luu thay doi len GitHub
cd /d "%~dp0.."
echo ==============================================
echo    LUU THAY DOI CUA CONG CU LEN GITHUB
echo ==============================================
echo.
where git >nul 2>nul
if errorlevel 1 (
  echo May chua cai Git. Hay cai Git roi chay lai.
  echo.
  pause
  exit /b 1
)
git add -A
git diff --cached --quiet
if not errorlevel 1 (
  echo Khong co thay doi nao can luu.
  echo.
  pause
  exit /b 0
)
echo Cac file da thay doi:
git status --short
echo.
set "GHICHU="
set /p GHICHU=Ghi chu ngan ve thay doi (Enter de bo qua):
if "%GHICHU%"=="" set "GHICHU=Cap nhat cong cu"
git commit -m "%GHICHU%"
git push
echo.
if errorlevel 1 (
  echo CHUA DAY LEN DUOC. Kiem tra mang hoac dang nhap GitHub roi chay lai.
) else (
  echo Da luu len GitHub.
)
echo.
pause
