@echo off
rem Mo cua so soan giao an (giao dien). Neu cua so khong hien, dung "TAO GIAO AN.bat" thay the.
start "" powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0giao_dien.ps1"
