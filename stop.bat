@echo off
setlocal

set "PORT=5000"
set "FOUND=0"

for /f "tokens=5" %%a in ('netstat -aon ^| findstr :%PORT% ^| findstr LISTENING') do (
  set "FOUND=1"
  echo Menghentikan proses PID %%a pada port %PORT% ...
  taskkill /PID %%a /F >nul 2>&1
)

if "%FOUND%"=="0" (
  echo Tidak ada proses aktif di port %PORT%.
) else (
  echo Server berhasil dihentikan.
)

endlocal
pause
