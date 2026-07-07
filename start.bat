@echo off
setlocal

cd /d "%~dp0"

set "PYTHON_EXE=C:\Users\Lenovo\AppData\Local\Programs\Python\Python312\python.exe"

if not exist "%PYTHON_EXE%" (
  echo Python tidak ditemukan di:
  echo %PYTHON_EXE%
  echo.
  echo Silakan sesuaikan path Python di file start.bat.
  pause
  exit /b 1
)

echo Menjalankan aplikasi di http://127.0.0.1:5000 ...
start "Aplikasi Akuisisi" cmd /k ""%PYTHON_EXE%" app.py"

endlocal
