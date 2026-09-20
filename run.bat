@echo off
rem Print600.ru - zapusk bez sborki (nuzhen Python 3.9+)
cd /d "%~dp0"
set "PY=python"
where py >nul 2>nul && set "PY=py -3"
%PY% -m pip install -q -r requirements.txt || goto err
%PY% print600_vvod.py
exit /b 0
:err
echo.
echo Could not install packages. Install Python 3.9+ from python.org
echo (tick "Add python.exe to PATH") and run this file again.
pause
exit /b 1
