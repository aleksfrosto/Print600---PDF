@echo off
rem Print600.ru - sborka odnogo EXE-faila (nuzhen Python 3.9+)
cd /d "%~dp0"
set "PY=python"
where py >nul 2>nul && set "PY=py -3"
if not exist .venv-build (
  %PY% -m venv .venv-build || goto err
)
call .venv-build\Scripts\activate.bat
python -m pip install -q --upgrade pip
python -m pip install -q -r requirements.txt pyinstaller || goto err
python -m PyInstaller --noconfirm --clean --onefile --windowed --collect-all pymupdf --collect-all tkinterdnd2 --collect-submodules fontTools.ttLib.tables --icon print600.ico --version-file version_info.txt --name "Print600.ru" print600_vvod.py || goto err
echo.
echo Done: dist\Print600.ru.exe
pause
exit /b 0
:err
echo.
echo Build failed. See messages above.
pause
exit /b 1
