@echo off
REM Builds SevakJoda.exe on Windows. Run this from a Windows machine
REM (or the GitHub Actions workflow does this automatically).
REM Requires Python 3.11+ installed and on PATH.

setlocal

echo Creating virtual environment...
python -m venv build_env
call build_env\Scripts\activate.bat

echo Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt

echo Building SevakJoda.exe with PyInstaller...
pyinstaller sevak_joda.spec --noconfirm

echo.
echo Done. Copy dist\SevakJoda.exe and credentials.json to the target machine.
echo A "photos" folder will be created automatically next to the exe on first run.

endlocal
