@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv-build\Scripts\python.exe" (
    py -3 -m venv .venv-build
    if errorlevel 1 goto :error
)

call .venv-build\Scripts\activate.bat
python -m pip install --upgrade pip
if errorlevel 1 goto :error
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto :error

python -m PyInstaller --noconfirm --clean OZDEKAN.spec
if errorlevel 1 goto :error

echo.
echo EXE hazir: %CD%\dist\OZDEKAN.exe
exit /b 0

:error
echo.
echo Paketleme basarisiz. Yukaridaki hatayi kontrol edin.
exit /b 1
