@echo off
setlocal
cd /d "%~dp0"
title OZDEKAN Windows Paketleme

echo OZDEKAN paketleme baslatiliyor...
echo Proje klasoru: %CD%
echo.

if not exist "launcher.py" (
    echo HATA: launcher.py bulunamadi. Tum proje klasorunu Windows'a kopyalayin.
    goto :error
)
if not exist "OZDEKAN.spec" (
    echo HATA: OZDEKAN.spec bulunamadi.
    goto :error
)
if not exist "static" (
    echo HATA: static klasoru bulunamadi.
    goto :error
)

if not exist ".venv-build\Scripts\python.exe" (
    echo Build ortami olusturuluyor...
    py -3 -m venv .venv-build
    if errorlevel 1 goto :error
)

call .venv-build\Scripts\activate.bat
python -m pip install --upgrade pip
if errorlevel 1 goto :error

echo Gerekli paketler kuruluyor...
python -m pip install fastapi "uvicorn[standard]" sqlalchemy databases pydantic psycopg2-binary "pymodbus[serial]" pyserial pyserial-asyncio "passlib[bcrypt]" python-multipart pandas xlsxwriter pyinstaller
if errorlevel 1 goto :error

echo EXE olusturuluyor...
python -m PyInstaller --noconfirm --clean OZDEKAN.spec
if errorlevel 1 goto :error

echo.
echo EXE hazir: %CD%\dist\OZDEKAN.exe
echo.
pause
exit /b 0

:error
echo.
echo Paketleme basarisiz. Yukaridaki hatayi kontrol edin.
echo Pencere artik otomatik kapanmayacak.
echo.
pause
exit /b 1
