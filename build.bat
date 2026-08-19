@echo off
setlocal enabledelayedexpansion

rem ---------------------------------------------------------------
rem  WordSplitter - Windows 11 build script
rem  Produces dist\WordSplitter.exe as a single self contained file.
rem ---------------------------------------------------------------

cd /d "%~dp0"

echo.
echo === WordSplitter build ===
echo.

where py >nul 2>&1
if %errorlevel%==0 (
    set "PYCMD=py -3"
) else (
    where python >nul 2>&1
    if !errorlevel! neq 0 (
        echo [ERROR] Python tidak ditemukan pada PATH.
        echo         Install Python 3.11 atau lebih baru dari python.org, lalu ulangi.
        exit /b 1
    )
    set "PYCMD=python"
)

echo [1/5] Membuat virtual environment build\venv ...
if not exist "build\venv" (
    %PYCMD% -m venv "build\venv"
    if !errorlevel! neq 0 (
        echo [ERROR] Pembuatan virtual environment gagal.
        exit /b 1
    )
)

set "VENV_PY=build\venv\Scripts\python.exe"

echo [2/5] Memutakhirkan pip ...
"%VENV_PY%" -m pip install --upgrade pip >nul
if !errorlevel! neq 0 (
    echo [ERROR] Pemutakhiran pip gagal.
    exit /b 1
)

echo [3/5] Memasang dependency ...
"%VENV_PY%" -m pip install -r requirements.txt
if !errorlevel! neq 0 (
    echo [ERROR] Instalasi dependency gagal.
    exit /b 1
)

echo [4/5] Membersihkan hasil build sebelumnya ...
if exist "dist\WordSplitter.exe" del /q "dist\WordSplitter.exe"
if exist "build\WordSplitter" rmdir /s /q "build\WordSplitter"

echo [5/5] Menjalankan PyInstaller ...
"%VENV_PY%" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name WordSplitter ^
    --paths src ^
    --distpath dist ^
    --workpath build ^
    --specpath build ^
    --hidden-import win32com.client ^
    --hidden-import win32timezone ^
    --hidden-import pythoncom ^
    --hidden-import pywintypes ^
    --exclude-module numpy ^
    --exclude-module pandas ^
    --exclude-module matplotlib ^
    --exclude-module PyQt5 ^
    --exclude-module PySide6 ^
    src\main.py

if !errorlevel! neq 0 (
    echo.
    echo [ERROR] Build gagal. Periksa keluaran PyInstaller di atas.
    exit /b 1
)

echo.
if exist "dist\WordSplitter.exe" (
    echo === BUILD BERHASIL ===
    echo Executable: %cd%\dist\WordSplitter.exe
) else (
    echo [ERROR] Executable tidak ditemukan setelah build.
    exit /b 1
)

endlocal
exit /b 0
