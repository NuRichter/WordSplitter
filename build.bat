@echo off
setlocal enabledelayedexpansion

rem ---------------------------------------------------------------
rem  WordSplitter - Windows build script
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

echo [1/6] Membuat virtual environment build\venv ...
if not exist "build\venv\Scripts\python.exe" (
    %PYCMD% -m venv "build\venv"
    if !errorlevel! neq 0 (
        echo [ERROR] Pembuatan virtual environment gagal.
        exit /b 1
    )
)
set "VENV_PY=build\venv\Scripts\python.exe"

echo [2/6] Memasang dependency ...
"%VENV_PY%" -m pip install --quiet --upgrade pip
"%VENV_PY%" -m pip install --quiet -r requirements.txt
if !errorlevel! neq 0 (
    echo [ERROR] Instalasi dependency gagal.
    exit /b 1
)

echo [3/6] Menjalankan pengujian ...
set "WS_SKIP_WORD=%WS_SKIP_WORD%"
"%VENV_PY%" -m unittest discover -s tests
if !errorlevel! neq 0 (
    echo [ERROR] Ada pengujian yang gagal. Build dihentikan.
    exit /b 1
)

echo [4/6] Membersihkan hasil build sebelumnya ...
if exist "dist\WordSplitter.exe" del /q "dist\WordSplitter.exe"
if exist "build\WordSplitter" rmdir /s /q "build\WordSplitter"

echo [5/6] Menjalankan PyInstaller ...
"%VENV_PY%" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --optimize 2 ^
    --name WordSplitter ^
    --icon "%cd%\assets\wordsplitter.ico" ^
    --version-file "%cd%\packaging\version_info.txt" ^
    --paths src ^
    --distpath dist ^
    --workpath build ^
    --specpath build ^
    --add-data "%cd%\src\locales;locales" ^
    --add-data "%cd%\assets\wordsplitter.ico;assets" ^
    --hidden-import pythoncom ^
    --hidden-import pywintypes ^
    --hidden-import win32com.client.dynamic ^
    --exclude-module numpy ^
    --exclude-module pandas ^
    --exclude-module matplotlib ^
    --exclude-module PIL ^
    --exclude-module PyQt5 ^
    --exclude-module PySide6 ^
    --exclude-module unittest ^
    --exclude-module pydoc ^
    --exclude-module setuptools ^
    --exclude-module pkg_resources ^
    --exclude-module tkinter.test ^
    src\main.py
if !errorlevel! neq 0 (
    echo.
    echo [ERROR] Build gagal. Periksa keluaran PyInstaller di atas.
    exit /b 1
)

echo [6/6] Memeriksa hasil ...
if not exist "dist\WordSplitter.exe" (
    echo [ERROR] Executable tidak ditemukan setelah build.
    exit /b 1
)
echo.
echo === BUILD BERHASIL ===
echo Executable: %cd%\dist\WordSplitter.exe

endlocal
exit /b 0
