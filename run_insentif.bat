@echo off
setlocal
title Insentif dan Performance
cd /d "%~dp0"

if not exist "data" (
    echo [GAGAL] Folder "data" tidak ditemukan di %CD%
    echo Isi folder data dengan file export sales dan file master.
    echo.
    pause
    exit /b 1
)

:menu
cls
echo ==============================================
echo    PEMBUAT INSENTIF / PERFORMANCE PER CHANNEL
echo ==============================================
echo.
echo    1. Modern Trade
echo    2. Direct        (belum tersedia)
echo    3. Project       (belum tersedia)
echo    4. Indirect      (belum tersedia)
echo    0. Transform saja (tanpa insentif)
echo.
set "pilihan="
set /p "pilihan=Pilih [0-4]: "

if "%pilihan%"=="1" set "CH=modern_trade" & goto jalan
if "%pilihan%"=="0" set "CH=none" & goto jalan
if "%pilihan%"=="2" goto belum
if "%pilihan%"=="3" goto belum
if "%pilihan%"=="4" goto belum

echo.
echo Pilihan tidak valid.
pause
goto menu

:belum
echo.
echo Script untuk channel tersebut belum tersedia, akan ditambahkan.
pause
goto menu

:jalan
echo.
echo Folder kerja: %CD%
echo.
python run_pipeline.py --channel %CH%
if errorlevel 1 (
    echo.
    echo [GAGAL] Cek pesan error di atas.
) else (
    echo.
    echo [SELESAI] Hasil ada di folder "output".
)
echo.
pause
endlocal
