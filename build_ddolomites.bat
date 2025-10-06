@echo off
title Building DDolomitesWpaper (Maguamale Edition)
echo ==============================================
echo   DDolomitesWpaper Build Script
echo   Developer: Maguamale
echo   Icon: DDolomitesWpaper.ico
echo   Logo resource: logo.png
echo ==============================================
echo.

REM Change working directory to script location
cd /d "%~dp0"

REM Clean previous build
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo Installing required packages...
pip install -U pyinstaller pystray pillow pywin32 requests tzdata

echo.
echo Packing, please wait...

pyinstaller --clean --onefile --noconsole ^
  --name DDolomitesWpaper ^
  --icon "%~dp0DDolomitesWpaper.ico" ^
  --hidden-import=pystray --hidden-import=pystray._win32 ^
  --hidden-import=win32api --hidden-import=win32gui --hidden-import=win32con ^
  --collect-submodules=pystray --collect-submodules=PIL ^
  --collect-data tzdata ^
  --add-data "logo.png;." ^
  ddolomites_wpaper.py

echo.
if exist dist\DDolomitesWpaper.exe (
    echo Build completed successfully!
    echo Output file:
    echo     %~dp0dist\DDolomitesWpaper.exe
    echo -----------------------------------------------------
    echo If the icon does not update immediately:
    echo   Run:  ie4uinit.exe -ClearIconCache
    echo   Or restart Windows Explorer
    echo -----------------------------------------------------
) else (
    echo Build failed. Please check the error log above.
)
echo.
pause
