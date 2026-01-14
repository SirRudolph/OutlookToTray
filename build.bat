@echo off
setlocal

cd /d "%~dp0"
set PATH=C:\msys64\mingw64\bin;%PATH%
set OUTDIR=bin

if not exist %OUTDIR% mkdir %OUTDIR%

echo Building DLL...
g++ -std=c++17 -Wall -O2 -DUNICODE -D_UNICODE -shared -static-libgcc -static-libstdc++ -o %OUTDIR%\OutlookToTray.dll OutlookToTray.Dll\OutlookToTray.Dll.cpp -lpsapi -lcomctl32
if errorlevel 1 (
    echo DLL build failed!
    pause
    exit /b 1
)

echo Building EXE resources...
windres OutlookToTray.Exe\OutlookToTray.rc -o %OUTDIR%\resources.o
if errorlevel 1 (
    echo Resource compile failed!
    pause
    exit /b 1
)

echo Building EXE...
g++ -std=c++17 -Wall -O2 -DUNICODE -D_UNICODE -mwindows -static-libgcc -static-libstdc++ -o %OUTDIR%\OutlookToTray.exe OutlookToTray.Exe\OutlookToTray.Exe.cpp %OUTDIR%\resources.o -lshell32 -lpsapi
if errorlevel 1 (
    echo EXE build failed!
    pause
    exit /b 1
)

del %OUTDIR%\resources.o 2>nul

echo.
echo Build successful!
echo Output: %OUTDIR%\OutlookToTray.exe and %OUTDIR%\OutlookToTray.dll
pause
