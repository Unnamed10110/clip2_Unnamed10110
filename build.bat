@echo off
echo Building Clipboard Manager...

if not exist build mkdir build
cd build

cmake .. -G "MinGW Makefiles"
if errorlevel 1 (
    cmake .. -G "Visual Studio 16 2019" -A x64
    if errorlevel 1 (
        cmake .. -G "Visual Studio 17 2022" -A x64
    )
)

cmake --build . --config Release

if exist Release\ClipboardManager.exe (
    copy Release\ClipboardManager.exe ..\ClipboardManager.exe
    echo Build successful! ClipboardManager.exe created.
) else if exist ClipboardManager.exe (
    copy ClipboardManager.exe ..
    echo Build successful! ClipboardManager.exe created.
) else (
    echo Build may have failed. Check output above.
)

cd ..
pause

