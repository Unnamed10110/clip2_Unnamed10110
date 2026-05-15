@echo off
setlocal EnableExtensions
set "ROOT=%~dp0"
cd /d "%ROOT%"

echo Building clip2...
if not exist build mkdir build
cd build

if exist CMakeCache.txt (
    cmake ..
    if errorlevel 1 (
        echo Reconfigure failed.
        cd ..
        exit /b 1
    )
) else (
    cmake .. -G "MinGW Makefiles" 2>nul
    if errorlevel 1 (
        cmake .. -G "Visual Studio 17 2022" -A x64
        if errorlevel 1 (
            cmake .. -G "Visual Studio 16 2019" -A x64
            if errorlevel 1 (
                echo CMake configuration failed. Install MinGW, CMake, and/or Visual Studio 2022/2019.
                cd ..
                exit /b 1
            )
        )
    )
)

cmake --build . --config Release
if errorlevel 1 (
    echo Build failed.
    cd ..
    exit /b 1
)

set "OUT="
if exist "Release\clip2.exe" set "OUT=Release\clip2.exe"
if not defined OUT if exist "clip2.exe" set "OUT=clip2.exe"

if defined OUT (
    copy /Y "%OUT%" "%ROOT%clip2.exe" >nul
    echo.
    echo Build successful: %ROOT%clip2.exe
    echo   ^(also: %CD%\%OUT%^)
) else (
    echo Build finished but clip2.exe was not found in expected locations.
    echo Look under: %CD%
    cd ..
    exit /b 1
)

cd ..
exit /b 0
