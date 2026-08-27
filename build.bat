@echo off
setlocal EnableExtensions
set "ROOT=%~dp0"
cd /d "%ROOT%"

echo Building clip2...
if not exist build mkdir build

rem ---------------------------------------------------------------------------------
rem Reuse the existing CMake cache when it still works, but never let a bad one stop
rem the build. A CMakeCache.txt records the absolute source/binary paths and the
rem generator it was created for, so one that came from another checkout (this repo
rem has historically committed build/), or that names a Visual Studio version which
rem is not installed here, makes every later configure fail. Wipe it and start over
rem instead of giving up.
rem ---------------------------------------------------------------------------------
if exist "build\CMakeCache.txt" (
    cmake -S . -B build >nul 2>&1
    if not errorlevel 1 goto build
    echo Existing CMake cache is unusable - reconfiguring from scratch.
)

rem Generators in order of preference. Each attempt wipes the cache first: a failed
rem configure still leaves one behind, and CMake then rejects a different generator
rem with a mismatch error instead of trying it.
call :try "MinGW Makefiles" ""
if defined OK goto build
call :try "Visual Studio 18 2026" "-A x64"
if defined OK goto build
call :try "Visual Studio 17 2022" "-A x64"
if defined OK goto build
call :try "Visual Studio 16 2019" "-A x64"
if defined OK goto build

echo.
echo CMake configuration failed. Install CMake plus either MinGW or Visual Studio
echo ^(2019 or newer^), then run build.bat again.
exit /b 1

:build
cmake --build build --config Release
if errorlevel 1 (
    echo Build failed.
    exit /b 1
)

set "OUT="
if exist "build\Release\clip2.exe" set "OUT=build\Release\clip2.exe"
if not defined OUT if exist "build\clip2.exe" set "OUT=build\clip2.exe"

if not defined OUT (
    echo Build finished but clip2.exe was not found under build\.
    exit /b 1
)

copy /Y "%OUT%" "%ROOT%clip2.exe" >nul
echo.
echo Build successful: %ROOT%clip2.exe
echo   ^(also: %ROOT%%OUT%^)
exit /b 0

rem --- helpers ---------------------------------------------------------------------

:try
call :wipe
echo Configuring with %~1 ...
cmake -S . -B build -G %1 %~2 >nul 2>&1
if errorlevel 1 (
    set "OK="
    goto :eof
)
set "OK=1"
goto :eof

:wipe
if exist "build\CMakeCache.txt" del /q "build\CMakeCache.txt" >nul 2>&1
if exist "build\CMakeFiles" rmdir /s /q "build\CMakeFiles" >nul 2>&1
goto :eof
