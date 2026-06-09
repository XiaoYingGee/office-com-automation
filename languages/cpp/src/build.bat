@echo off
REM build.bat - Build Excel COM automation C++ project
REM Uses MSVC cl.exe with #import for Excel type library

setlocal

REM Setup MSVC environment
call "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvarsall.bat" x64

REM Navigate to source directory
cd /d "%~dp0"

REM Create output directory
if not exist out mkdir out

REM Clean cached type library headers to force regeneration
del /Q *.tlh *.tli *.obj 2>NUL

echo.
echo === Building Excel COM Automation backend (C++) ===
echo.

cl.exe /nologo /EHsc /std:c++17 /W3 /O2 /bigobj ^
    /DWIN32 /D_UNICODE /DUNICODE ^
    /Fe:excel-ops-cpp.exe ^
    main.cpp ops.cpp excel_com.cpp ^
    /link ole32.lib oleaut32.lib uuid.lib

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo === BUILD FAILED ===
    exit /b 1
)

echo.
echo === BUILD SUCCEEDED ===
echo Output: %~dp0excel-ops-cpp.exe
echo.
echo Usage: one OpRequest JSON on stdin -^> one OpResponse JSON on stdout
