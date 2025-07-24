@echo off
chcp 65001 > nul
cls

echo ========================================
echo BfTaskBoard Build Script
echo ========================================
echo.
echo 1. Build only (for development)
echo 2. Publish as single .exe file
echo.
set /p choice="Select option (1-2): "

if "%choice%"=="1" goto build
if "%choice%"=="2" goto publish
echo Invalid choice!
pause
exit /b

:build
echo.
echo Building Release version...
dotnet build -c Release

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo BUILD SUCCESSFUL!
    echo ========================================
    echo.
    echo Executable location:
    echo bin\Release\net6.0-windows\BfTaskBoard.exe
    echo.
    echo To run: 
    echo bin\Release\net6.0-windows\BfTaskBoard.exe
    echo.
) else (
    echo.
    echo ========================================
    echo BUILD FAILED!
    echo ========================================
    echo.
    echo Please check the error messages above.
    echo.
)
pause
exit /b

:publish
echo.
echo Publishing as single executable file...
echo This may take a few minutes...
echo.

REM Clean previous publish
if exist "publish" rmdir /s /q "publish"

REM Publish as single file
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -o publish

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo PUBLISH SUCCESSFUL!
    echo ========================================
    echo.
    echo Single executable location:
    echo %cd%\publish\BfTaskBoard.exe
    echo.
    
    REM Get file size
    for %%I in ("publish\BfTaskBoard.exe") do set size=%%~zI
    set /a sizeMB=%size%/1048576
    echo File size: %sizeMB% MB
    echo.
    echo This file can be copied to any Windows 10/11 computer and run directly.
    echo No .NET runtime installation required!
    echo.
    
    REM Copy favicon if exists
    if exist "favicon.ico" (
        copy /Y "favicon.ico" "publish\favicon.ico" > nul
    )
    
    echo Opening publish folder...
    start "" "publish"
) else (
    echo.
    echo ========================================
    echo PUBLISH FAILED!
    echo ========================================
    echo.
    echo Please check the error messages above.
    echo Make sure you have .NET 6.0 SDK installed.
    echo.
)
pause
exit /b