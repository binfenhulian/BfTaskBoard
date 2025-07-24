@echo off
chcp 65001 > nul
cls

echo ========================================
echo Testing BfTaskBoard Single File Publish
echo ========================================
echo.

echo Step 1: Publishing as single file...
call build.bat

echo.
echo ========================================
echo Please test the published executable:
echo 1. Copy publish\BfTaskBoard.exe to another location
echo 2. Run it without .NET installed
echo 3. Verify all features work correctly
echo ========================================
echo.

pause