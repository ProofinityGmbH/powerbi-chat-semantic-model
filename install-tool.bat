@echo off
echo ========================================
echo Power BI Chat - External Tool Installer
echo ========================================
echo.
echo This script will register Power BI Chat as a Power BI External Tool.
echo You need to run this as Administrator.
echo.
pause

echo Copying registration file...
copy "%~dp0PowerBIChat.pbitool.json" "C:\Program Files (x86)\Common Files\Microsoft Shared\Power BI Desktop\External Tools\PowerBIChat.pbitool.json"

if errorlevel 1 (
    echo.
    echo ERROR: Failed to copy file. Make sure you're running as Administrator!
    echo.
    pause
    exit /b 1
)

echo.
echo SUCCESS! Power BI Chat has been registered.
echo.
echo Next steps:
echo 1. Restart Power BI Desktop
echo 2. Open any .pbix file
echo 3. Look for "Power BI Chat" in the External Tools ribbon
echo.
pause
