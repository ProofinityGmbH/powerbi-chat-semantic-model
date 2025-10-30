@echo off
setlocal enabledelayedexpansion

echo ========================================
echo Power BI Chat - Installer Build Script
echo ========================================
echo.

:: Check if Node.js is installed
where node >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Node.js is not installed or not in PATH
    echo Please install Node.js from https://nodejs.org/
    pause
    exit /b 1
)

:: Check if .NET SDK is installed
where dotnet >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] .NET SDK is not installed or not in PATH
    echo Please install .NET SDK from https://dotnet.microsoft.com/download
    pause
    exit /b 1
)

echo [1/6] Installing npm dependencies...
call npm install
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] npm install failed
    pause
    exit /b 1
)
echo [OK] Dependencies installed
echo.

echo [2/6] Building .NET Bridge...
cd XmlaBridge
call dotnet restore
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] dotnet restore failed
    cd ..
    pause
    exit /b 1
)

call dotnet build -c Release
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] dotnet build failed
    cd ..
    pause
    exit /b 1
)
cd ..
echo [OK] .NET Bridge compiled successfully
echo.

echo [3/6] Verifying .NET Bridge DLL...
if not exist "XmlaBridge\bin\Release\net48\XmlaBridge.dll" (
    echo [ERROR] XmlaBridge.dll not found at expected location
    echo Expected: XmlaBridge\bin\Release\net48\XmlaBridge.dll
    pause
    exit /b 1
)
echo [OK] XmlaBridge.dll verified
echo.

echo [4/6] Creating published .NET assemblies folder...
if not exist "XmlaBridge-Published" mkdir XmlaBridge-Published
xcopy /Y /I /E "XmlaBridge\bin\Release\net48\*" "XmlaBridge-Published\"
echo [OK] Published assemblies copied
echo.

echo [5/6] Building Electron installer...
call npm run build
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Electron build failed
    pause
    exit /b 1
)
echo [OK] Electron installer built
echo.

echo [6/6] Verifying installer output...
if exist "dist\*.exe" (
    echo [OK] Installer created successfully!
    echo.
    echo ========================================
    echo BUILD COMPLETE!
    echo ========================================
    echo.
    echo Installer location: dist\
    echo.
    dir /B dist\*.exe
    echo.
    echo To install:
    echo 1. Run the installer from the dist\ folder
    echo 2. The installer will:
    echo    - Install Power BI Chat to Program Files
    echo    - Register as a Power BI External Tool
    echo    - Create desktop and start menu shortcuts
    echo.
    echo After installation, launch from:
    echo    - Power BI Desktop ^> External Tools ribbon
    echo    - Start Menu ^> Power BI Chat
    echo    - Desktop shortcut
    echo.
) else (
    echo [WARNING] No installer found in dist\ folder
    echo Please check the build output above for errors
)

pause
