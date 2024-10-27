@echo off
setlocal enabledelayedexpansion

:: Change to the script's directory
cd /d "%~dp0"

echo Starting installation process...

:: Check if Node.js is installed
where node >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Error: Node.js is not installed or not in PATH
    echo Please install Node.js from https://nodejs.org/
    pause
    exit /b 1
)

:: Check if npm is installed
where npm >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Error: npm is not installed or not in PATH
    echo Please install Node.js from https://nodejs.org/
    pause
    exit /b 1
)

echo Installing node modules...
call npm install
if %ERRORLEVEL% neq 0 (
    echo Error: npm install failed
    pause
    exit /b 1
)
echo Node modules installed successfully.

echo Starting the application...
call npm start
if %ERRORLEVEL% neq 0 (
    echo Error: Failed to start the application
    pause
    exit /b 1
)

pause
exit /b 0