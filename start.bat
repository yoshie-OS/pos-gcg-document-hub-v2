@echo off
REM GCG Document Hub - Windows Start Script
REM This script checks dependencies and starts the development server

echo ==========================================
echo   GCG Document Hub - Starting Server
echo ==========================================
echo.

REM Check if Node.js is installed
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Node.js is not installed!
    echo Please install Node.js from: https://nodejs.org/
    pause
    exit /b 1
)

REM Check if Python is installed
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed!
    echo Please install Python from: https://www.python.org/
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('node --version') do set NODE_VERSION=%%i
for /f "tokens=2" %%i in ('python --version') do set PYTHON_VERSION=%%i

echo [OK] Node.js %NODE_VERSION% detected
echo [OK] Python %PYTHON_VERSION% detected
echo.

REM Check if node_modules exists
if not exist "node_modules\" (
    echo [INFO] Installing Node.js dependencies...
    call npm install
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to install Node.js dependencies
        pause
        exit /b 1
    )
    echo [OK] Node.js dependencies installed
    echo.
) else (
    echo [OK] Node.js dependencies found
)

REM Check if Python virtual environment exists
if not exist "venv\" (
    echo [INFO] Creating Python virtual environment...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create virtual environment
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created
    echo.
)

REM Activate virtual environment and install dependencies
echo [INFO] Installing Python dependencies...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip --quiet
pip install -r backend\requirements.txt --quiet
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install Python dependencies
    pause
    exit /b 1
)
echo [OK] Python dependencies installed
echo.

REM Check if .env file exists
if not exist ".env" (
    echo [WARNING] .env file not found
    if exist ".env.example" (
        echo Creating .env from .env.example...
        copy .env.example .env >nul
        echo [OK] .env file created - please configure it
    ) else (
        echo [ERROR] No .env.example found - you may need to create .env manually
    )
    echo.
)

echo ==========================================
echo [START] Starting Development Server...
echo ==========================================
echo.
echo Frontend: http://localhost:8080
echo Backend:  http://localhost:5001
echo.
echo Press Ctrl+C to stop the server
echo.

REM Start the development server
npm run dev
