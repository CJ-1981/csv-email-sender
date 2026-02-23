@echo off
echo ============================================
echo CSV Email Sender - Local Development Server
echo ============================================
echo.
echo Starting local web server on http://localhost:8000
echo.
echo Press Ctrl+C to stop the server
echo.
echo ============================================
echo.

REM Try Python 3 first
python --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo Using Python 3 to start server...
    python -m http.server 8000
    goto :end
)

REM Try Python 2
python2 --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo Using Python 2 to start server...
    python -m SimpleHTTPServer 8000
    goto :end
)

REM Try Python command (might be Python 2 or 3)
py --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo Using Python to start server...
    py -m http.server 8000
    goto :end
)

REM If no Python found
echo ============================================
echo ERROR: Python not found!
echo ============================================
echo.
echo Please install Python or use another method:
echo.
echo Option 1: Install Python from https://python.org
echo.
echo Option 2: Use Node.js:
echo   npx serve
echo.
echo Option 3: Use PHP:
echo   php -S localhost:8000
echo.
echo ============================================
echo.
pause
goto :end

:end
