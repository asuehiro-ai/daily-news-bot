@echo off
cd /d "%~dp0"

echo ============================================
echo  Kessan Tool - 3-Year Comparison + Valuation
echo ============================================
echo.
echo  Industry options  (set INDUSTRY= below):
echo    seizo      : Manufacturing
echo    it         : IT / Software
echo    service    : Services
echo    kensetsu   : Construction
echo    shokuhin   : Food
echo    inshoku    : Restaurant
echo    iryo       : Medical
echo    unyu       : Transport
echo    kouri      : Retail
echo    fudosan    : Real Estate
echo ============================================
echo.

rem *** EDIT INDUSTRY BELOW ***
set INDUSTRY=seizo
rem ***************************

rem -- Find Python (3.9 - 3.13) --
set PYEXE=
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set PYEXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set PYEXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set PYEXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe
if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" set PYEXE=%LOCALAPPDATA%\Programs\Python\Python310\python.exe
if exist "%LOCALAPPDATA%\Programs\Python\Python39\python.exe"  set PYEXE=%LOCALAPPDATA%\Programs\Python\Python39\python.exe
if exist "C:\Python312\python.exe" set PYEXE=C:\Python312\python.exe

if "%PYEXE%"=="" (
    echo [ERROR] Python not found.
    echo Install Python 3.x: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Industry : %INDUSTRY%
echo Python   : %PYEXE%
echo.

"%PYEXE%" "%~dp0kessan_tool.py" %INDUSTRY%

echo.
if %ERRORLEVEL% == 0 (
    echo [OK] Excel created: kessan_output.xlsx
    start "" "%~dp0kessan_output.xlsx"
) else (
    echo [ERROR] See messages above.
    echo   - Check that PDF files exist in this folder
    echo   - Scanned image PDFs cannot be read automatically
    echo   - pip install pdfplumber openpyxl yfinance pandas
)
pause
