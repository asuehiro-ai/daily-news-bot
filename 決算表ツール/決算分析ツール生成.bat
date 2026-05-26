@echo off
cd /d "%~dp0"
chcp 65001 > nul

echo ============================================
echo  決算分析ツール生成スクリプト
echo  (業界平均比較 / 企業価値算定 修正前・後)
echo ============================================
echo.
echo  前提: 「3期比較表.xlsx」に決算数値を入力済みであること
echo.

rem -- Python 検索 --
set PYEXE=
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set PYEXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set PYEXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set PYEXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe
if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" set PYEXE=%LOCALAPPDATA%\Programs\Python\Python310\python.exe
if exist "%LOCALAPPDATA%\Programs\Python\Python39\python.exe"  set PYEXE=%LOCALAPPDATA%\Programs\Python\Python39\python.exe
if exist "C:\Python312\python.exe" set PYEXE=C:\Python312\python.exe
for /f "tokens=*" %%i in ('where python 2^>nul') do if "%PYEXE%"=="" set PYEXE=%%i

if "%PYEXE%"=="" (
    echo [ERROR] Python が見つかりません。
    echo   https://www.python.org/downloads/ からインストールしてください。
    pause
    exit /b 1
)

echo Python: %PYEXE%
echo.
echo ---- 生成開始 ----

"%PYEXE%" "%~dp0valuation_tool.py"

echo.
if %ERRORLEVEL% == 0 (
    echo [OK] 「決算分析ツール.xlsx」を開きます...
    start "" "%~dp0決算分析ツール.xlsx"
) else (
    echo [ERROR] エラーが発生しました。上記メッセージを確認してください。
    echo   pip install openpyxl
)
pause
