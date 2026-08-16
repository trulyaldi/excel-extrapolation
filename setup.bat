@echo off
setlocal

REM ============================================================
REM Extraplus Excel Setup
REM ============================================================

cd /d "%~dp0"

echo.
echo ============================================================
echo                   EXTRAPLUS SETUP
echo ============================================================
echo.

REM ------------------------------------------------------------
REM 1. Check required project files
REM ------------------------------------------------------------

echo [1/5] Checking project files...

if not exist "extrap.py" (
    echo.
    echo ERROR: extrap.py was not found.
    echo Make sure setup.bat is in the root Extraplus folder.
    goto :error
)

if not exist "extrapolation.py" (
    echo.
    echo ERROR: extrapolation.py was not found.
    echo Make sure setup.bat is in the root Extraplus folder.
    goto :error
)

if not exist "requirements.txt" (
    echo.
    echo ERROR: requirements.txt was not found.
    goto :error
)

echo Project files found.
echo.

REM ------------------------------------------------------------
REM 2. Check Python 3.11
REM ------------------------------------------------------------

echo [2/5] Checking Python 3.11...

py -3.11 --version >nul 2>&1

if errorlevel 1 (
    echo.
    echo ERROR: Python 3.11 was not found.
    echo.
    echo Please install Python 3.11 and then run setup.bat again.
    echo.
    echo During Python installation, enable:
    echo     "Add Python to PATH"
    echo.
    goto :error
)

py -3.11 --version
echo.

REM ------------------------------------------------------------
REM 3. Create virtual environment
REM ------------------------------------------------------------

echo [3/5] Setting up Python virtual environment...

if exist ".venv\Scripts\python.exe" (
    echo Existing .venv found. Reusing it.
) else (
    echo Creating .venv...
    py -3.11 -m venv .venv

    if errorlevel 1 (
        echo.
        echo ERROR: Could not create the Python virtual environment.
        goto :error
    )
)

echo.

REM ------------------------------------------------------------
REM 4. Install Python dependencies
REM ------------------------------------------------------------

echo [4/5] Installing Python dependencies...
echo.

".venv\Scripts\python.exe" -m pip install --upgrade pip

if errorlevel 1 (
    echo.
    echo ERROR: Could not upgrade pip.
    goto :error
)

".venv\Scripts\python.exe" -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ERROR: Could not install Python dependencies.
    goto :error
)

echo.
echo Python dependencies installed successfully.
echo.

REM ------------------------------------------------------------
REM 5. Install xlwings Excel add-in
REM ------------------------------------------------------------

echo [5/5] Installing xlwings Excel add-in...
echo.

if not exist ".venv\Scripts\xlwings.exe" (
    echo ERROR: xlwings was not installed.
    echo Make sure xlwings is included in requirements.txt.
    goto :error
)

".venv\Scripts\xlwings.exe" addin install

if errorlevel 1 (
    echo.
    echo ERROR: xlwings add-in installation failed.
    echo.
    echo Make sure all Excel windows are closed and run setup.bat again.
    goto :error
)

REM ------------------------------------------------------------
REM Finished
REM ------------------------------------------------------------

echo.
echo ============================================================
echo              EXTRAPLUS INSTALLATION COMPLETE
echo ============================================================
echo.
echo Python environment:
echo.
echo     %CD%\.venv
echo.
echo ------------------------------------------------------------
echo Excel configuration
echo ------------------------------------------------------------
echo.
echo 1. Open Excel.
echo.
echo 2. Open the xlwings ribbon.
echo.
echo 3. Set Interpreter to:
echo.
echo     %CD%\.venv\Scripts\python.exe
echo.
echo 4. Set UDF Modules to:
echo.
echo     extrapolation
echo.
echo 5. Enable:
echo.
echo     Add Workbook to PYTHONPATH
echo.
echo 6. Click:
echo.
echo     Import Python UDFs
echo.
echo 7. If necessary, open:
echo.
echo     Alt+F11 ^> Tools ^> References
echo.
echo    and make sure "xlwings" is checked.
echo.
echo ------------------------------------------------------------
echo You should then be able to use functions such as:
echo.
echo     =EXTRAP1(A2:A20,B2:B20)
echo     =UNCERTY1(A2:A20,B2:B20)
echo     =PARAMETER_B1(A2:A20,B2:B20)
echo ------------------------------------------------------------
echo.
pause
exit /b 0


:error
echo.
echo ============================================================
echo                    SETUP FAILED
echo ============================================================
echo.
echo Please fix the error shown above and run setup.bat again.
echo.
pause
exit /b 1