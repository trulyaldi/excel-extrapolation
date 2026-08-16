@echo off
setlocal EnableExtensions

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
REM 2. Find Python 3.11 or newer
REM ------------------------------------------------------------

echo [2/5] Checking Python...
echo.

set "PYTHON_CMD="

REM First try the Windows Python Launcher.
py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3,11) else 1)" >nul 2>&1

if not errorlevel 1 (
    set "PYTHON_CMD=py -3"
    goto :python_found
)

REM Fall back to python on PATH.
python -c "import sys; raise SystemExit(0 if sys.version_info >= (3,11) else 1)" >nul 2>&1

if not errorlevel 1 (
    set "PYTHON_CMD=python"
    goto :python_found
)

echo ERROR: Python 3.11 or newer was not found.
echo.
echo Please install Python 3.11 or newer and run setup.bat again.
echo.
echo During Python installation, it is recommended to enable:
echo     "Add Python to PATH"
echo.
goto :error


:python_found

echo Compatible Python installation found:
%PYTHON_CMD% --version

echo.

REM ------------------------------------------------------------
REM 3. Create or validate virtual environment
REM ------------------------------------------------------------

echo [3/5] Setting up Python virtual environment...
echo.

if exist ".venv\Scripts\python.exe" (

    REM Make sure the existing environment is using Python 3.11+.
    ".venv\Scripts\python.exe" -c "import sys; raise SystemExit(0 if sys.version_info >= (3,11) else 1)" >nul 2>&1

    if errorlevel 1 (
        echo Existing .venv uses an unsupported Python version.
        echo Recreating .venv...
        echo.

        rmdir /s /q ".venv"

        %PYTHON_CMD% -m venv .venv

        if errorlevel 1 (
            echo.
            echo ERROR: Could not recreate the Python virtual environment.
            goto :error
        )
    ) else (
        echo Existing .venv found. Reusing it.
    )

) else (

    echo Creating .venv...

    %PYTHON_CMD% -m venv .venv

    if errorlevel 1 (
        echo.
        echo ERROR: Could not create the Python virtual environment.
        goto :error
    )
)

echo.
echo Virtual environment Python:
".venv\Scripts\python.exe" --version
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

echo.

".venv\Scripts\python.exe" -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ERROR: Could not install Python dependencies.
    echo Check your internet connection and requirements.txt.
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
    echo Make sure ALL Excel windows are closed and run setup.bat again.
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
echo Python interpreter:
echo.
echo     %CD%\.venv\Scripts\python.exe
echo.
echo ============================================================
echo                 EXCEL CONFIGURATION
echo ============================================================
echo.
echo Complete the following steps in Excel:
echo.
echo 1. Open Extraplus.xlsm.
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
echo 6. Go to:
echo.
echo     File
echo     ^> Options
echo     ^> Trust Center
echo     ^> Trust Center Settings
echo     ^> Macro Settings
echo.
echo    Enable:
echo.
echo     Trust access to the VBA project object model
echo.
echo 7. Press:
echo.
echo     Alt + F11
echo.
echo    Then go to:
echo.
echo     Tools ^> References
echo.
echo    Tick:
echo.
echo     xlwings
echo.
echo    Then click OK.
echo.
echo 8. Return to Excel and open the xlwings ribbon.
echo.
echo 9. Click:
echo.
echo     Import Python UDFs
echo.
echo 10. If necessary, click:
echo.
echo     Restart UDF Server
echo.
echo ============================================================
echo                     TEST FUNCTIONS
echo ============================================================
echo.
echo You should now be able to use:
echo.
echo     =EXTRAP1(A2:A20,B2:B20)
echo     =UNCERTAINTY1(A2:A20,B2:B20)
echo     =DECAY_RATE1(A2:A20,B2:B20)
echo.
echo Other available model families:
echo.
echo     EXTRAP1 / EXTRAP2 / EXTRAP3
echo     UNCERTAINTY1 / UNCERTAINTY2 / UNCERTAINTY3
echo     DECAY_RATE1 / DECAY_RATE2 / DECAY_RATE3
echo.
echo ============================================================
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