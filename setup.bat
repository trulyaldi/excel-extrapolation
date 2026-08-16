@echo off
setlocal EnableExtensions

REM ============================================================
REM Extraplus Excel Setup
REM Supported Python versions: 3.11 - 3.14
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
    echo Make sure requirements.txt is in the root Extraplus folder.
    goto :error
)

echo Project files found.
echo.

REM ------------------------------------------------------------
REM 2. Find a supported Python installation
REM ------------------------------------------------------------

echo [2/5] Checking Python...
echo.

set "PYTHON_CMD="

REM Prefer the newest explicitly supported Python version.

py -3.14 --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py -3.14"
    goto :python_found
)

py -3.13 --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py -3.13"
    goto :python_found
)

py -3.12 --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py -3.12"
    goto :python_found
)

py -3.11 --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py -3.11"
    goto :python_found
)

REM Fall back to "python" if the Windows Python Launcher is unavailable.

python -c "import sys; raise SystemExit(0 if (3,11) <= sys.version_info[:2] <= (3,14) else 1)" >nul 2>&1

if not errorlevel 1 (
    set "PYTHON_CMD=python"
    goto :python_found
)

echo ERROR: A supported Python version was not found.
echo.
echo Extraplus currently supports:
echo.
echo     Python 3.11
echo     Python 3.12
echo     Python 3.13
echo     Python 3.14
echo.
echo Please install Python 3.11 or newer, up to Python 3.14,
echo and then run setup.bat again.
echo.
echo During Python installation, it is recommended to enable:
echo.
echo     Add Python to PATH
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

    REM Verify that the existing virtual environment uses
    REM a supported Python version.

    ".venv\Scripts\python.exe" -c "import sys; raise SystemExit(0 if (3,11) <= sys.version_info[:2] <= (3,14) else 1)" >nul 2>&1

    if errorlevel 1 (
        echo Existing .venv uses an unsupported Python version.
        echo Recreating .venv...
        echo.

        rmdir /s /q ".venv"

        if exist ".venv" (
            echo.
            echo ERROR: Could not remove the existing .venv folder.
            echo.
            echo Close Excel, Command Prompt, VS Code, Python,
            echo or any other program that may be using the environment.
            echo Then run setup.bat again.
            goto :error
        )

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

if not exist ".venv\Scripts\python.exe" (
    echo.
    echo ERROR: The virtual environment was not created correctly.
    goto :error
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
    echo.
    echo Check:
    echo     - Your internet connection
    echo     - requirements.txt
    echo     - Your Python installation
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

REM Warn if Excel is currently running.

tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul

if not errorlevel 1 (
    echo.
    echo ERROR: Microsoft Excel is currently running.
    echo.
    echo Close ALL Excel windows before installing the xlwings add-in.
    echo Then run setup.bat again.
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
echo 1. Open your .xlsm workbook.
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
echo 10. Click:
echo.
echo     Restart UDF Server
echo.
echo 11. Save the workbook as an .xlsm file.
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