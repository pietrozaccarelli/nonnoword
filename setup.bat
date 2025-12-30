@echo off
setlocal enabledelayedexpansion

:: --- Configuration ---
set "SCRIPT_NAME=script.py"
set "VENV_NAME=env"
set "LAUNCHER_NAME=NonnoWord.bat"

:: Get absolute paths
set "BASE_DIR=%~dp0"
set "ABS_SCRIPT=%BASE_DIR%%SCRIPT_NAME%"
set "ABS_VENV=%BASE_DIR%%VENV_NAME%"
set "ABS_PYTHON_EXE=%ABS_VENV%\Scripts\pythonw.exe"

echo ====================================================
echo           NonnoWord Installer Setup
echo ====================================================

:: 1. Check if word.py exists
if not exist "%ABS_SCRIPT%" (
    echo [ERROR] %SCRIPT_NAME% not found in this folder.
    echo Please place setup.bat and %SCRIPT_NAME% in the same directory.
    pause
    exit /b
)

:: 2. Create Virtual Environment
echo [1/3] Creating Virtual Environment in: "%ABS_VENV%"...
python -m venv "%ABS_VENV%"
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Failed to create virtual environment. 
    echo Ensure Python is installed and added to your PATH.
    pause
    exit /b
)

:: 3. Install Dependencies
echo [2/3] Installing required libraries (python-docx)...
"%ABS_VENV%\Scripts\pip" install python-docx --quiet
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Failed to install dependencies. Check your internet connection.
    pause
    exit /b
)

:: 4. Create the Launcher .bat
echo [3/3] Creating Launcher: %LAUNCHER_NAME%...

echo @echo off > "%LAUNCHER_NAME%"
echo :: This file triggers the Python Word Emulator using its specific venv >> "%LAUNCHER_NAME%"
echo start "" "%ABS_PYTHON_EXE%" "%ABS_SCRIPT%" %%* >> "%LAUNCHER_NAME%"

echo.
echo ====================================================
echo Setup Successful!
echo ====================================================
echo 1. A virtual environment was created in the 'env' folder.
echo 2. 'python-docx' was installed.
echo 3. '%LAUNCHER_NAME%' was created.
echo.
echo You can now move '%LAUNCHER_NAME%' to your Desktop or
echo anywhere else. It is linked to this directory.
echo ====================================================
pause