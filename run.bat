@echo off
:: Ziara TalentX - One-Click Launcher
:: Developed by Ziara TechQ Labs (Md Aquil)
title Ziara TalentX - Automated CV Contact Extractor
cls

echo ======================================================================
echo     Ziara TalentX - Automated Candidate Contact Extractor
echo                  Ziara TechQ Labs (Md Aquil)
echo ======================================================================
echo.

REM 1. Detect Python
set "PY_CMD="

python --version >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=python"
    goto :python_found
)

py --version >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=py"
    goto :python_found
)

REM Search in AppData
for /d %%D in ("%LocalAppData%\Programs\Python\Python3*") do (
    if exist "%%D\python.exe" (
        set "PY_CMD=%%D\python.exe"
        goto :python_found
    )
)

REM Search in Program Files
for /d %%D in ("C:\Program Files\Python3*") do (
    if exist "%%D\python.exe" (
        set "PY_CMD=%%D\python.exe"
        goto :python_found
    )
)

:python_not_found
echo [ERROR] Python is NOT detected on this computer!
echo.
echo ======================================================================
echo  HOW TO FIX THIS (Simple 2-Step Setup):
echo ======================================================================
echo  1. Download Python from: https://www.python.org/downloads/
echo  2. IMPORTANT: During setup, make sure to check:
echo        [X] Add python.exe to PATH
echo  3. Click "Install Now".
echo  4. Once installed, double-click this run.bat file again!
echo ======================================================================
echo.
pause
exit /b 1

:python_found
echo [OK] Python detected!
%PY_CMD% --version
echo.

REM 2. Fast check if required packages are already installed
%PY_CMD% -c "import fitz, pandas, openpyxl, phonenumbers" >nul 2>&1
if not errorlevel 1 (
    echo [OK] All required packages are already installed!
    goto :packages_ready
)

echo [Notice] First-time setup: Downloading required packages...
echo (Please wait, downloading libraries from the internet - happens only once)
echo.
%PY_CMD% -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [WARNING] Some optional packages failed to install. Proceeding anyway...
) else (
    echo.
    echo [OK] All required packages installed successfully!
)

:packages_ready
echo.

REM 3. Candidate folder selection
set "TARGET_DIR=%~1"

if "%TARGET_DIR%"=="" (
    echo Tip: You can drag and drop your candidate folder directly onto this run.bat file!
    echo.
    set /p "TARGET_DIR=Enter folder name/path [or press ENTER for auto-detect]: "
)

echo.
echo [2/2] Running CV and Contact extraction...
echo ----------------------------------------------------------------------
echo.

if "%TARGET_DIR%"=="" (
    %PY_CMD% extract_resumes.py
) else (
    %PY_CMD% extract_resumes.py --folder "%TARGET_DIR%"
)

echo.
echo ======================================================================
echo  Process completed!
echo  Check generated spreadsheet files:
echo    - candidates_resume_data.xlsx  (Formatted Excel sheet)
echo    - candidates_resume_data.csv   (Raw CSV data)
echo ======================================================================
echo.
pause
