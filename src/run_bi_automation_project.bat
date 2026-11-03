@echo off
setlocal

set "PROJECT_DIR=%~dp0.."
for %%I in ("%PROJECT_DIR%") do set "PROJECT_DIR=%%~fI"
cd /d "%PROJECT_DIR%"

set "PYTHONPATH=%PROJECT_DIR%;%PROJECT_DIR%\src;%PYTHONPATH%"

if exist "%PROJECT_DIR%\.venv\Scripts\python.exe" (
    "%PROJECT_DIR%\.venv\Scripts\python.exe" "%PROJECT_DIR%\run_project.py"
) else (
    python "%PROJECT_DIR%\run_project.py"
)

if errorlevel 1 (
    echo.
    echo BI Automation failed. Review the error messages above.
    exit /b 1
)

endlocal
