@echo off
cd /d "%~dp0"

python -c "import tkinter" >nul 2>&1
if not errorlevel 1 goto RunApp

echo Python 3.13 with Tkinter was not found. Installing the bundled Python...
if not exist "%~dp0instalador\python-3.13.17-amd64.exe" (
    echo Bundled Python installer not found.
    echo Please install Python 3.13 manually and run this file again.
    pause
    exit /b 1
)

"%~dp0instalador\python-3.13.17-amd64.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_tcltk=1 SimpleInstall=1
if errorlevel 1 (
    echo Python installation failed or was cancelled.
    pause
    exit /b 1
)

set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python313\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=python"

:RunApp
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0recursos\criar_atalho.ps1" >nul 2>&1
where pythonw >nul 2>&1
if not errorlevel 1 (
    start "" pythonw app\app.py
) else (
    python app\app.py
    if errorlevel 1 pause
)
