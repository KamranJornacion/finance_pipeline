@echo off
REM Windows: double-click this file. Sets up Python the first time, then builds the report.
cd /d "%~dp0"
set PY=
where py >nul 2>nul && set PY=py -3
if not defined PY where python >nul 2>nul && set PY=python
if not defined PY (
  echo Python 3.10 or newer is needed. Install it from https://www.python.org/downloads/
  echo During install, tick "Add python.exe to PATH". Then try again.
  pause & exit /b 1
)
if not exist .venv\Scripts\python.exe (
  echo First-time setup: installing what is needed ^(this takes a minute^)...
  %PY% -m venv .venv || (echo Could not create the Python environment. & pause & exit /b 1)
)
fc /b requirements.txt .venv\installed-requirements.txt >nul 2>nul
if errorlevel 1 (
  .venv\Scripts\python.exe -m pip install --quiet --disable-pip-version-check -r requirements.txt || (echo Could not install the required packages - check your internet connection. & pause & exit /b 1)
  copy /y requirements.txt .venv\installed-requirements.txt >nul
)
.venv\Scripts\python.exe -m finpipe %*
if errorlevel 1 (
  echo.
  echo Finished with problems - please read the messages above.
  pause
  exit /b 1
)
timeout /t 3 >nul
