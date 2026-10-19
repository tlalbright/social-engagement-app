@echo off
cd /d "%~dp0"
echo Starting Social Engagement Processor...
echo.

REM Find Python — try 'py' launcher first (standard Windows install), then 'python'
where py >nul 2>&1
if %errorlevel%==0 (
    set PYTHON=py
    goto :run
)
where python >nul 2>&1
if %errorlevel%==0 (
    set PYTHON=python
    goto :run
)
echo ERROR: Python not found. Please install Python 3 from https://www.python.org/downloads/
echo Make sure to check "Add Python to PATH" during installation.
pause
exit /b 1

:run
echo Installing / verifying dependencies...
%PYTHON% -m pip install -r requirements.txt -q
echo.
%PYTHON% -m streamlit run app.py
pause
