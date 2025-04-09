@echo off
setlocal

:begin
REM Check if Python is already installed
echo Checking for existing Python installation...
python --version >nul 2>&1
if %ERRORLEVEL%==0 (
    echo Python is already installed. Upgrading pip and setuptools to the latest version...
    python -m pip install --upgrade pip
    python -m pip install --upgrade setuptools
    goto :end
)

REM Define temporary folder and installer name
set TEMP_FOLDER=%TEMP%\python_install
set INSTALLER_NAME=python-3.13.2-amd64.exe

REM Create a temporary directory
if not exist %TEMP_FOLDER% mkdir %TEMP_FOLDER%
cd %TEMP_FOLDER%

REM Kill any existing Python installer process
tasklist /FI "IMAGENAME eq %INSTALLER_NAME%" 2>NUL | find /I "%INSTALLER_NAME%" >NUL
if %ERRORLEVEL%==0 (
    echo Terminating existing Python installer process...
    taskkill /F /IM %INSTALLER_NAME%
)

REM Download the specific version of Python from the official website
echo Downloading Python 3.13.2...
powershell -Command "Invoke-WebRequest -Uri https://www.python.org/ftp/python/3.13.2/python-3.13.2-amd64.exe -OutFile %INSTALLER_NAME%"

REM Check if download was successful
if not exist %INSTALLER_NAME% (
    echo Failed to download Python. Please check your internet connection.
    pause
    exit /b
)

REM Run the installer silently with recommended settings
echo Installing Python 3.13.2...
start /wait %INSTALLER_NAME% /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1

REM Clean up the temporary files
del %INSTALLER_NAME%
cd ..
rmdir /s /q %TEMP_FOLDER%

REM Verify installation
echo Verifying Python installation...
python --version
pip --version

:end
echo Python installation and setup completed.

REM Restart the batch file once
if "%REPEAT%"=="" (
    set REPEAT=1
    echo Restarting the batch file completely...
    start "" "%~f0"
    exit
)

echo Installation process completed after one restart.
pause