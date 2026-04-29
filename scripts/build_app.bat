@echo off
setlocal

set "ROOT_DIR=%~dp0.."
set "VENV_DIR=%ROOT_DIR%\venv"
set "PACKAGE_DIR=%ROOT_DIR%\build\package"
set "PYINSTALLER_BUILD_DIR=%ROOT_DIR%\build\main"
set "PYINSTALLER_ALT_BUILD_DIR=%ROOT_DIR%\build\WordCompare"

if not defined PYTHON_EXE set "PYTHON_EXE=python"

cd /d "%ROOT_DIR%"

echo [build_app] Cleaning previous outputs
if exist "%ROOT_DIR%\dist" rmdir /s /q "%ROOT_DIR%\dist"
if exist "%VENV_DIR%" rmdir /s /q "%VENV_DIR%"
if exist "%PACKAGE_DIR%" rmdir /s /q "%PACKAGE_DIR%"
if exist "%PYINSTALLER_BUILD_DIR%" rmdir /s /q "%PYINSTALLER_BUILD_DIR%"
if exist "%PYINSTALLER_ALT_BUILD_DIR%" rmdir /s /q "%PYINSTALLER_ALT_BUILD_DIR%"
mkdir "%PACKAGE_DIR%"

echo [build_app] Creating virtual environment with %PYTHON_EXE%
call %PYTHON_EXE% -m venv "%VENV_DIR%"
if errorlevel 1 exit /b %errorlevel%

"%VENV_DIR%\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 exit /b %errorlevel%

"%VENV_DIR%\Scripts\pip.exe" install -r requirements.txt
if errorlevel 1 exit /b %errorlevel%

"%VENV_DIR%\Scripts\pip.exe" install pyinstaller
if errorlevel 1 exit /b %errorlevel%

echo [build_app] Building native Rust extractor
call "%ROOT_DIR%\scripts\build_native_extractor.bat"
if errorlevel 1 exit /b %errorlevel%

echo [build_app] Running test suite
"%VENV_DIR%\Scripts\python.exe" -m unittest discover -s tests -v
if errorlevel 1 exit /b %errorlevel%

echo [build_app] Building PyInstaller package
"%VENV_DIR%\Scripts\pyinstaller.exe" main.spec
if errorlevel 1 exit /b %errorlevel%

for /f %%i in ('"%VENV_DIR%\Scripts\python.exe" -c "from version import __version__; print(__version__)"') do set "VERSION=%%i"
set "PACKAGE_NAME=WordCompare-Windows-%VERSION%"
set "PACKAGE_ROOT=%PACKAGE_DIR%\%PACKAGE_NAME%"
set "PACKAGE_ARCHIVE=%PACKAGE_DIR%\%PACKAGE_NAME%.zip"

if exist "%PACKAGE_ROOT%" rmdir /s /q "%PACKAGE_ROOT%"
if exist "%PACKAGE_ARCHIVE%" del /f /q "%PACKAGE_ARCHIVE%"
mkdir "%PACKAGE_ROOT%"

copy /Y "%ROOT_DIR%\dist\WordCompare.exe" "%PACKAGE_ROOT%\WordCompare.exe" >nul
if errorlevel 1 exit /b %errorlevel%
copy /Y "%ROOT_DIR%\README.md" "%PACKAGE_ROOT%\README.md" >nul
if errorlevel 1 exit /b %errorlevel%
copy /Y "%ROOT_DIR%\LICENSE" "%PACKAGE_ROOT%\LICENSE" >nul
if errorlevel 1 exit /b %errorlevel%

echo [build_app] Creating zip package
powershell -NoProfile -Command "Compress-Archive -Path '%PACKAGE_ROOT%' -DestinationPath '%PACKAGE_ARCHIVE%' -Force"
if errorlevel 1 exit /b %errorlevel%

echo Packaged artifact: %PACKAGE_ARCHIVE%
