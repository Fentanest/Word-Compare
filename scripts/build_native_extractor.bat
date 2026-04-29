@echo off
setlocal

set "ROOT_DIR=%~dp0.."
set "CRATE_DIR=%ROOT_DIR%\native\docx-structure-extractor"
set "TARGET_DIR=%CRATE_DIR%\target\release"
set "OUTPUT_DIR=%ROOT_DIR%\build\native"
set "OUTPUT_NAME=word_compare_native_extractor.exe"

if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

cargo build --manifest-path "%CRATE_DIR%\Cargo.toml" --release
if errorlevel 1 exit /b %errorlevel%

copy /Y "%TARGET_DIR%\%OUTPUT_NAME%" "%OUTPUT_DIR%\%OUTPUT_NAME%" >nul
if errorlevel 1 exit /b %errorlevel%

echo Native extractor built into %OUTPUT_DIR%
