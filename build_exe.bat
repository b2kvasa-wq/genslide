@echo off
title GenSlide EXE Builder
color 0b
echo.
echo ===================================================================
echo   BUILD: GenSlide PyInstaller EXE
echo ===================================================================
echo.

REM Clean previous build artifacts
echo [1/4] Cleaning old build artifacts...
if exist "build" rd /s /q "build"
if exist "dist" rd /s /q "dist"
echo       Done.
echo.

REM Run PyInstaller with the spec file
echo [2/4] Running PyInstaller...
echo ---------------------------------------------------------------
pyinstaller GenSlide.spec --clean
echo ---------------------------------------------------------------
echo.

REM Check if build succeeded
if not exist "dist\GenSlide.exe" (
    echo.
    echo  [FAILED] GenSlide.exe was NOT created!
    echo  Check the errors above.
    echo.
    pause
    exit /b 1
)

REM The Vosk speech model is 100% embedded directly inside GenSlide.exe
echo [3/3] Build complete! Single-file standalone executable ready.
echo.
echo.
echo ===================================================================
echo   SUCCESS: dist\GenSlide.exe
echo   Size: 
for %%A in ("dist\GenSlide.exe") do echo          %%~zA bytes
echo.
echo   To run: double-click dist\GenSlide.exe
echo   Or drag a .pptx file onto GenSlide.exe to auto-load it.
echo ===================================================================
echo.
pause
