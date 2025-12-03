@echo off
setlocal

echo ---------------------------------------------
echo   Task Bar Saver - EXE Builder
echo ---------------------------------------------

REM === Auto-locate Python interpreter ===
for %%P in (python.exe python3.exe py.exe) do (
    where %%P >nul 2>&1
    if not errorlevel 1 (
        set PYTHON=%%P
        goto :foundpython
    )
)

echo ❌ ERROR: Python not found on this system.
pause
exit /b


:foundpython
echo ✔ Using Python interpreter: %PYTHON%
echo.

REM === Find the .py file in the current folder ===
for %%F in ("*.py") do (
    set SCRIPT=%%F
    goto :foundscript
)

echo ❌ ERROR: No .py file found in this folder.
pause
exit /b

:foundscript
echo ✔ Found Python script: %SCRIPT%
echo.

REM === Set Desktop output directory ===
set DESKTOP=%USERPROFILE%\Desktop

echo ✔ EXE will be built to: %DESKTOP%
echo.

REM === Run PyInstaller ===
echo 🛠 Building EXE...
%PYTHON% -m pip install --upgrade pyinstaller >nul
%PYTHON% -m PyInstaller --onefile --noconsole --distpath "%DESKTOP%" "%SCRIPT%"

echo.
echo ---------------------------------------------
echo   ✔ Build Complete!
echo   Your EXE is now on your Desktop.
echo ---------------------------------------------
pause
