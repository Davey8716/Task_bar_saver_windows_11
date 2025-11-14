# Task Bar Saver
A lightweight, safe Windows utility to:

- ✔ Backup your pinned taskbar shortcuts  
- ✔ Skip duplicates (no repeated backups)  
- ✔ Open the backup folder instantly  
- ✔ Take desktop screenshots (UI auto-hides itself)  
- ✔ Choose your own backup folder  
- ✔ Use a clean modern ttkbootstrap UI  
- ✔ Build your own EXE easily with the included tool  

Task Bar Saver **does NOT modify the registry**,  
does NOT restart Explorer,  
and never edits your real taskbar pins.  
It’s a **safe manual backup utility**.

---

# Project Setup
Before using or building the EXE, create a folder on your computer:

```
Task-Bar-Saver/
│
├── Task Bar Saver Final.py
├── build_exe.txt
└── README.md
```

All files **must be in the same folder**.

---

# Running the Python Program Directly

## 1. Install dependencies
```
pip install ttkbootstrap pillow
```

## 2. Run the program
```
python "Task Bar Saver Final.py"
```

---

# Building the EXE (Beginner Friendly)

This project includes a universal EXE builder:  
`build_exe.txt` (renamed by Windows for safety).

To build the EXE:

---

## Step 1 — Rename the EXE Builder
Rename:

```
build_exe.txt → build_exe.bat
```

IMPORTANT:  
When renaming, use **“All Files (*.*)”** so Windows does NOT make it `build_exe.bat.txt`.

---

## Step 2 — Place the BAT file in the same folder  
Your folder must look like:

```
Task-Bar-Saver/
│
├── Task Bar Saver Final.py
├── build_exe.bat
└── README.md
```

---

## Step 3 — Run the EXE Builder
Double‑click:

```
build_exe.bat
```

The builder will:

- ✔ Auto-detect your Python  
- ✔ Auto-find the `.py` file  
- ✔ Install PyInstaller automatically  
- ✔ Build a standalone EXE  
- ✔ Place the EXE on your **real Windows Desktop** (not OneDrive)

---

# EXE Output Location
After building, your EXE will appear on:

```
C:\Users\<your-name>\Desktop\
```

---

# Included EXE Builder (for reference)

```
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
```

---

# Notes
- Task Bar Saver does **not** write to your Windows taskbar.  
- Backup folder contents are safe to delete, move, or edit.  
- Screenshot tool hides the UI before capturing.  

---

# Done
You're ready to use or build your own Task Bar Saver EXE.
