@echo off
echo ==========================================
echo  Recomp-to-Playnite GUI - Build
echo ==========================================
python -m PyInstaller --onefile --windowed recomps_gui.py
if %errorlevel% neq 0 (
    echo BUILD FEHLGESCHLAGEN!
    pause
    exit /b 1
)
echo ==========================================
echo  BUILD ERFOLGREICH!
echo  EXE liegt unter: dist\recomps_gui.exe
echo ==========================================
pause
