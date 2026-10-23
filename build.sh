#!/bin/sh
echo "=========================================="
echo " Recomp-to-Playnite GUI - Build"
echo "=========================================="
python -m PyInstaller --onefile --windowed recomps_gui.py
if [ $? -ne 0 ]; then
    echo "BUILD FEHLGESCHLAGEN!"
    exit 1
fi
echo "=========================================="
echo " BUILD ERFOLGREICH!"
echo " EXE liegt unter: dist/recomps_gui"
echo "=========================================="
