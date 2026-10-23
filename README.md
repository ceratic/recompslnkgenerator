# Recomp-to-Playnite GUI

Ein kleines Windows/Linux-Tool mit grafischer Oberfläche, um **Quiver**- und **DecompBuddy**-Recompilationen automatisch zu scannen und Playnite-kompatible Verknüpfungen (`.lnk` / `.desktop`) zu erstellen.

---

## Features

- Scannt **Quiver** (`quiver-launch.log`) und **DecompBuddy** (`decomp-buddy.json`)
- Erstellt `portable.txt` in den Spielordnern
- Generiert Verknüpfungen: `.lnk` (Windows) / `.desktop` (Linux)
- **Cleanup**: Entfernt veraltete Verknüpfungen, deren Ziel-EXE nicht mehr existiert
- Erstellt `scan_bericht.txt` mit Zusammenfassung
- Einstellbare Basis-Pfade (werden in `recomps_config.json` gespeichert)
- Fortschrittsbalken + Live-Log während des Scans
- Dialog zur manuellen EXE-Auswahl, falls eine Executable fehlt
- One-File-Exe via PyInstaller

---

## Schnellstart (ohne Build)

```bash
python recomps_gui.py
```

---

## Build der `.exe` (Windows)

```bat
build.bat
```

Ergebnis: `dist/recomps_gui.exe`

---

## Build auf Linux

```bash
./build.sh
```

---

## Anforderungen

- Python 3.8+
- Windows: `pywin32` (wird automatisch für Verknüpfungen genutzt)
- Linux: Standard-Python + Tkinter (meist vorinstalliert)

---

## Lizenz

MIT – frei nutzbar, anpassbar, weiterverteilbar.