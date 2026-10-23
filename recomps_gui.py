import os
import sys
import json
import platform
import queue
import threading
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ==========================================
# KONFIGURATION
# ==========================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, "frozen", False):
    # Bei einer PyInstaller-/One-File-Exe läuft Python aus dem temporären Entpack-Ordner
    CONFIG_PATH = os.path.join(sys._MEIPASS, "recomps_config.json")
else:
    CONFIG_PATH = os.path.join(SCRIPT_DIR, "recomps_config.json")

DEFAULT_CONFIG = {
    "quiver_base_dir": r"",
    "decompbuddy_base_dir": r"",
    "output_base_dir": r"",
}

STATUS_ORDER = [
    "NEU HINZUGEFÜGT",
    "ENTFERNT (Spiel gelöscht)",
    "DUPLIKAT (INTERN)",
    "BEREITS VORHANDEN",
    "FEHLER",
]

DONE_MARKER = "___SCAN_DONE___"
MISSING_EXE_MARKER = "__MISSING_EXE__"


# ==========================================
# KONFIGURATION LESEN / SCHREIBEN
# ==========================================
def read_config():
    cfg = dict(DEFAULT_CONFIG)
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg.update(json.load(f))
    except Exception:
        pass
    return cfg


def write_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


# ==========================================
# SCAN BERICHT KLASSE
# ==========================================
class ScanReport:
    def __init__(self, output_dir, log=None):
        self.file_path = os.path.join(output_dir, "scan_bericht.txt")
        self.entries = []
        self.log = log or (lambda msg: print(msg))

    def add(self, status, tool, game_name, path, details=""):
        self.entries.append({
            "status": status,
            "tool": tool,
            "game_name": game_name,
            "path": path,
            "details": details,
        })

    def save(self):
        with open(self.file_path, "w", encoding="utf-8") as f:
            f.write(f"=== RECOMP TO PLAYNITE SCAN BERICHT ===\n")
            f.write(f"Erstellt am: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 50 + "\n\n")

            for status in STATUS_ORDER:
                matches = [e for e in self.entries if e["status"] == status]
                if matches:
                    f.write(f"\n--- {status} ({len(matches)}) ---\n")
                    for entry in matches:
                        f.write(f"🎮 {entry['game_name']}\n")
                        f.write(f"   Tool : {entry['tool']}\n")
                        f.write(f"   Pfad : {entry['path']}\n")
                        if entry["details"]:
                            f.write(f"   Info : {entry['details']}\n")
                        f.write("-" * 40 + "\n")

            f.write(f"\n=== ZUSAMMENFASSUNG ===\n")
            f.write(f"Neu hinzugefügt : {len([e for e in self.entries if e['status'] == 'NEU HINZUGEFÜGT'])}\n")
            f.write(f"Entfernt        : {len([e for e in self.entries if e['status'] == 'ENTFERNT (Spiel gelöscht)'])}\n")
            f.write(f"Übersprungen    : {len([e for e in self.entries if e['status'] in ['DUPLIKAT (INTERN)', 'BEREITS VORHANDEN']])}\n")
            f.write(f"Fehler          : {len([e for e in self.entries if e['status'] == 'FEHLER'])}\n")

        self.log(f"\n📄 Scan-Bericht gespeichert unter: {self.file_path}")


# ==========================================
# HILFSFUNKTIONEN
# ==========================================
def sanitize_filename(filename):
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '')
    return filename.strip()


def ensure_portable_txt(game_dir, report, tool_name, game_name, log=None):
    if log is None:
        log = lambda msg: print(msg)
    portable_path = os.path.join(game_dir, "portable.txt")
    if not os.path.exists(portable_path):
        try:
            with open(portable_path, 'w', encoding='utf-8') as f:
                pass
            log(f"  [+] portable.txt erstellt in: {game_dir}")
            report.add("NEU HINZUGEFÜGT", tool_name, game_name, portable_path, "portable.txt wurde automatisch erstellt")
        except Exception as e:
            log(f"  [!] Fehler beim Erstellen von portable.txt: {e}")
            report.add("FEHLER", tool_name, game_name, portable_path, str(e))


def create_shortcut_windows(name, target_exe, work_dir, output_dir, log=None):
    if log is None:
        log = lambda msg: print(msg)
    shortcut_path = os.path.join(output_dir, f"{sanitize_filename(name)}.lnk")

    if os.path.exists(shortcut_path):
        return False, "Verknüpfung existiert bereits im Ausgabeordner"

    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = target_exe
        shortcut.WorkingDirectory = work_dir
        shortcut.save()
        log(f"  [+] Verknüpfung erstellt: {name}")
        return True, "Erfolgreich erstellt"
    except ImportError:
        reason = "Fehler: 'pywin32' ist nicht installiert. Führe 'pip install pywin32' aus."
        log(f"  [!] Übersprungen: {name} ({reason})")
        return False, reason
    except Exception as e:
        reason = str(e)
        log(f"  [!] Übersprungen: {name} ({reason})")
        return False, reason


def create_shortcut_linux(name, target_exe, output_dir, log=None):
    if log is None:
        log = lambda msg: print(msg)
    shortcut_path = os.path.join(output_dir, f"{sanitize_filename(name)}.desktop")

    if os.path.exists(shortcut_path):
        return False, "Desktop-Datei existiert bereits im Ausgabeordner"

    desktop_content = f"""[Desktop Entry]
Name={name}
Exec="{target_exe}"
Type=Application
Terminal=false
Icon=application-x-executable
"""
    try:
        with open(shortcut_path, 'w', encoding='utf-8') as f:
            f.write(desktop_content)
        os.chmod(shortcut_path, 0o755)
        log(f"  [+] Verknüpfung erstellt: {name}")
        return True, "Erfolgreich erstellt"
    except Exception as e:
        reason = str(e)
        log(f"  [!] Übersprungen: {name} ({reason})")
        return False, reason


# ==========================================
# CLEANUP FUNKTIONEN
# ==========================================
def cleanup_invalid_shortcuts_windows(tool_output_dir, tool_name, report, log=None):
    if log is None:
        log = lambda msg: print(msg)
    log(f"\n--- Cleanup (Windows): Prüfe Verknüpfungen in {tool_output_dir} ---")
    removed = 0

    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
    except ImportError:
        log("  [!] pywin32 nicht installiert. Cleanup übersprungen.")
        return

    for file in os.listdir(tool_output_dir):
        if file.lower().endswith(".lnk"):
            shortcut_path = os.path.join(tool_output_dir, file)
            try:
                shortcut = shell.CreateShortCut(shortcut_path)
                target_path = shortcut.Targetpath

                if not target_path or not os.path.exists(target_path):
                    game_name = os.path.splitext(file)[0]
                    os.remove(shortcut_path)
                    log(f"  [-] Entfernt (Ziel fehlt): {game_name}")
                    report.add("ENTFERNT (Spiel gelöscht)", tool_name, game_name, shortcut_path, f"Ziel-Executable existiert nicht mehr: {target_path}")
                    removed += 1
            except Exception as e:
                log(f"  [!] Fehler beim Prüfen von {file}: {e}")
                report.add("FEHLER", tool_name, file, shortcut_path, f"Cleanup-Fehler: {e}")

    log(f"  ✅ {removed} ungültige Verknüpfung(en) entfernt.")


def cleanup_invalid_shortcuts_linux(tool_output_dir, tool_name, report, log=None):
    if log is None:
        log = lambda msg: print(msg)
    log(f"\n--- Cleanup (Linux): Prüfe Verknüpfungen in {tool_output_dir} ---")
    removed = 0

    for file in os.listdir(tool_output_dir):
        if file.lower().endswith(".desktop"):
            shortcut_path = os.path.join(tool_output_dir, file)
            try:
                target_path = None
                with open(shortcut_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        if line.startswith("Exec="):
                            target_path = line[5:].strip().strip('"').strip("'")
                            target_path = target_path.split()[0] if ' ' in target_path else target_path
                            break

                if not target_path or not os.path.exists(target_path):
                    game_name = os.path.splitext(file)[0]
                    os.remove(shortcut_path)
                    log(f"  [-] Entfernt (Ziel fehlt): {game_name}")
                    report.add("ENTFERNT (Spiel gelöscht)", tool_name, game_name, shortcut_path, f"Ziel-Executable existiert nicht mehr: {target_path}")
                    removed += 1
            except Exception as e:
                log(f"  [!] Fehler beim Prüfen von {file}: {e}")
                report.add("FEHLER", tool_name, file, shortcut_path, f"Cleanup-Fehler: {e}")

    log(f"  ✅ {removed} ungültige Verknüpfung(en) entfernt.")


def cleanup_invalid_shortcuts(tool_output_dir, tool_name, report, log=None):
    if platform.system() == "Windows":
        cleanup_invalid_shortcuts_windows(tool_output_dir, tool_name, report, log)
    else:
        cleanup_invalid_shortcuts_linux(tool_output_dir, tool_name, report, log)


# ==========================================
# SCAN FUNKTIONEN
# ==========================================
def process_quiver(base_dir, tool_output_dir, report, log=None, progress=None):
    if log is None:
        log = lambda msg: print(msg)
    log(f"\n--- Durchsuche Quiver: {base_dir} ---")
    processed_executables = set()

    for root, dirs, files in os.walk(base_dir):
        if "quiver-launch.log" in files:
            if progress:
                progress()
            log_path = os.path.join(root, "quiver-launch.log")
            game_name = None
            exe_path = None

            try:
                with open(log_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith("game="):
                            game_name = line[5:].strip()
                        elif line.startswith("chosenExecutable="):
                            exe_path = line[17:].strip()
            except Exception as e:
                log(f"  [!] Fehler beim Lesen von {log_path}: {e}")
                report.add("FEHLER", "Quiver", "Unbekannt", log_path, str(e))
                continue

            if game_name and exe_path and exe_path.lower() != "(none)" and os.path.exists(exe_path):
                if exe_path in processed_executables:
                    log(f"  [=] Übersprungen (Duplikat): {game_name}")
                    report.add("DUPLIKAT (INTERN)", "Quiver", game_name, exe_path, "Bereits in diesem Durchlauf verarbeitet")
                    continue

                log(f"  Gefunden: {game_name}")
                game_dir = os.path.dirname(exe_path)
                ensure_portable_txt(game_dir, report, "Quiver", game_name, log)

                if platform.system() == "Windows":
                    success, reason = create_shortcut_windows(game_name, exe_path, game_dir, tool_output_dir, log)
                else:
                    success, reason = create_shortcut_linux(game_name, exe_path, tool_output_dir, log)

                if success:
                    log(f"  [+] Verknüpfung erstellt: {game_name}")
                    report.add("NEU HINZUGEFÜGT", "Quiver", game_name, exe_path, reason)
                else:
                    log(f"  [!] Übersprungen: {game_name} ({reason})")
                    report.add("BEREITS VORHANDEN", "Quiver", game_name, exe_path, reason)

                processed_executables.add(exe_path)
            else:
                msg = "Ungültige oder unvollständige Daten" if not (game_name and exe_path) else f"Executable existiert nicht: {exe_path}"
                log(f"  [!] {msg} in {log_path}")
                report.add("FEHLER", "Quiver", game_name or "Unbekannt", log_path, msg)


def process_decompbuddy(base_dir, tool_output_dir, report, log=None, progress=None, on_missing_exe=None):
    if log is None:
        log = lambda msg: print(msg)
    log(f"\n--- Durchsuche DecompBuddy: {base_dir} ---")
    processed_executables = set()

    for root, dirs, files in os.walk(base_dir):
        if "decomp-buddy.json" in files:
            if progress:
                progress()
            json_path = os.path.join(root, "decomp-buddy.json")

            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                plan = data.get("plan", {})
                game_name = plan.get("game_title") or data.get("folder", "Unbekanntes Spiel")
                executable = plan.get("build", {}).get("executable")

                if not executable:
                    log(f"  [!] Keine Executable in {json_path} gefunden.")
                    report.add("FEHLER", "DecompBuddy", game_name, json_path, "Keine Executable im JSON-Plan gefunden")
                    continue

                game_dir = os.path.dirname(json_path)
                target_exe = os.path.join(game_dir, executable)

                if not os.path.exists(target_exe):
                    alt_exe = os.path.join(game_dir, data.get("folder", ""), executable)
                    if os.path.exists(alt_exe):
                        target_exe = alt_exe

                exe_to_use = target_exe
                if not os.path.exists(exe_to_use) and on_missing_exe:
                    chosen = on_missing_exe(game_name, exe_to_use)
                    if chosen and os.path.exists(chosen):
                        exe_to_use = chosen
                        log(f"  [+] Manuelles EXE gewählt: {chosen}")

                if os.path.exists(exe_to_use):
                    if exe_to_use in processed_executables:
                        log(f"  [=] Übersprungen (Duplikat): {game_name}")
                        report.add("DUPLIKAT (INTERN)", "DecompBuddy", game_name, exe_to_use, "Bereits in diesem Durchlauf verarbeitet")
                        continue

                    log(f"  Gefunden: {game_name}")
                    game_dir = os.path.dirname(json_path)
                    ensure_portable_txt(game_dir, report, "DecompBuddy", game_name, log)

                    if platform.system() == "Windows":
                        success, reason = create_shortcut_windows(game_name, exe_to_use, game_dir, tool_output_dir, log)
                    else:
                        success, reason = create_shortcut_linux(game_name, exe_to_use, tool_output_dir, log)

                    if success:
                        log(f"  [+] Verknüpfung erstellt: {game_name}")
                        report.add("NEU HINZUGEFÜGT", "DecompBuddy", game_name, exe_to_use, reason)
                    else:
                        log(f"  [!] Übersprungen: {game_name} ({reason})")
                        report.add("BEREITS VORHANDEN", "DecompBuddy", game_name, exe_to_use, reason)

                    processed_executables.add(exe_to_use)
                else:
                    log(f"  [!] Executable nicht gefunden: {exe_to_use}")
                    report.add("FEHLER", "DecompBuddy", game_name, json_path, f"Ziel-Executable existiert nicht: {exe_to_use}")

            except json.JSONDecodeError:
                log(f"  [!] Ungültiges JSON in {json_path}")
                report.add("FEHLER", "DecompBuddy", "Unbekannt", json_path, "JSON-Dekodierungsfehler")
            except Exception as e:
                log(f"  [!] Fehler beim Verarbeiten von {json_path}: {e}")
                report.add("FEHLER", "DecompBuddy", "Unbekannt", json_path, str(e))


# ==========================================
# SCAN AUSFÜHREN
# ==========================================
def summarize(report):
    def count(status):
        return len([e for e in report.entries if e["status"] == status])

    return {
        "neu": len([e for e in report.entries if e["status"] == "NEU HINZUGEFÜGT"]),
        "entfernt": len([e for e in report.entries if e["status"] == "ENTFERNT (Spiel gelöscht)"]),
        "uebersprungen": len([e for e in report.entries if e["status"] in ["DUPLIKAT (INTERN)", "BEREITS VORHANDEN"]]),
        "fehler": len([e for e in report.entries if e["status"] == "FEHLER"]),
    }


def count_target_files(base_dir):
    count = 0
    for root, dirs, files in os.walk(base_dir):
        if "quiver-launch.log" in files:
            count += 1
        if "decomp-buddy.json" in files:
            count += 1
    return count


def run_scan(quiver_base_dir, decompbuddy_base_dir, output_base_dir, log=None, progress=None, on_missing_exe=None):
    if log is None:
        log = lambda msg: print(msg)

    os.makedirs(output_base_dir, exist_ok=True)
    report = ScanReport(output_base_dir, log)

    quiver_output_dir = os.path.join(output_base_dir, "Quiver")
    decompbuddy_output_dir = os.path.join(output_base_dir, "DecompBuddy")

    os.makedirs(quiver_output_dir, exist_ok=True)
    os.makedirs(decompbuddy_output_dir, exist_ok=True)

    total = 0
    for base_dir in (quiver_base_dir, decompbuddy_base_dir):
        if os.path.exists(base_dir):
            total += count_target_files(base_dir)

    counter = {"handled": 0}

    def update_progress():
        if progress and total > 0:
            progress(min(1.0, counter["handled"] / total))

    def bump():
        counter["handled"] += 1
        update_progress()

    # SCHRITT 1: Cleanup - Ungültige Verknüpfungen entfernen
    if os.path.exists(quiver_output_dir):
        cleanup_invalid_shortcuts(quiver_output_dir, "Quiver", report, log)
    if os.path.exists(decompbuddy_output_dir):
        cleanup_invalid_shortcuts(decompbuddy_output_dir, "DecompBuddy", report, log)

    # SCHRITT 2: Neue Spiele finden und Verknüpfungen erstellen
    if os.path.exists(quiver_base_dir):
        process_quiver(quiver_base_dir, quiver_output_dir, report, log, bump)
    else:
        log(f"  [!] Quiver Basis-Ordner nicht gefunden: {quiver_base_dir}")
        report.add("FEHLER", "System", "N/A", quiver_base_dir, "Basis-Ordner nicht gefunden")

    if os.path.exists(decompbuddy_base_dir):
        process_decompbuddy(decompbuddy_base_dir, decompbuddy_output_dir, report, log, bump, on_missing_exe)
    else:
        log(f"  [!] DecompBuddy Basis-Ordner nicht gefunden: {decompbuddy_base_dir}")
        report.add("FEHLER", "System", "N/A", decompbuddy_base_dir, "Basis-Ordner nicht gefunden")

    if progress:
        progress(1.0)

    report.save()

    log("\n✅ Fertig!")
    log(f"💡 Tipp: Füge in Playnite unter 'Bibliothek integrieren' -> 'Von Ordner importieren' den Ordner '{output_base_dir}' hinzu.")

    return summarize(report)


# ==========================================
# GUI
# ==========================================
class App:
    def __init__(self, root):
        self.root = root
        self.busy = False
        self.scan_summary = None
        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        self.missing_exe_event = threading.Event()
        self.missing_exe_context = None

        root.title("Recomp-to-Playnite – Scan Tool")
        root.geometry("940x780")

        self.quiver_var = tk.StringVar(value=DEFAULT_CONFIG["quiver_base_dir"])
        self.decompbuddy_var = tk.StringVar(value=DEFAULT_CONFIG["decompbuddy_base_dir"])
        self.output_var = tk.StringVar(value=DEFAULT_CONFIG["output_base_dir"])

        self._build_ui()
        self.load_config()

    # ---------- UI ----------
    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)

        settings = ttk.LabelFrame(main, text="Basis-Pfade", padding=10)
        settings.pack(fill="x")

        ttk.Label(settings, text="Quiver Base Dir:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(settings, textvariable=self.quiver_var, width=60).grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        ttk.Button(settings, text="Durchsuchen…", command=self.browse_quiver).grid(row=0, column=2, padx=5, pady=2)

        ttk.Label(settings, text="DecompBuddy Base Dir:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(settings, textvariable=self.decompbuddy_var, width=60).grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        ttk.Button(settings, text="Durchsuchen…", command=self.browse_decompbuddy).grid(row=1, column=2, padx=5, pady=2)

        ttk.Label(settings, text="Output Base Dir:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(settings, textvariable=self.output_var, width=60).grid(row=2, column=1, sticky="ew", padx=5, pady=2)
        ttk.Button(settings, text="Durchsuchen…", command=self.browse_output).grid(row=2, column=2, padx=5, pady=2)

        settings.columnconfigure(1, weight=1)

        actions = ttk.Frame(main)
        actions.pack(fill="x", pady=(10, 0))

        self.scan_btn = ttk.Button(actions, text="🔍 Scan starten", command=self.start_scan)
        self.scan_btn.pack(side="left")

        ttk.Button(actions, text="Log leeren", command=self.clear_log).pack(side="left", padx=(8, 0))

        self.progress_bar = ttk.Progressbar(main, orient="horizontal", mode="determinate", maximum=100)
        self.progress_bar.pack(fill="x", pady=(0, 8))

        log_frame = ttk.Frame(main)
        log_frame.pack(fill="both", expand=True, pady=(10, 0))

        self.log_text = tk.Text(log_frame, height=20, width=120, state="disabled", wrap="none", font=("Consolas", 9))
        yscroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=yscroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")

        self.status_var = tk.StringVar(value="Status: Bereit")
        ttk.Label(main, textvariable=self.status_var, anchor="w").pack(fill="x")

    # ---------- Pfade ----------
    def load_config(self):
        cfg = read_config()
        self.quiver_var.set(cfg.get("quiver_base_dir", DEFAULT_CONFIG["quiver_base_dir"]))
        self.decompbuddy_var.set(cfg.get("decompbuddy_base_dir", DEFAULT_CONFIG["decompbuddy_base_dir"]))
        self.output_var.set(cfg.get("output_base_dir", DEFAULT_CONFIG["output_base_dir"]))

    def save_current_config(self):
        write_config({
            "quiver_base_dir": self.quiver_var.get().strip(),
            "decompbuddy_base_dir": self.decompbuddy_var.get().strip(),
            "output_base_dir": self.output_var.get().strip(),
        })

    def browse_quiver(self):
        path = filedialog.askdirectory(title="Quiver Base Dir wählen")
        if path:
            self.quiver_var.set(path)

    def browse_decompbuddy(self):
        path = filedialog.askdirectory(title="DecompBuddy Base Dir wählen")
        if path:
            self.decompbuddy_var.set(path)

    def browse_output(self):
        path = filedialog.askdirectory(title="Output Base Dir wählen")
        if path:
            self.output_var.set(path)

    # ---------- Scan ----------
    def start_scan(self):
        if self.busy:
            return
        quiver_base_dir = self.quiver_var.get().strip()
        decompbuddy_base_dir = self.decompbuddy_var.get().strip()
        output_base_dir = self.output_var.get().strip()

        if not quiver_base_dir or not decompbuddy_base_dir or not output_base_dir:
            messagebox.showwarning("Warnung", "Bitte alle drei Basis-Pfade angeben.")
            return

        self.busy = True
        self.scan_btn.config(state="disabled")
        self.status_var.set("Status: Scan läuft …")
        self.log_queue = queue.Queue()
        self.log_queue.put(f"\n--- Starte Scan: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n")
        self.root.after(100, self._poll_log)

        thread = threading.Thread(
            target=self._scan_worker,
            args=(quiver_base_dir, decompbuddy_base_dir, output_base_dir),
            daemon=True,
        )
        thread.start()

    def _scan_worker(self, quiver_base_dir, decompbuddy_base_dir, output_base_dir):
        try:
            self.scan_summary = run_scan(
                quiver_base_dir,
                decompbuddy_base_dir,
                output_base_dir,
                log=lambda msg: self.log_queue.put(str(msg)),
                progress=lambda f: self.progress_queue.put(("PCT", f)),
                on_missing_exe=self._request_missing_exe,
            )
        except Exception as e:
            self.log_queue.put(f"  [!] Unbekannter Fehler: {e}")
            self.scan_summary = None
        finally:
            self.log_queue.put(DONE_MARKER)

    def _on_scan_done(self):
        self.busy = False
        self.scan_btn.config(state="normal")
        self.progress_bar["value"] = 0
        if self.scan_summary is None:
            self.status_var.set("Status: Scan fehlgeschlagen – Fehler siehe Log")
        else:
            self.status_var.set(
                f"Status: Scan abgeschlossen | Neu hinzugefügt: {self.scan_summary['neu']} | "
                f"Entfernt: {self.scan_summary['entfernt']} | Übersprungen: {self.scan_summary['uebersprungen']} | Fehler: {self.scan_summary['fehler']}"
            )
        self.root.after(100, self._poll_log)

    def _poll_log(self):
        try:
            msg = self.log_queue.get_nowait()
        except queue.Empty:
            self._drain_progress()
            self.root.after(100, self._poll_log)
            return

        if msg == DONE_MARKER:
            self._on_scan_done()
            return

        if msg == MISSING_EXE_MARKER:
            self._handle_missing_exe_marker()
        else:
            self._insert_log(msg)

        self._drain_progress()
        self.root.after(100, self._poll_log)

    def _drain_progress(self):
        try:
            kind, value = self.progress_queue.get_nowait()
            if kind == "PCT":
                self.progress_bar["value"] = float(value) * 100
        except queue.Empty:
            pass

    def _handle_missing_exe_marker(self):
        ctx = self.missing_exe_context or {}
        suggested = ctx.get("suggested") or ""
        game_name = ctx.get("game_name") or "Spiel"
        path = filedialog.askopenfilename(
            title=f"EXE für {game_name} auswählen",
            filetypes=[("Ausführbare Dateien", "*.*"), ("Alle Dateien", "*.*")],
            initialdir=os.path.dirname(suggested) or os.path.expanduser("~"),
        )
        self.missing_exe_context["chosen"] = path
        self.missing_exe_event.set()
        self._insert_log(f"  [+] Manuelles EXE gewählt: {path}")

    def _request_missing_exe(self, game_name, suggested_exe):
        self.missing_exe_context = {"game_name": game_name, "suggested": suggested_exe, "chosen": None}
        self.missing_exe_event.clear()
        self.log_queue.put(MISSING_EXE_MARKER)
        self.missing_exe_event.wait()
        chosen = self.missing_exe_context.get("chosen")
        return chosen or None

    def _insert_log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, str(msg) + "\n")
        self.log_text.configure(state="disabled")
        self.log_text.see(tk.END)

    def clear_log(self):
        if self.busy:
            return
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

    # ---------- Events ----------
    def _on_close(self):
        self.save_current_config()
        self.root.destroy()

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()


def main():
    root = tk.Tk()
    App(root).run()


if __name__ == "__main__":
    main()
