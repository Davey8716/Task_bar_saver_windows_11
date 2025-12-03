import os
import shutil
import subprocess
import json
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

APPDATA = os.getenv("APPDATA")
TASKBAR_DIR = Path(APPDATA) / r"Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
BACKUP_DIR = Path(__file__).parent / "taskbar_backup"
LAYOUT_FILE = Path(__file__).parent / "layout.json"

def restart_explorer():
    subprocess.run("taskkill /f /im explorer.exe", shell=True)
    subprocess.run("start explorer.exe", shell=True)

def backup_classic_shortcuts():
    if not TASKBAR_DIR.exists():
        return 0
    BACKUP_DIR.mkdir(exist_ok=True)
    # Clear backup folder first
    for f in BACKUP_DIR.glob("*"):
        if f.is_file():
            f.unlink()
    count = 0
    for shortcut in TASKBAR_DIR.glob("*.lnk"):
        try:
            shutil.copy2(shortcut, BACKUP_DIR)
            count += 1
        except Exception as e:
            print(f"Failed to backup {shortcut}: {e}")
    return count

def backup(self):
    self.log("Backup started")
    count = backup_classic_shortcuts()
    self.log(f"Backed up {count} classic pinned shortcuts.")
    # ...
    self.log("Backup finished")


def list_uwp_apps():
    # Placeholder for real UWP app listing
    return [
        "Microsoft Store",
        "Photos",
        "Mail",
        "Calculator",
    ]

def save_layout(classic, uwp):
    layout = {
        "classic": classic,
        "uwp": uwp,
    }
    with open(LAYOUT_FILE, "w", encoding="utf-8") as f:
        json.dump(layout, f, indent=2)

def load_layout():
    if not LAYOUT_FILE.exists():
        return {"classic": [], "uwp": []}
    with open(LAYOUT_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

class LayoutSnapApp:
    def __init__(self, master):
        self.master = master
        master.title("LayoutSnap - Taskbar & Desktop Backup")
        master.geometry("420x380")

        self.layout_window = None

        ttk.Label(master, text="LayoutSnap", font=("Segoe UI", 18, "bold")).pack(pady=10)

        ttk.Button(master, text="💾 Backup Pinned Shortcuts", command=self.backup).pack(fill="x", padx=20, pady=5)
        ttk.Button(master, text="🧷 Open Backup Folder to Re-pin", command=self.open_backup_folder).pack(fill="x", padx=20, pady=5)
        ttk.Button(master, text="🔃 Restart Explorer", command=restart_explorer).pack(fill="x", padx=20, pady=5)
        ttk.Button(master, text="📋 View Saved Layout", command=self.show_saved).pack(fill="x", padx=20, pady=5)

        self.status_log = scrolledtext.ScrolledText(master, height=8, state="disabled")
        self.status_log.pack(fill="both", expand=True, padx=10, pady=10)

    def log(self, msg):
        self.status_log.config(state="normal")
        self.status_log.insert("end", msg + "\n")
        self.status_log.see("end")
        self.status_log.config(state="disabled")

    def backup(self):
        self.log("Backing up pinned shortcuts...")
        count = backup_classic_shortcuts()
        self.log(f"Backed up {count} classic pinned shortcuts.")

        self.log("Listing UWP pinned apps...")
        uwp_list = list_uwp_apps()
        self.log(f"Found {len(uwp_list)} UWP apps.")

        classic_list = [f.name for f in BACKUP_DIR.glob("*.lnk")] if BACKUP_DIR.exists() else []

        save_layout(classic_list, uwp_list)
        self.log("Layout saved to layout.json.")

    def open_backup_folder(self):
        if not BACKUP_DIR.exists():
            messagebox.showwarning("Folder not found", f"No backup folder found at:\n{BACKUP_DIR}")
            return
        subprocess.run(f'explorer "{BACKUP_DIR}"', shell=True)
        messagebox.showinfo(
            "Manual Re-pin Required",
            "Please right-click each shortcut in the opened folder and choose 'Pin to taskbar' manually."
        )

    def show_saved(self):
        # Toggle the saved layout window on/off
        if self.layout_window and self.layout_window.winfo_exists():
            self.layout_window.destroy()
            self.layout_window = None
            return

        layout = load_layout()

        self.layout_window = tk.Toplevel(self.master)
        self.layout_window.title("Saved Layout")
        self.layout_window.geometry("400x300")

        txt = tk.Text(self.layout_window, wrap="word")
        txt.pack(expand=True, fill="both")

        txt.insert("end", "📎 Classic Shortcuts:\n")
        for item in layout.get("classic", []):
            txt.insert("end", f"• {item}\n")

        txt.insert("end", "\n🧩 UWP Apps:\n")
        for item in layout.get("uwp", []):
            txt.insert("end", f"• {item}\n")

        txt.config(state="disabled")

def main():
    root = tk.Tk()
    app = LayoutSnapApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
