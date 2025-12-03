import os
import json
import shutil
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

# --- Constants ---
TASKBAR_DIR = Path(os.getenv("APPDATA")) / "Microsoft" / "Internet Explorer" / "Quick Launch" / "User Pinned" / "TaskBar"
BACKUP_DIR = Path(__file__).parent / "taskbar_backup"
LAYOUT_FILE = Path(__file__).parent / "taskbar_layout.json"

# --- Helpers ---

def get_classic_pinned():
    return sorted([f.name for f in TASKBAR_DIR.glob("*.lnk")])

def backup_classic_shortcuts():
    BACKUP_DIR.mkdir(exist_ok=True)
    for shortcut in TASKBAR_DIR.glob("*.lnk"):
        shutil.copy2(shortcut, BACKUP_DIR)
    return len(list(BACKUP_DIR.glob("*.lnk")))

def restore_classic_shortcuts():
    if not BACKUP_DIR.exists():
        return 0

    restored = 0
    log = []

    # Clear current .lnk files
    for shortcut in TASKBAR_DIR.glob("*.lnk"):
        try:
            shortcut.unlink()
            log.append(f"Deleted: {shortcut.name}")
        except Exception as e:
            log.append(f"Failed to delete: {shortcut.name} ({e})")

    # Copy from backup
    for file in BACKUP_DIR.glob("*.lnk"):
        try:
            shutil.copy2(file, TASKBAR_DIR)
            log.append(f"Restored: {file.name}")
            restored += 1
        except Exception as e:
            log.append(f"Failed to restore: {file.name} ({e})")

    with open("restore_log.txt", "w") as f:
        f.write("\n".join(log))

    return restored


def restart_explorer():
    subprocess.run("taskkill /f /im explorer.exe", shell=True)
    subprocess.run("start explorer.exe", shell=True)

def get_uwp_apps():
    script = r"""
    $apps = (New-Object -ComObject shell.application).Namespace('shell:AppsFolder').Items()
    $apps | ForEach-Object { $_.Name } | Sort-Object -Unique
    """
    result = subprocess.run(["powershell", "-Command", script], capture_output=True, text=True)
    return sorted(list(filter(None, result.stdout.splitlines())))

def save_layout(classic, uwp):
    with open(LAYOUT_FILE, "w") as f:
        json.dump({"classic": classic, "uwp": uwp}, f, indent=2)

def load_layout():
    if not LAYOUT_FILE.exists():
        return {"classic": [], "uwp": []}
    with open(LAYOUT_FILE, "r") as f:
        return json.load(f)

# --- GUI ---

class LayoutSnapApp:
    def __init__(self, master):
        self.master = master
        master.title("LayoutSnap - Taskbar Snapshot")
        master.geometry("400x460")
        master.resizable(False, False)

        self.show_filtered = tk.BooleanVar(value=False)
        self.uwp_apps = get_uwp_apps()
        self.classic_apps = get_classic_pinned()

        # UI Elements
        ttk.Label(master, text="Taskbar Snapshot Tool", font=("Segoe UI", 14, "bold")).pack(pady=10)

        ttk.Button(master, text="💾 Backup Pinned Shortcuts", command=self.backup).pack(pady=5)
        ttk.Button(master, text="♻ Restore Pinned Shortcuts", command=self.restore).pack(pady=5)
        ttk.Button(master, text="🔃 Restart Explorer", command=restart_explorer).pack(pady=5)
        ttk.Button(master, text="📋 View Saved Layout", command=self.show_saved).pack(pady=5)

        ttk.Checkbutton(master, text="Filter likely pinned UWP apps", variable=self.show_filtered,
                        command=self.update_uwp_list).pack(pady=10)

        self.text = tk.Text(master, wrap="word", height=15)
        self.text.pack(expand=True, fill="both", padx=10)
        self.text.config(state="disabled")

        self.update_uwp_list()

    def backup(self):
        count = backup_classic_shortcuts()
        self.classic_apps = get_classic_pinned()
        save_layout(self.classic_apps, self.uwp_apps)
        messagebox.showinfo("Success", f"Backed up {count} pinned shortcuts and saved layout.")

    def restore(self):
        count = restore_classic_shortcuts()
        messagebox.showinfo("Restored", f"Restored {count} shortcuts to taskbar.\nYou may need to restart Explorer.")

    def update_uwp_list(self):
        self.text.config(state="normal")
        self.text.delete("1.0", "end")

        if self.show_filtered.get():
            filtered = [app for app in self.uwp_apps if any(name.lower() in app.lower() for name in self.classic_apps)]
            self.text.insert("end", "📌 Likely Pinned UWP Apps:\n" + "\n".join(f"• {app}" for app in filtered))
        else:
            self.text.insert("end", "📋 All UWP Apps:\n" + "\n".join(f"• {app}" for app in self.uwp_apps))

        self.text.config(state="disabled")

    def show_saved(self):
        # Toggle logic: Close if already open
        if hasattr(self, 'layout_window') and self.layout_window.winfo_exists():
            self.layout_window.destroy()
            return

        # Otherwise, open new window
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


# --- Run ---
if __name__ == "__main__":
    root = tk.Tk()
    app = LayoutSnapApp(root)
    root.mainloop()
