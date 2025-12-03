import os
import shutil
import subprocess
import json
import re
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from datetime import datetime

# Pillow for screenshot
from PIL import ImageGrab
# Fix DPI scaling on Windows for crisp UI
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

APPDATA = os.getenv("APPDATA")
TASKBAR_DIR = Path(APPDATA) / r"Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
LAYOUT_FILE = Path(__file__).parent / "layout.json"

def restart_explorer():
    try:
        subprocess.run("taskkill /f /im explorer.exe", shell=True)
        subprocess.run("start explorer.exe", shell=True)
    except Exception as e:
        print(f"Failed to restart Explorer: {e}")


def is_duplicate(file_name):
    return re.search(r"\(\d+\)\.lnk$", file_name) is not None

class TaskbarBackupApp:
    def __init__(self, master):
        self.master = master
        master.title("Taskbar Backup - Pinned Shortcuts Saver")
        master.geometry("800x800")  # increased height for new button
        master.minsize(800, 800)
        master.resizable(False, False)  # Disable resizing/maximizing

        self.backup_dir = Path(__file__).parent / "taskbar_backup"
        self.layout_window = None

        # Setup ttk style for colored buttons
        style = ttk.Style(master)
        style.theme_use('clam')

        style.configure("Backup.TButton",
                        background="#4CAF50",
                        foreground="white",
                        font=("Segoe UI", 10, "bold"),
                        padding=8)
        style.map("Backup.TButton",
                background=[('active', '#45a049'), ('pressed', '#3e8e41')])

        style.configure("OpenBackup.TButton",
                        background="#2196F3",
                        foreground="white",
                        font=("Segoe UI", 10, "bold"),
                        padding=8)
        style.map("OpenBackup.TButton",
                background=[('active', '#1976D2'), ('pressed', '#1565C0')])
        
        style.configure("Restart.TButton",
                        background="#FF5722",
                        foreground="white",
                        font=("Segoe UI", 10, "bold"),
                        padding=8)
        style.map("Restart.TButton",
                background=[('active', '#E64A19'), ('pressed', '#D84315')])

        style.configure("ViewLayout.TButton",
                        background="#9C27B0",
                        foreground="white",
                        font=("Segoe UI", 10, "bold"),
                        padding=8)
        style.map("ViewLayout.TButton",
                background=[('active', '#7B1FA2'), ('pressed', '#6A1B9A')])

        style.configure("Screenshot.TButton",
                        background="#607D8B",
                        foreground="white",
                        font=("Segoe UI", 10, "bold"),
                        padding=8)
        style.map("Screenshot.TButton",
                background=[('active', '#455A64'), ('pressed', '#37474F')])

        ttk.Label(master, text="Taskbar Backup", font=("Segoe UI", 18, "bold")).pack(pady=10)

        ttk.Button(master, text="💾 Backup Pinned Shortcuts",
                command=self.backup,
                style="Backup.TButton").pack(fill="x", padx=20, pady=5)

        self.open_backup_btn = ttk.Button(master, text="🧷 Open Backup Folder to view saved Pins",
                                        command=self.open_backup_folder,
                                        style="OpenBackup.TButton")
        self.open_backup_btn.pack(fill="x", padx=20, pady=5)
                


        # NEW Desktop Screenshot button
        ttk.Button(master, text="🖼️ Desktop Screenshot",
                command=self.desktop_screenshot,
                style="Screenshot.TButton").pack(fill="x", padx=20, pady=5)

        self.status_log = scrolledtext.ScrolledText(master, height=8, state="disabled")
        self.status_log.pack(fill="both", expand=True, padx=10, pady=10)
        
        
        # Backup folder selector frame
        folder_frame = ttk.Frame(master)
        folder_frame.pack(fill="x", padx=20, pady=5)

        ttk.Label(folder_frame, text="Backup Folder:").pack(side="left", padx=(0, 5))

        self.backup_folder_var = tk.StringVar(value=str(self.backup_dir))
        self.backup_folder_entry = ttk.Entry(folder_frame, textvariable=self.backup_folder_var, state="readonly")
        self.backup_folder_entry.pack(side="left", fill="x", expand=True)

        self.change_folder_btn = ttk.Button(folder_frame, text="Change...", command=self.change_backup_folder)
        self.change_folder_btn.pack(side="left", padx=(5, 0))

        note_text = "Note: Right-click each shortcut in the opened folder to 'Pin to taskbar' manually."
        self.note_label = ttk.Label(master, text=note_text, font=("Segoe UI", 8), foreground="gray")
        self.note_label.pack(fill="x", padx=25, pady=(0, 10))

    def log(self, msg):
        self.status_log.config(state="normal")
        self.status_log.insert("end", msg + "\n")
        self.status_log.see("end")
        self.status_log.config(state="disabled")

    def backup_classic_shortcuts(self):
        if not TASKBAR_DIR.exists():
            return 0
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        for f in self.backup_dir.glob("*"):
            if f.is_file():
                f.unlink()
        count = 0
        for shortcut in TASKBAR_DIR.glob("*.lnk"):
            if is_duplicate(shortcut.name):
                continue
            try:
                shutil.copy2(shortcut, self.backup_dir)
                count += 1
            except Exception as e:
                self.log(f"Failed to backup {shortcut}: {e}")
        return count

    def save_layout(self, classic):
        layout = {
            "classic": classic
        }
        with open(LAYOUT_FILE, "w", encoding="utf-8") as f:
            json.dump(layout, f, indent=2)

    def load_layout(self):
        if not LAYOUT_FILE.exists():
            return {"classic": []}

        with open(LAYOUT_FILE, "r", encoding="utf-8") as f:
            layout = json.load(f)

        # Remove any leftover UWP entries from old backups
        layout.pop("uwp", None)

        # Remove any Windows duplicate shortcuts like "Chrome (2).lnk"
        classic_filtered = [
            name for name in layout.get("classic", [])
            if not is_duplicate(name)
        ]

        layout["classic"] = classic_filtered
        return layout

    def backup(self):
        self.log("Backing up pinned shortcuts...")
        count = self.backup_classic_shortcuts()
        self.log(f"Backed up {count} classic pinned shortcuts.")

        classic_list = [f.name for f in self.backup_dir.glob("*.lnk")] if self.backup_dir.exists() else []

        self.save_layout(classic_list)
        self.log("Layout saved to layout.json.")

    def open_backup_folder(self):
        if not self.backup_dir.exists():
            messagebox.showwarning("Folder not found", f"No backup folder found at:\n{self.backup_dir}")
            return
        subprocess.run(f'explorer "{self.backup_dir}"', shell=True)

    def change_backup_folder(self):
        new_folder = filedialog.askdirectory(title="Select Backup Folder", initialdir=str(self.backup_dir))
        if new_folder:
            self.backup_dir = Path(new_folder)
            self.backup_folder_var.set(str(self.backup_dir))
            self.log(f"Backup folder changed to: {self.backup_dir}")

    def desktop_screenshot(self):
        try:
            # Grab the full screen
            img = ImageGrab.grab()
            # Ask user where to save
            initial_dir = str(self.backup_dir) if self.backup_dir.exists() else str(Path.home())
            save_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG Image", "*.png")],
                initialdir=initial_dir,
                initialfile=f"DesktopScreenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                title="Save Desktop Screenshot"
            )
            if save_path:
                img.save(save_path)
                self.log(f"Desktop screenshot saved: {save_path}")
            else:
                self.log("Screenshot canceled by user.")
        except Exception as e:
            self.log(f"Failed to take screenshot: {e}")
            messagebox.showerror("Error", f"Failed to take screenshot:\n{e}")
            
    

def main():
    root = tk.Tk()
    app = TaskbarBackupApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
