import os
import shutil
import subprocess
import re
import ttkbootstrap as tb
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

def is_duplicate(file_name):
    return re.search(r"\(\d+\)\.lnk$", file_name) is not None

class TaskbarBackupApp:
    def __init__(self, master):
        self.master = master
        master.title("Taskbar Backup - Pinned Shortcuts Saver")
        master.geometry("600x600")  # increased height for new button
        master.minsize(600, 600)
        master.resizable(False, False)  # Disable resizing/maximizing

        self.backup_dir = Path(__file__).parent / "taskbar_backup"
        self.layout_window = None

        # Setup ttk style for colored buttons
        style = tb.Style()


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

        ttk.Label(folder_frame, text="Backup Folder:",
            font=("Segoe UI", 10, "bold")).pack(side="left", padx=(0, 5))


        self.backup_folder_var = tk.StringVar(value=str(self.backup_dir))
        self.backup_folder_entry = ttk.Entry(folder_frame, textvariable=self.backup_folder_var, state="readonly")
        self.backup_folder_entry.pack(side="left", fill="x", expand=True)

        self.change_folder_btn = ttk.Button(folder_frame, text="Change...", command=self.change_backup_folder)
        self.change_folder_btn.pack(side="left", padx=(5, 0))

        note_text = "Note: Right-click each shortcut in the opened folder to 'Pin to taskbar' manually."
        self.note_label = ttk.Label(master, text=note_text,
                            font=("Segoe UI", 10, "bold"),
                            foreground="black")

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

    def backup(self):
        try:
            # Make sure backup folder exists
            self.backup_dir.mkdir(parents=True, exist_ok=True)

            # List current pinned shortcuts
            current_shortcuts = {f.name for f in TASKBAR_DIR.glob("*.lnk") if not is_duplicate(f.name)}

            # List existing backed-up shortcuts
            existing_shortcuts = {f.name for f in self.backup_dir.glob("*.lnk")}

            # Compare sets
            if current_shortcuts == existing_shortcuts:
                self.log("No new shortcuts to save — everything already backed up.")
                return

            # Otherwise, perform a fresh backup
            self.log("Backing up pinned shortcuts...")

            # Clear old backup
            for f in self.backup_dir.glob("*"):
                if f.is_file():
                    f.unlink()

            # Copy in fresh versions
            count = 0
            for shortcut in TASKBAR_DIR.glob("*.lnk"):
                if is_duplicate(shortcut.name):
                    continue
                shutil.copy2(shortcut, self.backup_dir)
                count += 1

            self.log(f"Backed up {count} classic pinned shortcuts.")

        except Exception as e:
            self.log(f"Backup failed: {e}")


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
            # Minimize the window before taking screenshot
            self.master.withdraw()
            self.master.after(200)

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

        finally:
            # Restore window after screenshot/save
            self.master.deiconify()

def main():
    root = tb.Window(themename="litera")
    app = TaskbarBackupApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
