import os
import shutil
import subprocess
import json
import re
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from datetime import datetime

# Pillow for screenshotp
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
    restart_explorer()

def is_duplicate(file_name):
    return re.search(r"\(\d+\)\.lnk$", file_name) is not None

class TaskbarBackupApp:
    def __init__(self, master):
        self.master = master
        master.title("Taskbar Backup - Pinned Shortcuts Saver")
        master.geometry("580x550")  # increased height for new button
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
                
        ttk.Button(master,
                text="📌 Restore Pinned Shortcuts (Full Replace)",
                command=self.restore_pinned_shortcuts,
                style="Backup.TButton").pack(fill="x", padx=20, pady=5)

        # Clear all pinned shortcuts from the taskbar
        ttk.Button(master,
                text="🧹 Clear ALL Taskbar Pins",
                command=self.clear_all_taskbar_pins,
                style="Restart.TButton").pack(fill="x", padx=20, pady=5)

        ttk.Button(master, text="📋 View Saved Layout",
                command=self.show_saved,
                style="ViewLayout.TButton").pack(fill="x", padx=20, pady=5)

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

    def show_saved(self):
        if self.layout_window and self.layout_window.winfo_exists():
            self.layout_window.destroy()
            self.layout_window = None
            return

        layout = self.load_layout()

        self.layout_window = tk.Toplevel(self.master)
        self.layout_window.title("Saved Layout")
        self.layout_window.geometry("400x300")

        txt = tk.Text(self.layout_window, wrap="word")
        txt.pack(expand=True, fill="both")

        txt.insert("end", "📎 Classic Shortcuts:\n")
        for item in layout.get("classic", []):
            txt.insert("end", f"• {item}\n")

        txt.config(state="disabled")

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
            
    def clear_all_taskbar_pins(self):
        confirm = messagebox.askyesno(
            "Clear ALL Taskbar Pins",
            "This will remove ALL pinned programs from your taskbar for this user.\n\n"
            "Make sure you have a backup first.\n\n"
            "Do you want to continue?"
        )
        if not confirm:
            self.log("Clear all taskbar pins canceled by user.")
            return

        self.log("Clearing all taskbar pins...")

        # 1) Delete all .lnk files in the pinned Taskbar folder
        removed = 0
        if TASKBAR_DIR.exists():
            for shortcut in TASKBAR_DIR.glob("*.lnk"):
                try:
                    shortcut.unlink()
                    removed += 1
                except Exception as e:
                    self.log(f"Failed to delete pinned shortcut {shortcut}: {e}")
        else:
            self.log(f"Pinned taskbar folder does not exist: {TASKBAR_DIR}")

        self.log(f"Deleted {removed} .lnk files from pinned taskbar folder.")

        # 2) Clear Taskband registry key (removes cached taskbar layout)
        try:
            cmd = r'reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband" /f'
            completed = subprocess.run(cmd, shell=True)
            if completed.returncode == 0:
                self.log("Successfully cleared Taskband registry key (taskbar layout cache).")
            else:
                self.log(f"Failed to clear Taskband registry key. Exit code: {completed.returncode}")
        except Exception as e:
            self.log(f"Error while clearing Taskband registry key: {e}")

        # 3) Ask to restart Explorer so the taskbar updates cleanly
        restart = messagebox.askyesno(
        "Restart Explorer?",
        "All taskbar pins have been cleared (files + cache).\n\n"
        "To avoid 'item has been moved or deleted' errors when clicking old pins,\n"
        "Windows Explorer should be restarted.\n\n"
        "Restart Explorer now?"
        )
        if restart:
            self.log("Restarting Explorer to apply cleared pins...")
            try:
                restart_explorer()
                self.log("Explorer restarted successfully.")
            except Exception as e:
                self.log(f"Failed to restart Explorer: {e}")
        else:
            self.log("Explorer restart skipped — taskbar may show stale pins until you restart Explorer manually.")


        self.log("Clear pins complete.")

    def restore_pinned_shortcuts(self):
        """
        Restore classic pinned shortcuts by copying all .lnk files
        from the backup folder into the real Taskbar pinned folder,
        replacing anything that is there now.
        """
        # Ensure backup folder exists
        if not self.backup_dir.exists():
            messagebox.showwarning(
                "Backup Folder Not Found",
                f"No backup folder found at:\n{self.backup_dir}"
            )
            return

        # Collect .lnk files in the backup folder
        lnk_files = list(self.backup_dir.glob("*.lnk"))
        if not lnk_files:
            messagebox.showinfo(
                "No Shortcuts Found",
                "No .lnk files were found in the selected backup folder."
            )
            return

        # Confirm with the user (this replaces current pins)
        confirm = messagebox.askyesno(
            "Confirm Full Restore",
            "This will REPLACE your current pinned taskbar shortcuts\n"
            "with the ones stored in this backup folder.\n\n"
            "Do you want to continue?"
        )
        if not confirm:
            self.log("Restore canceled by user.")
            return

        # Make sure the Taskbar pinned folder exists
        try:
            TASKBAR_DIR.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self.log(f"Failed to ensure taskbar directory exists: {e}")
            messagebox.showerror(
                "Error",
                f"Failed to access the taskbar pinned folder:\n{e}"
            )
            return

        self.log(f"Restoring pinned shortcuts from: {self.backup_dir}")

        # Remove existing pinned .lnk files
        removed = 0
        for shortcut in TASKBAR_DIR.glob("*.lnk"):
            try:
                shortcut.unlink()
                removed += 1
            except Exception as e:
                self.log(f"Failed to remove existing shortcut {shortcut}: {e}")

        # Copy backup .lnk files into the pinned folder
        added = 0
        for src in lnk_files:
            try:
                dst = TASKBAR_DIR / src.name
                shutil.copy2(src, dst)
                added += 1
            except Exception as e:
                self.log(f"Failed to restore shortcut {src}: {e}")

        self.log(f"Removed {removed} existing pinned shortcuts.")
        self.log(f"Restored {added} shortcuts from backup.")

        if added == 0:
            messagebox.showinfo(
                "Nothing Restored",
                "No shortcuts were restored (all copies failed)."
            )
            return

        # Ask user if they want to restart Explorer now
        restart = messagebox.askyesno(
            "Restart Explorer?",
            "Pinned shortcuts have been restored.\n\n"
            "To apply changes to the taskbar, Windows Explorer usually "
            "needs to be restarted.\n\n"
            "Restart Explorer now?"
        )
        if restart:
            self.log("Restarting Explorer to apply restored pins...")
            restart_explorer()
        else:
            self.log("Restore complete (Explorer restart skipped by user).")

def main():
    root = tk.Tk()
    app = TaskbarBackupApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
