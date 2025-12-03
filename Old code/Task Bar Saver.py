import os
import json
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import messagebox

# Constants
taskbar_folder = Path(os.getenv('APPDATA')) / "Microsoft" / "Internet Explorer" / "Quick Launch" / "User Pinned" / "TaskBar"
layout_file = Path(__file__).parent / "taskbar_layout.json"

# Functions
def save_taskbar_layout():
    layout = []
    for file in taskbar_folder.glob("*.lnk"):
        layout.append({
            "name": file.stem,
            "shortcut_path": str(file)
        })
    with open(layout_file, "w") as f:
        json.dump(layout, f, indent=2)
    messagebox.showinfo("Saved", f"Saved {len(layout)} pinned items to:\n{layout_file}")

def restore_taskbar_layout():
    if not layout_file.exists():
        messagebox.showerror("Error", "No saved layout file found.")
        return

    with open(layout_file, "r") as f:
        layout = json.load(f)

    # Remove current pinned shortcuts
    for file in taskbar_folder.glob("*.lnk"):
        try:
            file.unlink()
        except Exception as e:
            print(f"Error deleting {file.name}: {e}")

    # Restore saved shortcuts
    missing = []
    for item in layout:
        src = Path(item["shortcut_path"])
        if src.exists():
            dst = taskbar_folder / src.name
            try:
                shutil.copy2(src, dst)
            except Exception as e:
                print(f"Error restoring {src.name}: {e}")
        else:
            missing.append(src.name)

    msg = "Restore complete!"
    if missing:
        msg += f"\nMissing shortcuts not restored:\n" + "\n".join(missing)
    msg += "\n\nRestart Explorer to see changes."

    messagebox.showinfo("Restore", msg)

# GUI Setup
root = tk.Tk()
root.title("LayoutSnap - Taskbar Manager")
root.geometry("300x160")
root.resizable(False, False)

tk.Label(root, text="Taskbar Layout Backup", font=("Segoe UI", 14, "bold")).pack(pady=10)

tk.Button(root, text="💾 Save Layout", width=20, command=save_taskbar_layout).pack(pady=5)
tk.Button(root, text="🔁 Restore Layout", width=20, command=restore_taskbar_layout).pack(pady=5)
tk.Button(root, text="❌ Exit", width=20, command=root.quit).pack(pady=5)

root.mainloop()
