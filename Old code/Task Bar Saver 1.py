import os
import json
import subprocess
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

layout_file = Path(__file__).parent / "taskbar_full_layout.json"

def get_taskbar_items():
    # Run PowerShell command to read taskbar layout
    powershell_script = r"""
    $pins = (New-Object -ComObject shell.application).Namespace('shell:AppsFolder').Items() |
        Where-Object { $_.IsLink -eq $false } |
        Select-Object -ExpandProperty Name

    # Get traditional pinned shortcuts
    $classicPath = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    $classicPins = Get-ChildItem -Path $classicPath -Filter *.lnk | Select-Object -ExpandProperty Name

    $output = @{
        "classic" = $classicPins
        "uwp" = $pins
    }

    $output | ConvertTo-Json
    """

    result = subprocess.run(["powershell", "-Command", powershell_script],
                            capture_output=True, text=True)

    if result.returncode != 0:
        print("PowerShell error:", result.stderr)
        return None

    try:
        layout = json.loads(result.stdout)
        return layout
    except Exception as e:
        print("JSON parse error:", e)
        return None

def save_layout():
    layout = get_taskbar_items()
    if not layout:
        messagebox.showerror("Error", "Failed to get taskbar items.")
        return

    with open(layout_file, "w") as f:
        json.dump(layout, f, indent=2)

    messagebox.showinfo("Saved", f"Saved {len(layout['classic']) + len(layout['uwp'])} items to:\n{layout_file}")

def show_layout():
    if not layout_file.exists():
        messagebox.showerror("Error", "No saved layout file.")
        return

    with open(layout_file, "r") as f:
        layout = json.load(f)

    win = tk.Toplevel()
    win.title("Saved Taskbar Items")
    win.geometry("400x300")

    txt = tk.Text(win, wrap="word")
    txt.pack(expand=True, fill="both")

    txt.insert("end", "🔹 Classic Shortcuts (.lnk):\n")
    for item in layout.get("classic", []):
        txt.insert("end", f"• {item}\n")

    txt.insert("end", "\n🔹 UWP Apps:\n")
    for item in layout.get("uwp", []):
        txt.insert("end", f"• {item}\n")

    txt.config(state="disabled")

# GUI
root = tk.Tk()
root.title("LayoutSnap - Taskbar Snapshot")
root.geometry("300x180")
root.resizable(False, False)

tk.Label(root, text="Full Taskbar Snapshot", font=("Segoe UI", 14, "bold")).pack(pady=10)

tk.Button(root, text="💾 Save All Items", width=22, command=save_layout).pack(pady=5)
tk.Button(root, text="📋 View Saved Layout", width=22, command=show_layout).pack(pady=5)
tk.Button(root, text="❌ Exit", width=22, command=root.quit).pack(pady=5)

root.mainloop()
