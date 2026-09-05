"""
Create a Desktop shortcut that launches the timer widget with NO console window.

Just run once:   python create_shortcut.py
Then double-click "Desk Timer" on your Desktop.
"""

import os
import sys
import subprocess

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TARGET = os.path.join(SCRIPT_DIR, "timer_widget.pyw")

# pythonw.exe runs the script windowless (no black console).
pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
if not os.path.exists(pythonw):
    pythonw = sys.executable  # fallback

# Build the .lnk via a small PowerShell command (no extra dependencies).
# Ask Windows for the REAL Desktop folder (handles OneDrive redirection).
ps = f'''
$ws = New-Object -ComObject WScript.Shell
$desktop = $ws.SpecialFolders("Desktop")
$s = $ws.CreateShortcut((Join-Path $desktop "Desk Timer.lnk"))
$s.TargetPath = "{pythonw}"
$s.Arguments = '"{TARGET}"'
$s.WorkingDirectory = "{SCRIPT_DIR}"
$s.WindowStyle = 7
$s.IconLocation = "{pythonw},0"
$s.Description = "Transparent desktop timer"
$s.Save()
Write-Output ("Shortcut created: " + (Join-Path $desktop "Desk Timer.lnk"))
'''

subprocess.run(
    ["powershell", "-NoProfile", "-Command", ps],
    check=True,
)
