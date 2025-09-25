# Script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Absolute path to pythonw.exe (adjusted to actual location)
$pythonExe = "C:\Users\NADLUROB\Documents\PythonProject\.venv\Scripts\pythonw.exe"

# Absolute path to GUI script (inside the same folder as this launch script)
$guiScript = Join-Path $scriptDir "gui.py"

# Launch GUI invisibly
Start-Process -FilePath $pythonExe -ArgumentList "`"$guiScript`"" -WindowStyle Hidden
