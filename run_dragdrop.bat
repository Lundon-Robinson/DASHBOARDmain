@echo off
setlocal
REM -- Set path to correct Python interpreter --
set PYTHON=C:\Users\NADLUROB\Documents\PythonProject\.venv\Scripts\python.exe
set SCRIPT=%~dp0process_delegation.py

"%PYTHON%" "%SCRIPT%" "%~1"
pause
