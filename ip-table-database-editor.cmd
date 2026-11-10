@echo off
rem Open the IP table database editor, in Windows Python (openpyxl and PySide6 installed there):
rem   ip-table-database-editor.cmd ["C:\...\Network Summary.xlsx" or sweep.json] [--db devices.json]
rem   pinger.exe -json 172.22.0.1/24 | ip-table-database-editor.cmd
rem Console python (iptable.launch) reads a pipe if there is one, starts the window in pythonw and returns.
set "PYTHONPATH=%~dp0;%PYTHONPATH%"
where python >nul 2>nul
if %errorlevel%==0 (python -m iptable.launch %*) else (py -3 -m iptable.launch %*)
