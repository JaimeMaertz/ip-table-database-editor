@echo off
rem The IP table database editor's command line, from any folder, in Windows Python (openpyxl installed there):
rem   ip-table-database-editor-cli.cmd "C:\...\Network Summary.xlsx" [--write] [--db devices.json] [--jobsite NAME] [-v]
rem   pinger.exe -json 172.22.0.1/24 | ip-table-database-editor-cli.cmd -      what a sweep found, new or known
set "PYTHONPATH=%~dp0;%PYTHONPATH%"
where python >nul 2>nul
if %errorlevel%==0 (python -m iptable %*) else (py -3 -m iptable %*)
