@echo off
rem Starts the SCCM Maintenance App from the folder this batch file is in
start "" /b powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0SCCM-MaintenanceApp.ps1"
