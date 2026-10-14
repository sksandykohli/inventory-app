@echo off
cd /d "%~dp0"
rem To keep backups in Google Drive / OneDrive, remove "rem" below and put your folder path:
rem set INVENTORY_BACKUP_DIR=G:\My Drive\InventoryBackup
python app.py
pause
