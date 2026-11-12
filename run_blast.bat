@echo off
:: ============================================================
:: run_blast.bat — Trigger untuk Windows Task Scheduler
:: Jalankan blast_worker.py dengan environment dari .env
:: ============================================================

:: Ganti path di bawah sesuai lokasi repo ini di komputer kamu
set REPO_DIR=C:\Users\fresn\OneDrive\Documents\INSP-Project\WA automation server - Copy

:: Pindah ke direktori repo (penting agar .env terbaca dengan benar)
cd /d "%REPO_DIR%"

:: Jalankan Python — pastikan python.exe sudah ada di PATH
:: Kalau pakai virtual environment, aktifkan dulu:
::   call "%REPO_DIR%\venv\Scripts\activate.bat"
python blast_worker.py >> "%REPO_DIR%\worker.log" 2>&1

