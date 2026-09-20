@echo off
cd /d "%~dp0"
python can_matrix_checker.py
if errorlevel 1 pause
