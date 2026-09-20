@echo off
cd /d "%~dp0"
set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python312\python.exe"
if not exist "%PYTHON_EXE%" (
    echo 未找到带 tkinter 的 Python 3.12：%PYTHON_EXE%
    echo 请先安装 Python 3.12 官方 Windows 版本，并勾选 Tcl/Tk。
    pause
    exit /b 1
)
"%PYTHON_EXE%" can_matrix_checker.py
if errorlevel 1 pause
