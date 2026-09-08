@echo off
title xPence Report Generator

:: ── Check Python ────────────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo Python not found. Install from https://www.python.org
    pause
    exit /b 1
)

:: ── Install / verify dependencies ───────────────────────────────────────────
python -c "import pandas" >nul 2>&1    || pip install pandas
python -c "import openpyxl" >nul 2>&1 || pip install openpyxl

:: ── Launch GUI ───────────────────────────────────────────────────────────────
python "%~dp0xpence_gui.py"
