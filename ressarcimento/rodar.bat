@echo off
cd /d "%~dp0"
pip install -q pandas pyarrow openpyxl python-calamine
python converter_e_analisar.py %*
pause
