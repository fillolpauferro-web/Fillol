@echo off
cd /d "%~dp0"
set PY=python
where python >nul 2>nul || set PY=py
%PY% -m pip install -q --user pandas pyarrow openpyxl python-calamine
%PY% ressarcimento_completo.py %*
pause
