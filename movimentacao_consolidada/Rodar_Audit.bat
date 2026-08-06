@echo off
setlocal

REM ============================================================
REM Rodar_Audit.bat
REM Duplo clique roda so o --audit: lista os Document Type do
REM export e mostra se algum esta sem categoria em
REM category_rules.csv. Nao grava nada.
REM ============================================================

set RAW_DIR=C:\Users\I0507867.FARMA\OneDrive - Sanofi\Desktop\OL Robo

REM --- localizar o Python (tenta o do Spyder primeiro, depois o do PATH) ---
set SPYDER_PY=C:\Users\I0507867.FARMA\AppData\Local\Programs\spyder-6\envs\spyder-runtime\python.exe

if exist "%SPYDER_PY%" (
    set "PYTHON_EXE=%SPYDER_PY%"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_EXE=python"
    ) else (
        where py >nul 2>nul
        if %errorlevel%==0 (
            set "PYTHON_EXE=py"
        ) else (
            echo Nao encontrei o Python nesta maquina.
            echo Abra o IPython que voce ja usa e rode:
            echo     import sys; print(sys.executable)
            echo Copie o caminho que aparecer e cole na linha SPYDER_PY deste .bat.
            pause
            exit /b 1
        )
    )
)

cd /d "%~dp0"

"%PYTHON_EXE%" etl.py --raw-dir "%RAW_DIR%" --audit

echo.
pause
