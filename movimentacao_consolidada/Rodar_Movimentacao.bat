@echo off
setlocal

REM ============================================================
REM Rodar_Movimentacao.bat
REM Duplo clique le TODOS os arquivos que estiverem em RAW_DIR
REM (Movimentacao + Saldo Inicial) e grava o resultado dentro da
REM propria pasta, em RAW_DIR\output\consolidado.xlsx -- tudo num
REM lugar so, sem depender da pasta do projeto.
REM
REM Edite as duas linhas abaixo se a pasta do export ou o mes de
REM ancoragem do Saldo Inicial mudar.
REM ============================================================

set RAW_DIR=C:\Users\I0507867.FARMA\OneDrive - Sanofi\Desktop\OL Robo
set OUTPUT_DIR=%RAW_DIR%\output
set SALDO_INICIAL_DATA=2026-05-01

REM --- localizar o Python (tenta o do Spyder primeiro, depois o do PATH) ---
set SPYDER_PY=C:\Users\I0507867.FARMA\AppData\Local\spyder-6\envs\spyder-runtime\python.exe

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

"%PYTHON_EXE%" etl.py --raw-dir "%RAW_DIR%" --output-dir "%OUTPUT_DIR%" --saldo-inicial-data %SALDO_INICIAL_DATA%

echo.
echo ============================================================
echo Terminado. Confira o resultado em:
echo %OUTPUT_DIR%\consolidado.xlsx
echo ============================================================
pause
