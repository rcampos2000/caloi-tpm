@echo off
title Sistema de Manutencao TPM - Caloi
cd /d "%~dp0"

echo.
echo  ============================================================
echo   SISTEMA DE CONTROLE DE MANUTENCAO - CALOI TPM v2.0
echo  ============================================================
echo.

echo  Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERRO] Python nao encontrado!
    echo  Baixe em: https://www.python.org/downloads/
    echo  Marque "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)

echo  [OK] Python encontrado.
echo.
echo  Instalando dependencias...
python -m pip install flask openpyxl --quiet
echo  [OK] Dependencias prontas.
echo.

echo  Iniciando servidor Flask...
echo  Acesse: http://localhost:5000
echo  Para encerrar: feche esta janela (Ctrl+C)
echo.

start "" /b cmd /c "timeout /t 3 >nul && start http://localhost:5000"
python app.py

pause
