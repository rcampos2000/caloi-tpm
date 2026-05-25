@echo off
chcp 65001 > nul
title Sistema de Manutenção TPM - Caloi

echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║     SISTEMA DE CONTROLE DE MANUTENÇÃO - CALOI       ║
echo  ║                    TPM v2.0                          ║
echo  ╚══════════════════════════════════════════════════════╝
echo.

:: Verificar se Python está instalado
python --version > nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo  [ERRO] Python não encontrado!
    echo.
    echo  Por favor, instale o Python em: https://www.python.org/downloads/
    echo  Marque a opção "Add Python to PATH" durante a instalação.
    echo.
    pause
    exit /b 1
)

echo  [OK] Python encontrado.
echo.

:: Instalar dependências
echo  Instalando dependências necessárias...
pip install flask openpyxl --quiet --no-warn-script-location
IF %ERRORLEVEL% NEQ 0 (
    echo  [AVISO] Tentando instalar com pip3...
    pip3 install flask openpyxl --quiet
)
echo  [OK] Dependências instaladas.
echo.

:: Criar pasta do banco de dados no Desktop
SET DB_FOLDER=%USERPROFILE%\Desktop\Banco de dados Manutenção
IF NOT EXIST "%DB_FOLDER%" (
    mkdir "%DB_FOLDER%"
    mkdir "%DB_FOLDER%\assinaturas"
    mkdir "%DB_FOLDER%\backups"
    echo  [OK] Pasta criada: %DB_FOLDER%
) ELSE (
    echo  [OK] Pasta existente: %DB_FOLDER%
)
echo.

:: Iniciar o servidor
echo  ┌──────────────────────────────────────────────────────┐
echo  │  Iniciando servidor... Aguarde alguns segundos       │
echo  │                                                      │
echo  │  Acesse no navegador:  http://localhost:5000         │
echo  │  Dashboard:            http://localhost:5000/dashboard│
echo  │                                                      │
echo  │  Para ENCERRAR: feche esta janela ou pressione       │
echo  │  CTRL+C nesta janela                                 │
echo  └──────────────────────────────────────────────────────┘
echo.

:: Aguardar 2 segundos e abrir o navegador automaticamente
timeout /t 2 /nobreak > nul
start "" "http://localhost:5000"

:: Executar o app
python app.py

echo.
echo  [!] Servidor encerrado.
pause
