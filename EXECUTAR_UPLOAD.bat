@echo off
chcp 65001 > nul
title CALOI TPM — Upload para GitHub

echo.
echo ╔══════════════════════════════════════════════════════════╗
echo ║       CALOI TPM — ENVIAR ATUALIZACOES PARA NUVEM        ║
echo ║       GitHub → Railway (deploy automatico)               ║
echo ╚══════════════════════════════════════════════════════════╝
echo.

:: Verificar se github_token.txt existe
if not exist "%~dp0github_token.txt" (
    echo ERRO: arquivo 'github_token.txt' nao encontrado!
    echo.
    echo    O arquivo com o token do GitHub e necessario para o upload.
    echo    Contate o administrador do sistema.
    echo.
    pause
    exit /b 1
)

:: Tentar Python
python --version > nul 2>&1
if errorlevel 1 (
    python3 "%~dp0UPLOAD_GITHUB.py"
) else (
    python "%~dp0UPLOAD_GITHUB.py"
)
pause
