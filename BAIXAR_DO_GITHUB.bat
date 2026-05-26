@echo off
chcp 65001 > nul
title CALOI TPM — Baixar do GitHub

echo.
echo ╔══════════════════════════════════════════════════════════╗
echo ║     CALOI TPM — BAIXAR VERSAO MAIS RECENTE DO GITHUB    ║
echo ║     Sincroniza os arquivos locais com a nuvem            ║
echo ╚══════════════════════════════════════════════════════════╝
echo.

python --version > nul 2>&1
if errorlevel 1 (
    python3 "%~dp0BAIXAR_DO_GITHUB.py"
) else (
    python "%~dp0BAIXAR_DO_GITHUB.py"
)
