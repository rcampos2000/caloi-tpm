# Sistema de Manutenção TPM - Caloi
# Script de inicialização PowerShell

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║     SISTEMA DE CONTROLE DE MANUTENÇÃO - CALOI       ║" -ForegroundColor Cyan
Write-Host "  ║                    TPM v2.0                          ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Mudar para a pasta do script
Set-Location -Path $PSScriptRoot

# Instalar dependências
Write-Host "  Instalando dependências..." -ForegroundColor Yellow
python -m pip install flask openpyxl --quiet 2>&1 | Out-Null
Write-Host "  [OK] Dependências prontas." -ForegroundColor Green
Write-Host ""

# Criar pasta de banco de dados se não existir
$dbFolder = "$env:USERPROFILE\Desktop\Banco de dados Manutenção"
if (-not (Test-Path $dbFolder)) {
    New-Item -ItemType Directory -Path $dbFolder | Out-Null
    New-Item -ItemType Directory -Path "$dbFolder\assinaturas" | Out-Null
    New-Item -ItemType Directory -Path "$dbFolder\backups" | Out-Null
    Write-Host "  [OK] Pasta criada: $dbFolder" -ForegroundColor Green
}

Write-Host "  ┌──────────────────────────────────────────────────────┐" -ForegroundColor Cyan
Write-Host "  │  Servidor iniciando... Aguarde alguns segundos       │" -ForegroundColor Cyan
Write-Host "  │                                                      │" -ForegroundColor Cyan
Write-Host "  │  Acesse: http://localhost:5000                       │" -ForegroundColor Cyan
Write-Host "  │                                                      │" -ForegroundColor Cyan
Write-Host "  │  Para ENCERRAR: feche esta janela (Ctrl+C)           │" -ForegroundColor Cyan
Write-Host "  └──────────────────────────────────────────────────────┘" -ForegroundColor Cyan
Write-Host ""

# Abrir navegador após 3 segundos em background
Start-Job -ScriptBlock {
    Start-Sleep -Seconds 3
    Start-Process "http://localhost:5000"
} | Out-Null

# Iniciar Flask
python app.py
