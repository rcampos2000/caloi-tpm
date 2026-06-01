# ============================================================
#  DEPLOY_GITHUB.ps1 — Caloi TPM v2.0
#  Envia o projeto para o GitHub (execute 1x)
# ============================================================

$ErrorActionPreference = "Stop"
$pasta = Split-Path -Parent $MyInvocation.MyCommand.Path

# URL do repositório já configurada
$repoUrl = "https://github.com/rcampos2000/caloi-tpm.git"

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   CALOI TPM — Deploy para o GitHub" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Repositório: $repoUrl" -ForegroundColor Gray
Write-Host ""

# ── 1. Verificar se git está instalado ──────────────────────
try {
    $gitVer = git --version 2>&1
    Write-Host "OK  Git encontrado: $gitVer" -ForegroundColor Green
} catch {
    Write-Host "ERRO  Git nao esta instalado!" -ForegroundColor Red
    Write-Host "   Baixe em: https://git-scm.com/download/win" -ForegroundColor Yellow
    Write-Host "   Apos instalar, execute este script novamente." -ForegroundColor Yellow
    pause
    exit 1
}

# ── 2. Navegar para a pasta do projeto ──────────────────────
Set-Location $pasta
Write-Host "Pasta do projeto: $pasta" -ForegroundColor Gray
Write-Host ""

# ── 3. Remover .git quebrado se existir e re-inicializar ────
if (Test-Path ".git") {
    Write-Host "Removendo .git antigo..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force ".git"
}

Write-Host "Inicializando repositorio git..." -ForegroundColor Cyan
git init
git branch -m main
git config user.email "campos_ri@hotmail.com"
git config user.name "Ricardo Campos"

# ── 4. Adicionar arquivos e fazer commit ─────────────────────
Write-Host ""
Write-Host "Adicionando arquivos..." -ForegroundColor Cyan
git add .

Write-Host "Criando commit inicial..." -ForegroundColor Cyan
git commit -m "feat: Caloi TPM v2.0 - sistema de manutencao industrial"

# ── 5. Configurar remote e fazer push ───────────────────────
Write-Host ""
Write-Host "Configurando remote origin..." -ForegroundColor Cyan
git remote add origin $repoUrl

Write-Host "Enviando para o GitHub..." -ForegroundColor Cyan
Write-Host "(Pode aparecer janela de autenticacao do Git)" -ForegroundColor Yellow
git push -u origin main

# ── 6. Confirmar ─────────────────────────────────────────────
Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "   SUCESSO! Projeto enviado ao GitHub." -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Proximo passo: configure o Railway em railway.app" -ForegroundColor Cyan
Write-Host "  1. New Project > Deploy from GitHub repo" -ForegroundColor Gray
Write-Host "  2. Selecione: caloi-tpm" -ForegroundColor Gray
Write-Host "  3. Adicione Volume em /data" -ForegroundColor Gray
Write-Host "  4. Variaveis: DATA_DIR=/data  SECRET_KEY=<gere-uma-chave>" -ForegroundColor Gray
Write-Host "  5. Gere o dominio publico em Settings > Networking" -ForegroundColor Gray
Write-Host ""
pause
