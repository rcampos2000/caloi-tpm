@echo off
chcp 65001 >nul
title Caloi TPM - Enviar para o GitHub
cd /d "%~dp0"

echo ============================================
echo   Caloi TPM  -  Envio para o GitHub
echo ============================================
echo.

where git >nul 2>&1
if errorlevel 1 (
  echo [ERRO] Git nao encontrado. Instale em https://git-scm.com/download/win
  pause & exit /b 1
)

rem 1) Inicializa o repositorio (apenas na 1a vez)
if not exist ".git" (
  echo Inicializando repositorio git...
  git init
  git branch -M main
)

rem 2) TRAVA DE SEGURANCA: o repo precisa ser ESTA pasta (e nao a pasta de usuario)
if not exist ".git" (
  echo [ERRO] Nao foi possivel criar o repositorio git aqui.
  echo Provavel bloqueio do OneDrive. Rode MOVER_PARA_PROJETOS.bat e tente de la.
  pause & exit /b 1
)
for /f "delims=" %%R in ('git rev-parse --show-toplevel 2^>nul') do set "TOPDIR=%%R"
set "TOPDIR=%TOPDIR:/=\%"
if /i not "%TOPDIR%"=="%CD%" (
  echo [ERRO] O git esta apontando para outra pasta:
  echo    %TOPDIR%
  echo Existe um .git em C:\Users\%USERNAME%. Rode MOVER_PARA_PROJETOS.bat
  echo e use C:\Projetos\Manutencao_TPM.
  pause & exit /b 1
)

rem 3) Define o repositorio remoto (apenas na 1a vez)
git remote get-url origin >nul 2>&1
if errorlevel 1 (
  echo.
  echo Cole a URL do repositorio GitHub do TPM
  echo  exemplo: https://github.com/rcampos2000/caloi-tpm.git
  set /p REPO=URL:
  git remote add origin "%REPO%"
)

set MSG=Atualizacao TPM %date% %time%

echo.
echo Adicionando arquivos (dados e token sao ignorados pelo .gitignore)...
git add -A
git commit -m "%MSG%"

echo.
echo Enviando para o GitHub (pode abrir o navegador para login)...
git push -u origin main

echo.
echo ============================================
echo   Concluido. Se o Railway estiver ligado ao
echo   repo, o redeploy comeca automaticamente.
echo ============================================
pause
