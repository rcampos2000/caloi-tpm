@echo off
echo ============================================
echo    CALOI TPM - Enviando para o GitHub...
echo ============================================
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0DEPLOY_GITHUB.ps1"
pause
