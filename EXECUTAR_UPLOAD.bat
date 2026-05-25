@echo off
echo ============================================
echo    CALOI TPM - Enviando arquivos ao GitHub
echo    (via Python - sem necessidade de git)
echo ============================================
echo.
python "%~dp0UPLOAD_GITHUB.py"
if errorlevel 1 (
    echo.
    echo ERRO: Python nao encontrado. Tentando python3...
    python3 "%~dp0UPLOAD_GITHUB.py"
)
pause
