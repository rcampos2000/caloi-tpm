@echo off
chcp 65001 >nul
title Caloi TPM - Copiar para C:\Projetos (fora do OneDrive)
cd /d "%~dp0"

set "DEST=C:\Projetos\Manutencao_TPM"

echo ============================================
echo   Copiando o Caloi TPM para uma pasta limpa
echo   Origem : %~dp0
echo   Destino: %DEST%
echo ============================================
echo.
echo (sem OneDrive e sem o .git da pasta de usuario - o git funciona certo la)
echo Os DADOS (xlsx, config) sao copiados para o TPM continuar funcionando,
echo mas o .gitignore impede que eles subam ao GitHub.
echo.
pause

if not exist "C:\Projetos" mkdir "C:\Projetos"

rem Copia tudo, menos cache, .git antigo, temporarios e o token
robocopy "%~dp0." "%DEST%" /E ^
  /XD __pycache__ .git ^
  /XF *.pyc *.pyo *.log _temp_*.xlsx github_token.txt UPLOAD_GITHUB.py

echo.
echo ============================================
echo   Pronto! Agora:
echo   1) Abra a pasta %DEST%
echo   2) De dois cliques em SUBIR_GITHUB.bat (de la)
echo   3) O Railway, se estiver ligado ao repo, faz o redeploy sozinho
echo ============================================
echo.
echo Dica: deixe o SGM em C:\Projetos\SGM_Caloi (lado a lado) -
echo assim o SGM acha os dados do TPM automaticamente.
pause
