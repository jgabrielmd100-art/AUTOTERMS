@echo off
TITLE AutoTermos - Iniciando Sistema...

echo ==========================================
echo    🚀 AUTO-TERMOS v1 - SIMPLES CONTABIL
echo ==========================================
echo.

:: Verifica se a pasta node_modules existe
if not exist "node_modules" (
    echo [!] Pasta node_modules nao encontrada.
    echo [!] Instalando dependencias...
    call npm install
)

echo [+] Iniciando Servidor e Frontend...
echo [+] O sistema abrira automaticamente no seu navegador.
echo.

:: Inicia o processo de desenvolvimento
start "" http://localhost:3000
call npm run dev:all

pause
