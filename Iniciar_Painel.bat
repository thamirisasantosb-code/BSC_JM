@echo off
echo Iniciando Servidor do Painel Gerot...
start cmd /k "node server.js"
timeout /t 2 /nobreak >nul
echo Abrindo navegador...
start http://localhost:3000
exit
