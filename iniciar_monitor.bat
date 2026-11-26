@echo off
chcp 65001 >nul
title Monitor de Escala Medica - Hospital Cruzeiro do Sul (Hapvida)
cd /d "%~dp0"
echo Iniciando monitoramento automatico da planilha OneDrive...
echo Sempre que a planilha for salva, o index.html sera atualizado e enviado ao GitHub Pages.
echo.
python atualizar_escala.py --watch --git-push
pause
