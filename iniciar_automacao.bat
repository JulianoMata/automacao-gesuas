@echo off
chcp 65001 > nul
title Sistema de Automação de Atendimentos - GESUAS / CRAS Sul
echo =========================================================================
echo  INICIANDO SISTEMA DE AUTOMACAO OPERACIONAL - GESUAS
echo  CRAS Sul - Campo Belo / MG
echo =========================================================================
echo.

call "%~dp0.venv\Scripts\activate.bat"
python "%~dp0SCRIPTS\main.py"

echo.
echo =========================================================================
echo  Execução encerrada. Pressione qualquer tecla para fechar...
echo =========================================================================
pause > nul
