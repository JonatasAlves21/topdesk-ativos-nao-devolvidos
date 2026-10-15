@echo off
REM Execucao agendada (Agendador de Tarefas): acrescenta a saida em exportar.log.
cd /d "%~dp0"
op run --env-file=".env" -- node export.js >> exportar.log 2>&1
