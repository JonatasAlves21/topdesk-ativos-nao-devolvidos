@echo off
REM Execucao manual: segredos injetados pelo 1Password CLI.
cd /d "%~dp0"
op run --env-file=".env" -- node export.js
pause
