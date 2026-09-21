@echo off
setlocal

set "PASTA_PROJETO=C:\Users\jaildo.junior\Desktop\ANALISE_PCP_RFK"
set "SCRIPT=%PASTA_PROJETO%\atualizar_dashboard.py"
set "LOG_BAT=%PASTA_PROJETO%\logs\bat_%date:~-4,4%-%date:~-7,2%-%date:~-10,2%.log"

REM Garante que a pasta de logs existe antes de escrever nela
if not exist "%PASTA_PROJETO%\logs" mkdir "%PASTA_PROJETO%\logs"

echo ===== Execucao do .bat iniciada em %date% %time% ===== >> "%LOG_BAT%"

cd /d "%PASTA_PROJETO%"

echo Rodando: py "%SCRIPT%" >> "%LOG_BAT%"
py "%SCRIPT%" >> "%LOG_BAT%" 2>&1

echo Codigo de saida do Python: %errorlevel% >> "%LOG_BAT%"
echo ===== Execucao do .bat finalizada em %date% %time% ===== >> "%LOG_BAT%"

endlocal
