@echo off
setlocal

set "PASTA_PROJETO=C:\Users\jaildo.junior\Desktop\ANALISE_PCP_RFK"
set "SCRIPT=%PASTA_PROJETO%\monitorar_index.py"

cd /d "%PASTA_PROJETO%"

REM Deixe esta janela aberta: ela fica de olho no index.html e envia
REM automaticamente para o GitHub sempre que o arquivo for alterado.
py "%SCRIPT%"

pause
endlocal
