@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python nao encontrado no PATH. Instale o Python marcando "Add python.exe to PATH".
  pause
  exit /b 1
)

python -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Instalando o pacote openpyxl...
  python -m pip install openpyxl
)

python atualizar_aderencia.py %*
set RC=%errorlevel%

echo.
echo Codigo de saida: %RC%  (0 = OK)
echo Log completo em: %~dp0logs\aderencia.log
pause
exit /b %RC%
