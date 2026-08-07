@echo off
setlocal
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
  echo Python nao encontrado. Instale em https://www.python.org/downloads/ e tente de novo.
  pause
  exit /b 1
)

echo Instalando/atualizando dependencias...
python -m pip install --quiet --disable-pip-version-check --user -r requirements.txt

set "TARGET=%~1"
if "%TARGET%"=="" (
  rem Sem pasta arrastada: usa a subpasta "xml" se existir, senao a propria pasta do .bat
  if exist "%~dp0xml\" (
    set "TARGET=%~dp0xml"
  ) else (
    set "TARGET=%~dp0"
  )
)

rem Remove a barra final (se houver) para nao quebrar as aspas no Windows
if "%TARGET:~-1%"=="\" set "TARGET=%TARGET:~0,-1%"

echo Gerando relatorio de notas de servico (NFS-e)...
python nfse_report.py "%TARGET%"
pause
