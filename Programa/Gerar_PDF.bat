@echo off
setlocal
title Relatorio do Gestor - PDF
if not exist "%~dp0gerar_relatorio.py" goto incompleto
set "PYTHONDONTWRITEBYTECODE=1"
set "PYTHONIOENCODING=utf-8"
chcp 65001 >nul
py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
if not errorlevel 1 goto pelo_py
python -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
if not errorlevel 1 goto pelo_python
echo Python 3.10 ou mais recente nao foi encontrado.
echo Use uma instalacao de Python autorizada no seu computador.
echo Este programa nao instala nada nem solicita acesso de administrador.
pause
exit /b 1
:pelo_py
py -3 "%~dp0gerar_relatorio.py"
exit /b %errorlevel%
:pelo_python
python "%~dp0gerar_relatorio.py"
exit /b %errorlevel%
:incompleto
echo Extraia todos os arquivos do ZIP para a mesma pasta antes de iniciar.
pause
exit /b 1
