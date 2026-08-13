@echo off
rem Abre a janela do Leitor para o PPA com dois cliques.
cd /d "%~dp0"

py --version > NUL 2>&1
if errorlevel 1 goto sem_python

py -m pip show python-docx > NUL 2>&1
if errorlevel 1 (
  echo Instalando as dependencias pela primeira vez, aguarde...
  py -m pip install -r requirements.txt
)

start "" py -w "interface.py"
exit /b 0

:sem_python
echo.
echo  Python nao foi encontrado neste computador.
echo  Instale em https://www.python.org/downloads/ marcando a opcao
echo  "Add python.exe to PATH" e execute este arquivo novamente.
echo.
pause
exit /b 1
