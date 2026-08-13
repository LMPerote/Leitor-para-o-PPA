@echo off
rem Gera o LeitorPPA.exe nesta maquina (pasta dist\).
cd /d "%~dp0"

py --version > NUL 2>&1
if errorlevel 1 (
  echo Python nao encontrado. Instale em https://www.python.org/downloads/
  pause
  exit /b 1
)

echo Instalando o PyInstaller e as dependencias...
py -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto erro

echo.
echo Gerando o executavel (pode levar alguns minutos)...
py -m PyInstaller --onefile --windowed --noconfirm --name LeitorPPA --paths . interface.py
if errorlevel 1 goto erro

echo.
echo Pronto: dist\LeitorPPA.exe
explorer dist
pause
exit /b 0

:erro
echo.
echo Falhou. Copie a mensagem acima para analisar.
pause
exit /b 1
