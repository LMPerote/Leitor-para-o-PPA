@echo off
rem Abre a janela do Leitor para o PPA com dois cliques.
cd /d "%~dp0"

py --version > NUL 2>&1
if errorlevel 1 goto sem_python

py -m pip show python-docx > NUL 2>&1
if errorlevel 1 (
  echo Instalando as dependencias pela primeira vez, aguarde...
  py -m pip install -r requirements.txt
  if errorlevel 1 goto erro_dependencias
)

rem Confere se o aplicativo carrega ANTES de abrir a janela. Sem isso, uma
rem falha no carregamento faria o programa simplesmente nao abrir, sem aviso.
py -c "import interface"
if errorlevel 1 goto erro_carregar

rem pyw.exe abre a janela sem o console preto atras; py.exe e o plano B.
where pyw > NUL 2>&1
if errorlevel 1 (
  start "" py "interface.py"
) else (
  start "" pyw "interface.py"
)
exit /b 0

:sem_python
echo.
echo  Python nao foi encontrado neste computador.
echo  Instale em https://www.python.org/downloads/ marcando a opcao
echo  "Add python.exe to PATH" e execute este arquivo novamente.
echo.
pause
exit /b 1

:erro_dependencias
echo.
echo  Nao foi possivel instalar as dependencias. A mensagem acima diz o motivo.
echo.
pause
exit /b 1

:erro_carregar
echo.
echo  O aplicativo nao pode ser carregado - veja a mensagem acima.
echo  Se ela citar algum modulo faltando, rode neste prompt:
echo      py -m pip install -r requirements.txt
echo.
pause
exit /b 1
