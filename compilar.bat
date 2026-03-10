@echo off

set ICON_PARAM=
if exist "doc\icone.ico" (
    echo [INFO] Icone personalizado encontrado.
    set ICON_PARAM=--icon="doc\icone.ico"
) else (
    echo [AVISO] Icone "doc\icone.ico" nao encontrado.
    echo O programa sera compilado com o icone padrao do Python.
)

echo Encerrando processos antigos (se existirem)...
taskkill /f /im "Configurador.exe" >nul 2>&1
taskkill /f /im "Automatizador de Combustivel.exe" >nul 2>&1

echo Limpando arquivos antigos...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /q *.spec

echo.
echo Compilando Configurador...
pyinstaller --noconsole --onefile --clean %ICON_PARAM% --name="Configurador" configurador.py

echo.
echo Compilando Robo Combustivel...
pyinstaller --noconsole --onefile --clean %ICON_PARAM% --hidden-import plyer.platforms.win.notification --name="Automatizador de Combustivel" combustivel.pyw

echo.
echo Copiando pasta 'doc' para 'dist'...
if not exist "dist\doc" mkdir "dist\doc"
if exist "doc\icone.ico" (
    copy "doc\icone.ico" "dist\doc\" >nul
)

echo.
echo -------------------------------------------------------
echo SUCESSO! Os arquivos .exe foram criados na pasta 'dist'.
echo -------------------------------------------------------
pause