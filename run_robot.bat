@echo off
:: Navega para o diretório onde este arquivo .bat está localizado
cd /d "%~dp0"

:: Ativa o ambiente virtual (agora o caminho relativo é seguro)
echo Ativando ambiente virtual...
call "venv\Scripts\activate.bat"

:: Executa o script Python
echo Executando o script Python...
python main.py

:: Desativa o ambiente virtual
echo Desativando ambiente virtual...
call "venv\Scripts\deactivate.bat"

echo.
echo Processo concluido.