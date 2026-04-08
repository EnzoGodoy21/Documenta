@echo off
python interface.py
if errorlevel 1 (
    echo.
    echo Erro ao executar. Verifique se o Python esta instalado.
    pause
)
