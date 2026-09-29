@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  Disparador de Emails - Leilao
echo  ==============================
echo  Verificando dependencias...
pip install -r requirements.txt -q
echo  Iniciando servidor...
echo  Acesse: http://localhost:5000
echo.
python app.py
pause
