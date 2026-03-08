@echo off
cd /d "%~dp0"
echo Avvio Ollama...
start "" "C:\Users\Alessandro\Desktop\mailhelper\ollama-windows-amd64\ollama.exe" serve
timeout /t 3 /nobreak >nul
echo Avvio Estrattore DDT...
python estrai_ddt.py
