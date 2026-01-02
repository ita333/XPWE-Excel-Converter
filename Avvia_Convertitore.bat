@echo off
title Avvio Convertitore XPWE
cd /d "%~dp0"
echo Avvio del convertitore in corso...
start /b pythonw.exe xpwe_converter_app.py
exit
