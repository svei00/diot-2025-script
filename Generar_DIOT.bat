@echo off
chcp 65001 > nul
REM ===== Generador DIOT (lee el Registro, escribe solo el .txt) =====
REM Sin argumentos abre la GUI: eliges el libro .xlsb, el mes y donde guardar.
cd /d "%~dp0"
python diot_generator.py
