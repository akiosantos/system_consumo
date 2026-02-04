@echo off
chcp 65001 >nul
title Sistema Sabesp + Enel

echo =========================================
echo   Iniciando Sistema Sabesp + Enel
echo =========================================

set PASTA=U:\BackupContabilidade\Custos\0 - Enel, Sabesp e Telefônica - Lucas\system\backend

if not exist "%PASTA%" (
    echo ERRO: Pasta nao encontrada:
    echo %PASTA%
    pause
    exit
)

cd /d "%PASTA%"

python -m uvicorn main:app --host 0.0.0.0 --port 8000

pause
