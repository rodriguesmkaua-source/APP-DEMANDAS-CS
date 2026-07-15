@echo off
cd /d "%~dp0"
title Demandas CS - Modo dev

echo.
echo  ==========================================
echo   Demandas CS - Iniciando em modo dev
echo  ==========================================
echo.

node --version >nul 2>&1
if errorlevel 1 (
    echo  ERRO: Node.js nao encontrado.
    echo  Instale em: https://nodejs.org
    pause & exit /b 1
)

if not exist node_modules (
    echo  Instalando dependencias pela primeira vez...
    npm install
    echo.
)

echo  Abrindo app...
npm start
