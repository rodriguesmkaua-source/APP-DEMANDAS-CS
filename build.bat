@echo off
cd /d "%~dp0"
title Demandas CS - Build

echo.
echo  ==========================================
echo   Demandas CS - Gerando pacote
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

:: ── 1. Compilar ──────────────────────────────────────────────────────────
echo  [1/3] Compilando app...
npx electron-builder --dir
if errorlevel 1 ( echo  ERRO na compilacao. & pause & exit /b 1 )

:: ── 2. Adicionar instalador ao pacote e zipar ─────────────────────────────
echo  [2/3] Criando ZIP para distribuicao...
copy /Y "%~dp0instalar.bat" "dist\win-unpacked\instalar.bat" >nul
if exist "dist\Demandas CS.zip" del "dist\Demandas CS.zip"
powershell -NoProfile -Command "Compress-Archive -Path 'dist\win-unpacked\*' -DestinationPath 'dist\Demandas CS.zip' -Force"

:: ── 3. Instalar localmente e criar atalho na area de trabalho ────────────
echo  [3/3] Instalando no computador e criando atalho...
set "DEST=%LOCALAPPDATA%\Demandas CS"
xcopy /E /I /Y "dist\win-unpacked" "%DEST%\" >nul
powershell -NoProfile -Command ^
  "$s=(New-Object -COM WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\Demandas CS.lnk');" ^
  "$s.TargetPath='%DEST%\Demandas CS.exe';" ^
  "$s.WorkingDirectory='%DEST%';" ^
  "$s.Description='KnowU CS - Log de Demandas';" ^
  "$s.Save()"

echo.
echo  ==========================================
echo   Pronto!
echo.
echo   Voce:
echo   -> Atalho criado na area de trabalho
echo.
echo   Colegas:
echo   -> Envie o arquivo: dist\Demandas CS.zip
echo   -> Eles extraem e clicam em instalar.bat
echo  ==========================================
echo.
pause
