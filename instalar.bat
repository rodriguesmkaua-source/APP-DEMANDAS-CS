@echo off
cd /d "%~dp0"
title Demandas CS - Instalador

echo.
echo  ==========================================
echo   Demandas CS - Instalacao
echo  ==========================================
echo.

set "DEST=%LOCALAPPDATA%\Demandas CS"

echo  Copiando arquivos para:
echo  %DEST%
echo.

xcopy /E /I /Y "%~dp0" "%DEST%\" >nul

echo  Criando atalho na area de trabalho...
powershell -NoProfile -Command ^
  "$s=(New-Object -COM WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\Demandas CS.lnk');" ^
  "$s.TargetPath='%DEST%\Demandas CS.exe';" ^
  "$s.WorkingDirectory='%DEST%';" ^
  "$s.Description='KnowU CS - Log de Demandas';" ^
  "$s.Save()"

echo.
echo  ==========================================
echo   Pronto!
echo   Atalho criado na area de trabalho.
echo   Pode fechar esta janela.
echo  ==========================================
echo.
pause
