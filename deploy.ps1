# deploy.ps1 — Empacota e instala o app
# Uso: powershell -ExecutionPolicy Bypass -File deploy.ps1

$SRC  = $PSScriptRoot
$TMP  = "C:\Users\$env:USERNAME\AppData\Local\Temp\cs-deploy"
$DEST = "$env:LOCALAPPDATA\Demandas CS\resources"

Write-Host "=== Deploy KnowU CS ===" -ForegroundColor Cyan

# 1. Limpa temp
if (Test-Path $TMP) { Remove-Item $TMP -Recurse -Force }
New-Item -ItemType Directory $TMP | Out-Null

# 2. Copia projeto excluindo pastas que NÃO devem entrar no app
Write-Host "Copiando projeto..." -ForegroundColor Yellow
$excludeDirs = @('dist','app.asar.unpacked','.git','node_modules')
Get-ChildItem $SRC -Exclude $excludeDirs | ForEach-Object {
  Copy-Item $_.FullName "$TMP\$($_.Name)" -Recurse -Force
}

# 3. Copia node_modules excluindo pacotes de desenvolvimento pesados
Write-Host "Copiando dependencias..." -ForegroundColor Yellow
New-Item -ItemType Directory "$TMP\node_modules" | Out-Null
$devOnlyPkgs = @('electron','app-builder-bin','app-builder-lib','electron-builder',
                  'electron-publish','7zip-bin','winCodeSign','typescript',
                  'dmg-builder','builder-util','builder-util-runtime','@types')
Get-ChildItem "$SRC\node_modules" | Where-Object {
  $devOnlyPkgs -notcontains $_.Name
} | ForEach-Object {
  Copy-Item $_.FullName "$TMP\node_modules\$($_.Name)" -Recurse -Force
}

$srcSize = [math]::Round((Get-ChildItem $TMP -Recurse -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum/1MB)
Write-Host "Tamanho da fonte: ${srcSize}MB" -ForegroundColor Gray

# 4. Empacota de dentro do TMP para evitar recursão
Write-Host "Empacotando ASAR..." -ForegroundColor Yellow
$outAsar = "$TMP\..\cs-app.asar"
Set-Location $SRC
npx asar pack $TMP $outAsar 2>&1
if ($LASTEXITCODE -ne 0) { Write-Host "ERRO no pack!" -ForegroundColor Red; Set-Location $SRC; exit 1 }

$asarSize = [math]::Round((Get-Item $outAsar).Length/1MB, 1)
Write-Host "ASAR gerado: ${asarSize}MB" -ForegroundColor Green

# 5. Remove app.asar.unpacked do destino (não mais necessário)
if (Test-Path "$DEST\app.asar.unpacked") {
  Remove-Item "$DEST\app.asar.unpacked" -Recurse -Force
}

# 6. Instala
Write-Host "Instalando..." -ForegroundColor Yellow
Copy-Item -Force $outAsar "$DEST\app.asar"

# 7. Limpa temp
Remove-Item $TMP -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $outAsar -Force -ErrorAction SilentlyContinue

Set-Location $SRC
Write-Host "=== Pronto! App instalado com ${asarSize}MB ===" -ForegroundColor Green
Write-Host "Reinicie 'Demandas CS' para ver as mudancas." -ForegroundColor Cyan
