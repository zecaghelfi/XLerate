# Instala a versao do XLerate deste fork (zecaghelfi/XLerate) no Excel desktop.
# Remove o registro do XLerate original e preserva os outros suplementos (OF Asset etc.).
$ErrorActionPreference = 'Stop'
$key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$idOriginal = 'e3afbfd8-06fa-47f9-a862-bdc61ed6e9bf'

$dir = Join-Path $env:LOCALAPPDATA 'XLerate'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$manifest = Join-Path $dir 'manifest-of.xml'
Invoke-WebRequest -UseBasicParsing -Uri 'https://zecaghelfi.github.io/XLerate/manifest.xml' -OutFile $manifest
[xml]$x = Get-Content -Raw $manifest
$id = $x.OfficeApp.Id
if ($x.OfficeApp.DefaultSettings.SourceLocation.DefaultValue -notlike 'https://zecaghelfi.github.io/*') {
  throw "O site ainda serve o manifesto antigo. Espere o deploy do GitHub Pages terminar e rode de novo."
}

if (-not (Test-Path $key)) { New-Item -Path $key | Out-Null }
Remove-ItemProperty -Path $key -Name $idOriginal -ErrorAction SilentlyContinue
New-ItemProperty -Path $key -Name $id -Value $manifest -PropertyType String -Force | Out-Null

# Limpa o cache de suplementos do Office (atalhos e painel antigos ficam guardados ali)
$cache = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
if (Test-Path $cache) { Remove-Item -Recurse -Force $cache -ErrorAction SilentlyContinue }

Write-Host ""
Write-Host "OK: XLerate (fork) registrado, Id $id."
Write-Host "Suplementos de desenvolvedor registrados:"
Get-ItemProperty -Path $key | Select-Object * -ExcludeProperty PS* | Format-List
Write-Host "Feche o Excel por completo e abra de novo: Inserir > Meus Suplementos > Suplementos do Desenvolvedor > XLerate."
