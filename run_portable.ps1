# Portable runner for Monatsbericht_Bonus_Quartal.py without admin rights
# Steps:
# 1) Downloads Python embeddable (Windows, 64-bit) to ./py-embed
# 2) Enables site-packages and bootstraps pip locally
# 3) Installs required packages into ./py-embed
# 4) Runs the script using the embedded Python

param(
  [string]$CsvPath,
  [string]$XmlPath,
  [string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Cyan }
function Write-Ok($msg)   { Write-Host "[OK]   $msg" -ForegroundColor Green }
function Write-Warn($msg) { Write-Host "[WARN] $msg" -ForegroundColor Yellow }
function Write-Err($msg)  { Write-Host "[ERR]  $msg" -ForegroundColor Red }

# Ensure TLS1.2 for downloads
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

$PyVersion    = '3.11.9'
$EmbedZipUrl  = "https://www.python.org/ftp/python/$PyVersion/python-$PyVersion-embed-amd64.zip"
$EmbedZipPath = Join-Path $ScriptDir "python-embed-$PyVersion.zip"
$EmbedDir     = Join-Path $ScriptDir 'py-embed'
$PyExe        = Join-Path $EmbedDir 'python.exe'

if (-not (Test-Path $PyExe)) {
  Write-Info "Downloading Python embeddable $PyVersion..."
  Invoke-WebRequest -Uri $EmbedZipUrl -OutFile $EmbedZipPath -UseBasicParsing
  if (Test-Path $EmbedDir) { Remove-Item -Recurse -Force $EmbedDir }
  Expand-Archive -Path $EmbedZipPath -DestinationPath $EmbedDir
  Remove-Item $EmbedZipPath -Force

  # Enable site and site-packages in the embeddable distribution
  $pth = Get-ChildItem -Path $EmbedDir -Filter 'python*.pth' | Select-Object -First 1
  if (-not $pth) { $pth = Get-ChildItem -Path $EmbedDir -Filter 'python*. _pth' | Select-Object -First 1 }
  if (-not $pth) { $pth = Get-ChildItem -Path $EmbedDir -Filter 'python*._pth' | Select-Object -First 1 }
  if (-not $pth) { throw "Could not locate python._pth file in $EmbedDir" }
  $lines = Get-Content $pth.FullName
  $lines = $lines | ForEach-Object { if ($_ -match '^#\s*import\s+site') { 'import site' } else { $_ } }
  if ($lines -notcontains 'Lib\site-packages') { $lines += 'Lib\site-packages' }
  Set-Content -Path $pth.FullName -Value $lines -Encoding ASCII

  # Bootstrap pip
  $getPip = Join-Path $EmbedDir 'get-pip.py'
  Write-Info "Downloading get-pip.py..."
  Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile $getPip -UseBasicParsing
  Write-Info "Installing pip into embedded Python..."
  & $PyExe $getPip --no-warn-script-location --disable-pip-version-check
}

# Install required packages into the embeddable's site-packages
Write-Info "Installing/upgrading required packages (local)..."
& $PyExe -m pip install --upgrade --no-warn-script-location --disable-pip-version-check pandas openpyxl lxml numpy

# Run the script
$ScriptPath = Join-Path $ScriptDir 'Monatsbericht_Bonus_Quartal.py'
if (-not (Test-Path $ScriptPath)) { throw "Could not find $ScriptPath" }

Write-Ok "Environment ready. Launching the report generator..."

# If parameters are passed, feed them via stdin prompts using here-strings is complex.
# Simpler: if all three are provided, set env vars the script reads via input() replacement.
# For now, run interactively.
& $PyExe $ScriptPath

