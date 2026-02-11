param(
  [int]$Port = 4173,
  [string]$DefaultSpreadsheetId = '1KULUSaQIEFKtusmi-1E0HrR65KUA-bVa1-1Bjji5j6k'
)

$ErrorActionPreference = 'Stop'

function Write-Step {
  param([string]$Message)
  Write-Host "[TikTok GUI] $Message" -ForegroundColor Cyan
}

function Ensure-Directory {
  param([string]$Path)
  if (-not (Test-Path $Path)) {
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
  }
}

function Ensure-LocalConfig {
  param(
    [string]$ConfigPath,
    [string]$SpreadsheetId
  )

  Ensure-Directory -Path (Split-Path -Parent $ConfigPath)

  $config = @{}
  if (Test-Path $ConfigPath) {
    try {
      $raw = Get-Content -Path $ConfigPath -Raw -Encoding UTF8
      if (-not [string]::IsNullOrWhiteSpace($raw)) {
        $obj = $raw | ConvertFrom-Json
        if ($obj) {
          foreach ($prop in $obj.PSObject.Properties) {
            $config[$prop.Name] = $prop.Value
          }
        }
      }
    } catch {
      Write-Step "local-config.json parse failed; recreating minimal config."
      $config = @{}
    }
  }

  $needsWrite = $false
  if (-not $config.ContainsKey('spreadsheetId') -or [string]::IsNullOrWhiteSpace([string]$config['spreadsheetId'])) {
    $config['spreadsheetId'] = $SpreadsheetId
    $needsWrite = $true
  }

  if ($needsWrite -or -not (Test-Path $ConfigPath)) {
    $json = $config | ConvertTo-Json -Depth 16
    Set-Content -Path $ConfigPath -Value "$json`n" -Encoding UTF8
    Write-Step "local-config.json has been updated."
  }
}

function Install-PortableNode {
  param([string]$ProjectRoot)

  $runtimeRoot = Join-Path $ProjectRoot '.runtime'
  $nodeRoot = Join-Path $runtimeRoot 'node'
  $zipPath = Join-Path $runtimeRoot 'node-win-x64.zip'
  $extractRoot = Join-Path $runtimeRoot 'node-extract'

  Ensure-Directory -Path $runtimeRoot

  Write-Step "Node.js not found. Downloading portable LTS runtime..."
  $index = Invoke-RestMethod -Uri 'https://nodejs.org/dist/index.json'
  $target = $index | Where-Object {
    $_.lts -and $_.files -and ($_.files -contains 'win-x64-zip')
  } | Select-Object -First 1

  if (-not $target) {
    throw 'Failed to resolve Node.js LTS win-x64 zip package.'
  }

  $version = [string]$target.version
  $downloadUrl = "https://nodejs.org/dist/$version/node-$version-win-x64.zip"
  Write-Step "Downloading $version ..."

  if (Test-Path $zipPath) {
    Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue
  }
  if (Test-Path $extractRoot) {
    Remove-Item -Path $extractRoot -Recurse -Force -ErrorAction SilentlyContinue
  }

  Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath
  Expand-Archive -Path $zipPath -DestinationPath $extractRoot -Force

  $extractedDir = Get-ChildItem -Path $extractRoot -Directory | Select-Object -First 1
  if (-not $extractedDir) {
    throw 'Failed to expand Node.js package.'
  }

  if (Test-Path $nodeRoot) {
    Remove-Item -Path $nodeRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
  Move-Item -Path $extractedDir.FullName -Destination $nodeRoot

  Remove-Item -Path $extractRoot -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue

  $nodeCmd = Join-Path $nodeRoot 'node.exe'
  $npmCmd = Join-Path $nodeRoot 'npm.cmd'
  if (-not (Test-Path $nodeCmd) -or -not (Test-Path $npmCmd)) {
    throw 'Portable Node.js installation failed.'
  }

  return @{
    Node = $nodeCmd
    Npm = $npmCmd
    Source = 'portable'
  }
}

function Resolve-NodeTools {
  param([string]$ProjectRoot)

  $portableNode = Join-Path $ProjectRoot '.runtime\node\node.exe'
  $portableNpm = Join-Path $ProjectRoot '.runtime\node\npm.cmd'
  if ((Test-Path $portableNode) -and (Test-Path $portableNpm)) {
    return @{
      Node = $portableNode
      Npm = $portableNpm
      Source = 'portable'
    }
  }

  $nodeCmd = Get-Command node -ErrorAction SilentlyContinue
  if ($nodeCmd) {
    $nodePath = $nodeCmd.Source
    $npmPath = Join-Path (Split-Path -Parent $nodePath) 'npm.cmd'
    if (-not (Test-Path $npmPath)) {
      $npmCmd = Get-Command npm -ErrorAction SilentlyContinue
      if ($npmCmd) {
        $npmPath = $npmCmd.Source
      } else {
        $npmPath = 'npm.cmd'
      }
    }
    return @{
      Node = $nodePath
      Npm = $npmPath
      Source = 'system'
    }
  }

  return Install-PortableNode -ProjectRoot $ProjectRoot
}

function Ensure-NpmDependencies {
  param(
    [string]$ProjectRoot,
    [string]$NpmCmd
  )

  $nodeModules = Join-Path $ProjectRoot 'node_modules'
  $lockPath = Join-Path $ProjectRoot 'package-lock.json'
  $markerPath = Join-Path $ProjectRoot 'app-data\.deps-lockhash'

  if (-not (Test-Path $lockPath)) {
    throw 'package-lock.json not found.'
  }

  $lockHash = (Get-FileHash -Path $lockPath -Algorithm SHA256).Hash
  $installNeeded = -not (Test-Path $nodeModules)
  if (-not $installNeeded) {
    if (Test-Path $markerPath) {
      $prev = (Get-Content -Path $markerPath -Raw -Encoding UTF8).Trim()
      $installNeeded = ($prev -ne $lockHash)
    } else {
      $installNeeded = $true
    }
  }

  if ($installNeeded) {
    Write-Step "Installing npm dependencies (npm ci)..."
    & $NpmCmd ci --no-fund --no-audit
    if ($LASTEXITCODE -ne 0) {
      throw "npm ci failed (exit=$LASTEXITCODE)"
    }
    Set-Content -Path $markerPath -Value "$lockHash`n" -Encoding UTF8
  } else {
    Write-Step "Dependencies are up to date."
  }
}

function Test-PortListening {
  param([int]$Port)
  try {
    $conn = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction Stop
    return [bool]$conn
  } catch {
    return $false
  }
}

function Start-ServerProcess {
  param(
    [string]$NodeCmd,
    [string]$ProjectRoot
  )

  $entry = Join-Path $ProjectRoot 'app/server.mjs'
  try {
    Start-Process -FilePath $NodeCmd -ArgumentList @($entry) -WorkingDirectory $ProjectRoot | Out-Null
    return
  } catch {
    & cmd /c "start `"`" /b `"$NodeCmd`" `"$entry`""
  }
}

function Wait-Health {
  param(
    [int]$Port,
    [int]$TimeoutSeconds = 40
  )

  $url = "http://127.0.0.1:$Port/api/health"
  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 500
    try {
      $res = Invoke-RestMethod -Uri $url -TimeoutSec 2
      if ($res.ok) {
        return $true
      }
    } catch {
      # keep waiting
    }
  }
  return $false
}

function Open-GuiPage {
  param([int]$Port)
  $url = "http://127.0.0.1:$Port"
  try {
    Start-Process $url | Out-Null
  } catch {
    & cmd /c "start `"`" `"$url`""
  }
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDir
Set-Location $projectRoot

Write-Step "Preparing startup..."

$configPath = Join-Path $projectRoot 'app-data\local-config.json'
Ensure-LocalConfig -ConfigPath $configPath -SpreadsheetId $DefaultSpreadsheetId

$tools = Resolve-NodeTools -ProjectRoot $projectRoot
Write-Step "Node source: $($tools.Source)"
Ensure-NpmDependencies -ProjectRoot $projectRoot -NpmCmd $tools.Npm

if (Test-PortListening -Port $Port) {
  Write-Step "Server already running. Opening browser."
  Open-GuiPage -Port $Port
  exit 0
}

Write-Step "Starting GUI server..."
Start-ServerProcess -NodeCmd $tools.Node -ProjectRoot $projectRoot

if (Wait-Health -Port $Port) {
  Write-Step "Opening GUI..."
  Open-GuiPage -Port $Port
  Write-Step "Ready: http://127.0.0.1:$Port"
  exit 0
}

throw "Server did not become healthy in time. Check: http://127.0.0.1:$Port/api/health"
