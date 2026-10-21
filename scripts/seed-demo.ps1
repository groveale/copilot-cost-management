param(
  [Parameter(Mandatory)][string] $SeedFile,
  [Parameter(Mandatory)][string] $TableUrl,
  [Parameter(Mandatory)][string] $BlobUrl,
  [Parameter(Mandatory)][string] $TenantId,
  [switch] $Apply
)
$ErrorActionPreference = 'Stop'
$mode = if ($Apply) { 'apply' } else { 'validate' }
$api = Join-Path $PSScriptRoot '..\api'
Push-Location $api
try {
  npm run build
  if ($LASTEXITCODE -ne 0) { throw 'API build failed.' }
  node dist\seed.js $SeedFile $TableUrl $BlobUrl $TenantId $mode
  if ($LASTEXITCODE -ne 0) { throw 'Seed validation/import failed; review any partial progress before retrying.' }
} finally { Pop-Location }
