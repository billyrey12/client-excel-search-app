$ErrorActionPreference = 'Stop'

$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compiler -PathType Leaf)) {
    throw 'The 64-bit .NET Framework C# compiler was not found.'
}

$distRoot = Join-Path $PSScriptRoot 'dist'
$manifest = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'version.json') -Raw | ConvertFrom-Json
$outputFolder = Join-Path $distRoot "Client Excel Search $($manifest.version)"
New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null

$exePath = Join-Path $outputFolder 'Client Excel Search.exe'
$sourcePath = Join-Path $PSScriptRoot 'ClientExcelSearchLauncher.cs'
$bundleFiles = @(
    'ClientSearchApp.ps1',
    'ClientSearchWorker.ps1',
    'UpdateClientSearch.ps1',
    'version.json'
)
$compilerArgs = @(
    '/nologo',
    '/target:winexe',
    '/platform:x64',
    ('/out:' + $exePath),
    '/reference:System.Windows.Forms.dll',
    $sourcePath
)
foreach ($name in $bundleFiles) {
    $bundlePath = Join-Path $PSScriptRoot $name
    $compilerArgs += ('/resource:' + $bundlePath + ',Payload.' + $name)
}
& $compiler @compilerArgs
if ($LASTEXITCODE -ne 0) { throw 'The launcher EXE could not be built.' }

Write-Output "Built: $exePath"

