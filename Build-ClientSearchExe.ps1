$ErrorActionPreference = 'Stop'

$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compiler -PathType Leaf)) {
    throw 'The 64-bit .NET Framework C# compiler was not found.'
}

$distRoot = Join-Path $PSScriptRoot 'dist'
$packageFolder = Join-Path $distRoot 'Client Excel Search'
New-Item -ItemType Directory -Path $packageFolder -Force | Out-Null

$manifest = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'version.json') -Raw | ConvertFrom-Json
$exePath = Join-Path $packageFolder 'Client Excel Search.exe'
$sourcePath = Join-Path $PSScriptRoot 'ClientExcelSearchLauncher.cs'
$compilerArgs = @(
    '/nologo',
    '/target:winexe',
    '/platform:x64',
    ('/out:' + $exePath),
    '/reference:System.Windows.Forms.dll',
    $sourcePath
)
& $compiler @compilerArgs
if ($LASTEXITCODE -ne 0) { throw 'The launcher EXE could not be built.' }

$packageFiles = @(
    'ClientSearchApp.ps1',
    'ClientSearchWorker.ps1',
    'UpdateClientSearch.ps1',
    'version.json',
    'Start Client Search.cmd'
)
foreach ($name in $packageFiles) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination $packageFolder -Force
}

$instructions = @"
Client Excel Search $($manifest.version)

Keep all files in this folder together, then double-click Client Excel Search.exe.
Use Browse to choose the folder containing the Excel workbooks to search.

Requirements: 64-bit Windows and Microsoft Excel desktop. Internet is needed to check for updates.
When an update is available, choose Yes to install it.
"@
$instructionsPath = Join-Path $packageFolder 'Client Instructions.txt'
[System.IO.File]::WriteAllText($instructionsPath, $instructions, [System.Text.UTF8Encoding]::new($false))

$zipPath = Join-Path $distRoot "Client Excel Search-$($manifest.version).zip"
$archiveFiles = @((Join-Path $packageFolder 'Client Excel Search.exe'), $instructionsPath) + @(
    $packageFiles | ForEach-Object { Join-Path $packageFolder $_ }
)
Compress-Archive -Path $archiveFiles -DestinationPath $zipPath -CompressionLevel Optimal -Force

Write-Output "Built: $exePath"
Write-Output "Package: $zipPath"

