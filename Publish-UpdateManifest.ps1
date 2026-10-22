param(
    [Parameter(Mandatory = $true)]
    [version]$Version
)

$files = foreach ($name in @('ClientSearchApp.ps1', 'ClientSearchWorker.ps1')) {
    $path = Join-Path $PSScriptRoot $name
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing update file: $name" }
    $text = Get-Content -LiteralPath $path -Raw
    $canonicalBytes = [System.Text.UTF8Encoding]::new($false).GetBytes(($text -replace "`r`n?", "`n"))
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $hash = [System.BitConverter]::ToString($sha.ComputeHash($canonicalBytes)).Replace('-', '').ToLowerInvariant() } finally { $sha.Dispose() }
    [pscustomobject]@{ name = $name; sha256 = $hash }
}

$manifest = [ordered]@{ version = $Version.ToString(); files = @($files) }
$json = ConvertTo-Json -InputObject $manifest -Depth 5
[System.IO.File]::WriteAllText((Join-Path $PSScriptRoot 'version.json'), $json, [System.Text.UTF8Encoding]::new($false))
Write-Output "Created version.json for version $Version. Commit it with the app scripts to the main branch."

