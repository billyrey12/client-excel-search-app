Add-Type -AssemblyName System.Windows.Forms

$repository = 'billyrey12/client-excel-search-app'
$branch = 'main'
$manifestPath = Join-Path $PSScriptRoot 'version.json'
$stagePath = Join-Path $PSScriptRoot ('.update-' + [Guid]::NewGuid().ToString('N'))
$allowedFiles = @('ClientSearchApp.ps1', 'ClientSearchWorker.ps1')
$filesToReplace = @($allowedFiles + 'version.json')
$updatedFiles = New-Object 'System.Collections.Generic.List[string]'
$backupFiles = @{}
$installRequested = $false

function Show-UpdateMessage {
    param([string]$Message, [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information)
    [void][System.Windows.Forms.MessageBox]::Show($Message, 'Client Excel Search Update', [System.Windows.Forms.MessageBoxButtons]::OK, $Icon)
}

function Restore-UpdatedFiles {
    foreach ($name in @($updatedFiles)) {
        $destination = Join-Path $PSScriptRoot $name
        if ($backupFiles.ContainsKey($name)) {
            Copy-Item -LiteralPath $backupFiles[$name] -Destination $destination -Force
        } elseif (Test-Path -LiteralPath $destination) {
            Remove-Item -LiteralPath $destination -Force
        }
    }
}

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $manifestUrl = "https://raw.githubusercontent.com/$repository/$branch/version.json"
    $remoteManifest = Invoke-RestMethod -Uri $manifestUrl -TimeoutSec 5 -ErrorAction Stop
    $remoteVersion = [version]$remoteManifest.version

    $localVersion = [version]'0.0.0'
    if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
        $localManifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json -ErrorAction Stop
        $localVersion = [version]$localManifest.version
    }
    if ($remoteVersion -le $localVersion) { exit 0 }

    $manifestFiles = @($remoteManifest.files)
    if ($manifestFiles.Count -ne $allowedFiles.Count) { throw 'The update manifest has an unexpected file list.' }
    foreach ($name in $allowedFiles) {
        $entry = @($manifestFiles | Where-Object { $_.name -ceq $name })
        if ($entry.Count -ne 1 -or [string]$entry[0].sha256 -notmatch '^[A-Fa-f0-9]{64}$') {
            throw "The update manifest is invalid for $name."
        }
    }

    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Version $remoteVersion is available. Install it now?`r`n`r`nThe app will download and verify the update before opening.",
        'Client Excel Search Update',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { exit 0 }
    $installRequested = $true

    New-Item -ItemType Directory -Path $stagePath -Force | Out-Null
    foreach ($name in $allowedFiles) {
        $entry = $manifestFiles | Where-Object { $_.name -ceq $name } | Select-Object -First 1
        $url = "https://raw.githubusercontent.com/$repository/$branch/$([Uri]::EscapeDataString($name))"
        $stagedPath = Join-Path $stagePath $name
        Invoke-WebRequest -Uri $url -OutFile $stagedPath -TimeoutSec 30 -UseBasicParsing -ErrorAction Stop
        $actualHash = (Get-FileHash -LiteralPath $stagedPath -Algorithm SHA256).Hash
        if ($actualHash -ine [string]$entry.sha256) { throw "The downloaded file failed verification: $name" }
    }

    $backupDirectory = Join-Path $stagePath 'backup'
    New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
    $manifestJson = ConvertTo-Json -InputObject $remoteManifest -Depth 5
    [System.IO.File]::WriteAllText((Join-Path $stagePath 'version.json'), $manifestJson, [System.Text.UTF8Encoding]::new($false))
    foreach ($name in $filesToReplace) {
        $destination = Join-Path $PSScriptRoot $name
        if (Test-Path -LiteralPath $destination -PathType Leaf) {
            $backupPath = Join-Path $backupDirectory $name
            Copy-Item -LiteralPath $destination -Destination $backupPath -Force
            $backupFiles[$name] = $backupPath
        }
        $updatedFiles.Add($name)
        Move-Item -LiteralPath (Join-Path $stagePath $name) -Destination $destination -Force
    }

    Show-UpdateMessage "Client Excel Search was updated to version $remoteVersion. Click OK to open the app."
} catch {
    try { Restore-UpdatedFiles } catch { }
    if ($installRequested) {
        Show-UpdateMessage ("The update could not be installed. The current app will open.`r`n`r`n" + $_.Exception.Message) ([System.Windows.Forms.MessageBoxIcon]::Warning)
    }
} finally {
    if (Test-Path -LiteralPath $stagePath) {
        Remove-Item -LiteralPath $stagePath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

