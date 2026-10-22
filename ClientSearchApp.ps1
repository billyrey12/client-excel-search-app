Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$defaultFolder = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
$script:SearchWhitespaceChars = [char[]]@(9, 10, 13, 160, 8199, 8239, 8203, 65279)
$script:WorkerPath = Join-Path $PSScriptRoot 'ClientSearchWorker.ps1'
$script:WorkerProcess = $null
$script:ProgressReader = $null
$script:ProgressFile = $null
$script:CancelFile = $null
$script:ReadBuffer = New-Object 'System.Collections.Generic.List[byte]'
$script:WorkerDone = $null
$script:WorkerError = ''
$script:WorkerCacheWarning = ''
$script:WorkerWarnings = New-Object 'System.Collections.Generic.List[string]'
$script:ResultCount = 0
$script:CloseWhenDone = $false
$script:IndexRows = New-Object 'System.Collections.Generic.List[object]'
$script:IndexReady = $false
$script:IndexFolder = ''
$script:IndexSignature = ''
$script:IndexFileCount = 0
$script:IndexHeaderCount = 0
$script:IndexCachePath = ''

function ConvertTo-CellString {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal]) {
        return ([double]$Value).ToString('0.################', [Globalization.CultureInfo]::InvariantCulture)
    }
    return ([string]$Value).Trim()
}

function Normalize-SearchText {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value.IndexOf('  ') -lt 0 -and $Value.IndexOfAny($script:SearchWhitespaceChars) -lt 0) { return $Value.Trim() }
    $text = $Value.Replace(([char]0x00A0).ToString(), ' ').Replace(([char]0x2007).ToString(), ' ').Replace(([char]0x202F).ToString(), ' ')
    $text = $text.Replace(([char]0x200B).ToString(), '').Replace(([char]0xFEFF).ToString(), '')
    return [regex]::Replace($text, '\s+', ' ').Trim()
}

function Test-SearchMatch {
    param([string]$Value, [string]$Query, [switch]$CofNumber, [AllowNull()][string]$QueryDigits = $null, [AllowNull()][string]$NormalizedQuery = $null)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    if ($null -eq $NormalizedQuery) { $NormalizedQuery = Normalize-SearchText $Query }
    $normalizedValue = Normalize-SearchText $Value
    if ($normalizedValue.IndexOf($NormalizedQuery, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    if ($CofNumber) {
        if ($null -eq $QueryDigits) { $QueryDigits = [regex]::Replace($Query, '\D', '').TrimStart([char]'0') }
        $valueDigits = [regex]::Replace($Value, '\D', '').TrimStart([char]'0')
        if ($QueryDigits.Length -gt 0 -and $valueDigits.Length -gt 0 -and $valueDigits.Contains($QueryDigits)) { return $true }
    }
    return $false
}

function Get-ExcelFiles {
    param([string]$Folder)
    return @(Get-ChildItem -LiteralPath $Folder -File -Recurse -ErrorAction SilentlyContinue | Where-Object {
        $_.Extension -in @('.xls', '.xlsx', '.xlsm') -and $_.Name -notlike '~$*'
    } | Sort-Object FullName)
}

function Get-FolderSignature {
    param($Files)
    $parts = foreach ($file in $Files) { "$($file.FullName)`t$($file.Length)`t$($file.LastWriteTimeUtc.Ticks)" }
    return [string]::Join([Environment]::NewLine, [string[]]$parts)
}

function Get-IndexCachePath {
    param($Files)
    $filePaths = foreach ($file in $Files) { $file.FullName.ToUpperInvariant() }
    $cacheIdentity = [string]::Join([Environment]::NewLine, [string[]]$filePaths)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($cacheIdentity)
        $key = [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '')
    } finally {
        $sha.Dispose()
    }
    $cacheRoot = Join-Path $env:LOCALAPPDATA 'ClientExcelSearch'
    return Join-Path $cacheRoot "index-$key.jsonl"
}

function Load-IndexCache {
    param([string]$Path, [string]$FolderKey, [string]$Signature)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    $reader = $null
    try {
        $reader = [System.IO.StreamReader]::new($Path, [System.Text.Encoding]::UTF8, $true)
        $line = $reader.ReadLine()
        if (-not $line) { return $false }
        $metadata = ConvertFrom-Json -InputObject $line -ErrorAction Stop
        if ($metadata.type -ne 'meta' -or $metadata.version -ne 3 -or $metadata.signature -ne $Signature) { return $false }
        $script:IndexRows.Clear()
        $stats = $null
        while ($null -ne ($line = $reader.ReadLine())) {
            if (-not $line) { continue }
            $batch = ConvertFrom-Json -InputObject $line -ErrorAction Stop
            if ($batch.type -eq 'batch') {
                foreach ($entry in @($batch.rows)) { $script:IndexRows.Add($entry) }
            } elseif ($batch.type -eq 'stats') {
                $stats = $batch
            }
        }
        if ($null -eq $stats -or $script:IndexRows.Count -ne [int]$stats.rows) {
            $script:IndexRows.Clear()
            return $false
        }
        $script:IndexFolder = $FolderKey
        $script:IndexSignature = $Signature
        $script:IndexFileCount = [int]$metadata.fileCount
        $script:IndexHeaderCount = [int]$stats.headers
        $script:IndexReady = $true
        return $true
    } catch {
        $script:IndexRows.Clear()
        return $false
    } finally {
        if ($null -ne $reader) { $reader.Dispose() }
    }
}

function Release-ComObject {
    param($Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

function ConvertTo-PowerShellLiteral {
    param([string]$Value)
    return "'" + $Value.Replace("'", "''") + "'"
}

function Open-ResultRow {
    param($Result)
    if ($null -eq $Result) { return }
    $excel = $null
    $book = $null
    $sheet = $null
    $target = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        $book = $excel.Workbooks.Open($Result.Path, 0, $false)
        $sheet = $book.Worksheets.Item($Result.Sheet)
        $sheet.Activate()
        $column = if ($Result.NameColumn -gt 0) { $Result.NameColumn } else { $Result.CofColumn }
        if ($column -le 0) { $column = 1 }
        $target = $sheet.Cells.Item($Result.Row, $column)
        $excel.Goto($target, $true) | Out-Null
        $excel.ActiveWindow.ScrollRow = [Math]::Max(1, $Result.Row - 5)
    } catch {
        [System.Windows.Forms.MessageBox]::Show(('Could not open the file in Excel.' + [Environment]::NewLine + $_.Exception.Message), 'Client Excel Search', 'OK', 'Error') | Out-Null
    } finally {
        Release-ComObject $target
        Release-ComObject $sheet
        Release-ComObject $book
        Release-ComObject $excel
    }
}

function Remove-WorkerFiles {
    if ($null -ne $script:ProgressReader) {
        $script:ProgressReader.Dispose()
        $script:ProgressReader = $null
    }
    if ($null -ne $script:WorkerProcess) {
        $script:WorkerProcess.Dispose()
        $script:WorkerProcess = $null
    }
    foreach ($path in @($script:ProgressFile, $script:CancelFile)) {
        if ($path -and (Test-Path -LiteralPath $path)) {
            Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
        }
    }
    $script:ProgressFile = $null
    $script:CancelFile = $null
    $script:ReadBuffer.Clear()
}

function Complete-Search {
    param($Form, $SearchButton, $BrowseButton, $CancelButton, $OpenButton, $ProgressBar, $StatusLabel, $Timer)
    $Timer.Stop()
    $processExitCode = 0
    if ($null -ne $script:WorkerProcess) {
        try { $processExitCode = $script:WorkerProcess.ExitCode } catch { $processExitCode = -1 }
    }
    $SearchButton.Enabled = $true
    $BrowseButton.Enabled = $true
    $CancelButton.Enabled = $false
    $OpenButton.Enabled = ($script:ResultCount -gt 0)
    $ProgressBar.Style = 'Continuous'
    $ProgressBar.MarqueeAnimationSpeed = 0
    if ($ProgressBar.Maximum -gt 0) { $ProgressBar.Value = $ProgressBar.Maximum }

    if ($null -ne $script:WorkerDone) {
        if (-not $script:WorkerDone.cancelled -and $script:WorkerDone.errors -eq 0 -and -not $script:WorkerError -and $script:WorkerDone.indexValid) {
            # The worker writes the full index to disk; load it on demand so the UI does not deserialize every row during the scan.
            $script:IndexReady = $false
            $script:IndexFileCount = [int]$script:WorkerDone.files
            $script:IndexHeaderCount = [int]$script:WorkerDone.headers
        } else {
            $script:IndexReady = $false
        }
        if ($script:WorkerDone.cancelled) {
            $StatusLabel.Text = "Search cancelled. $($script:ResultCount) result(s) were found before cancellation."
        } elseif ($script:WorkerDone.files -eq 0) {
            $StatusLabel.Text = 'Finished: no Excel files were found in the folder.'
        } elseif ($script:ResultCount -gt 0) {
            $StatusLabel.Text = "Finished: $($script:ResultCount) result(s) found in $($script:WorkerDone.files) Excel file(s). Select a result and click Open Row."
        } elseif ($script:WorkerDone.headers -eq 0) {
            $StatusLabel.Text = "Finished: no Name or COF NO column was detected in $($script:WorkerDone.files) file(s)."
        } else {
            $StatusLabel.Text = "Finished: no matches found in $($script:WorkerDone.files) Excel file(s)."
        }
        if ($script:WorkerDone.errors -gt 0) {
            $StatusLabel.Text += " Could not read $($script:WorkerDone.errors) file(s)."
        }
        if ($script:WorkerCacheWarning) {
            $StatusLabel.Text += ' The local search index could not be saved, so the next search may take longer.'
        }
    } elseif ($script:WorkerError) {
        $StatusLabel.Text = 'The search did not finish because of an error.'
    } else {
        $StatusLabel.Text = "The search process ended (exit code $processExitCode)."
    }

    if ($script:WorkerError) {
        [System.Windows.Forms.MessageBox]::Show($script:WorkerError, 'Client Excel Search', 'OK', 'Error') | Out-Null
    }
    if ($script:WorkerWarnings.Count -gt 0) {
        $details = ($script:WorkerWarnings | Select-Object -First 5) -join [Environment]::NewLine
        if ($script:WorkerWarnings.Count -gt 5) { $details += [Environment]::NewLine + "...at $($script:WorkerWarnings.Count - 5) pang file." }
        $warningText = 'Some files could not be read:' + [Environment]::NewLine + [Environment]::NewLine + $details
        [System.Windows.Forms.MessageBox]::Show($warningText, 'Client Excel Search', 'OK', 'Warning') | Out-Null
    }

    Remove-WorkerFiles
    if ($script:CloseWhenDone) { $Form.Close() }
}

function Read-WorkerMessages {
    param($Form, $SearchButton, $BrowseButton, $CancelButton, $OpenButton, $ProgressBar, $StatusLabel, $Grid, $Timer)
    if ($null -eq $script:ProgressReader) { return }
    $messagesRead = 0
    while ($script:ProgressReader.Position -lt $script:ProgressReader.Length -and $messagesRead -lt 20) {
        $byte = $script:ProgressReader.ReadByte()
        if ($byte -lt 0) { break }
        if ($byte -eq 10) {
            $line = [System.Text.Encoding]::UTF8.GetString($script:ReadBuffer.ToArray())
            $script:ReadBuffer.Clear()
            if (-not $line) { continue }
            try { $message = ConvertFrom-Json -InputObject $line } catch { continue }
            $messagesRead++
            switch ([string]$message.type) {
                'start' {
                    $total = [int]$message.total
                    $ProgressBar.Style = 'Marquee'
                    $ProgressBar.MarqueeAnimationSpeed = 25
                    $StatusLabel.Text = if ($total -gt 0) { "Starting search in $total Excel file(s)..." } else { 'No Excel files were found in the folder.' }
                }
                'progress' {
                    $StatusLabel.Text = "Reading file $($message.current) of $($message.total): $($message.file)"
                    $ProgressBar.Style = 'Marquee'
                    $ProgressBar.MarqueeAnimationSpeed = 25
                }
                'detail' {
                    $StatusLabel.Text = [string]$message.text
                }
                'fileDone' {
                    $total = [Math]::Max(1, [int]$message.total)
                    $ProgressBar.Style = 'Continuous'
                    $ProgressBar.MarqueeAnimationSpeed = 0
                    $ProgressBar.Maximum = $total
                    $ProgressBar.Value = [Math]::Min($total, [int]$message.current)
                }
                'result' {
                    $index = $Grid.Rows.Add($message.Name, $message.CofNo, $message.MatchedIn, $message.File, $message.Sheet, $message.Row)
                    $Grid.Rows[$index].Tag = $message
                    $script:ResultCount++
                    $OpenButton.Enabled = $true
                    $StatusLabel.Text = "$($script:ResultCount) result(s) found. Searching the remaining files..."
                }
                'warning' {
                    $script:WorkerWarnings.Add([string]$message.text)
                }
                'cacheWarning' {
                    $script:WorkerCacheWarning = [string]$message.text
                }
                'error' {
                    $script:WorkerError = [string]$message.text
                }
                'done' {
                    $script:WorkerDone = $message
                }
            }
        } elseif ($byte -ne 13) {
            $script:ReadBuffer.Add([byte]$byte)
        }
    }

    if ($null -ne $script:WorkerProcess) {
        $script:WorkerProcess.Refresh()
        if ($script:WorkerProcess.HasExited -and $script:ProgressReader.Position -ge $script:ProgressReader.Length) {
            Complete-Search $Form $SearchButton $BrowseButton $CancelButton $OpenButton $ProgressBar $StatusLabel $Timer
        }
    }
}

function Search-IndexedRows {
    param([string]$Query, $Grid, $StatusLabel, $OpenButton, $ProgressBar)
    $Grid.Rows.Clear()
    $script:ResultCount = 0
    $queryDigits = [regex]::Replace($Query, '\D', '').TrimStart([char]'0')
    $normalizedQuery = Normalize-SearchText $Query
    foreach ($entry in $script:IndexRows) {
        $nameMatch = Test-SearchMatch $entry.Name $Query -QueryDigits $queryDigits -NormalizedQuery $normalizedQuery
        $cofMatch = Test-SearchMatch $entry.CofNo $Query -CofNumber -QueryDigits $queryDigits -NormalizedQuery $normalizedQuery
        if (-not $nameMatch -and -not $cofMatch) { continue }
        $matchedIn = if ($nameMatch -and $cofMatch) { 'Name, COF NO' } elseif ($nameMatch) { 'Name' } else { 'COF NO' }
        $result = [pscustomobject]@{
            Name = $entry.Name
            CofNo = $entry.CofNo
            MatchedIn = $matchedIn
            File = $entry.File
            Sheet = $entry.Sheet
            Row = $entry.Row
            Path = $entry.Path
            NameColumn = $entry.NameColumn
            CofColumn = $entry.CofColumn
        }
        $rowIndex = $Grid.Rows.Add($result.Name, $result.CofNo, $result.MatchedIn, $result.File, $result.Sheet, $result.Row)
        $Grid.Rows[$rowIndex].Tag = $result
        $script:ResultCount++
    }
    $OpenButton.Enabled = ($script:ResultCount -gt 0)
    $ProgressBar.Style = 'Continuous'
    $ProgressBar.MarqueeAnimationSpeed = 0
    $ProgressBar.Maximum = [Math]::Max(1, $script:IndexFileCount)
    $ProgressBar.Value = $ProgressBar.Maximum
    if ($script:IndexHeaderCount -eq 0 -and $script:IndexFileCount -gt 0) {
        $StatusLabel.Text = "Finished: no Name or COF NO column was detected in the $($script:IndexFileCount) indexed file(s)."
    } elseif ($script:ResultCount -gt 0) {
        $StatusLabel.Text = "Finished: $($script:ResultCount) result(s) found in the $($script:IndexFileCount) indexed file(s)."
    } else {
        $StatusLabel.Text = "Finished: no matches found in the $($script:IndexFileCount) indexed file(s)."
    }
}

function Start-Search {
    param($FolderBox, $QueryBox, $Grid, $StatusLabel, $SearchButton, $BrowseButton, $CancelButton, $OpenButton, $ProgressBar, $Timer)
    $query = $QueryBox.Text.Trim()
    $folder = $FolderBox.Text.Trim()
    if (-not $query) {
        [System.Windows.Forms.MessageBox]::Show('Enter a Name or COF NO to search for.', 'Client Excel Search', 'OK', 'Information') | Out-Null
        $QueryBox.Focus()
        return
    }
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show('The selected folder could not be found.', 'Client Excel Search', 'OK', 'Warning') | Out-Null
        return
    }
    if (-not (Test-Path -LiteralPath $script:WorkerPath -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show('ClientSearchWorker.ps1 is missing. Restore it to the app folder.', 'Client Excel Search', 'OK', 'Error') | Out-Null
        return
    }

    $Grid.Rows.Clear()
    $files = @(Get-ExcelFiles $folder)
    $signature = Get-FolderSignature $files
    $folderKey = [System.IO.Path]::GetFullPath($folder).TrimEnd([char[]]@('\', '/'))
    $script:IndexCachePath = Get-IndexCachePath $files
    if ($script:IndexReady -and $script:IndexFolder -eq $folderKey -and $script:IndexSignature -eq $signature) {
        Search-IndexedRows $query $Grid $StatusLabel $OpenButton $ProgressBar
        return
    }
    if (Test-Path -LiteralPath $script:IndexCachePath -PathType Leaf) {
        $StatusLabel.Text = 'Checking the saved search index...'
        $ProgressBar.Style = 'Marquee'
        $ProgressBar.MarqueeAnimationSpeed = 25
        $StatusLabel.Refresh()
        $ProgressBar.Refresh()
    }
    if (Load-IndexCache $script:IndexCachePath $folderKey $signature) {
        Search-IndexedRows $query $Grid $StatusLabel $OpenButton $ProgressBar
        return
    }
    if ($files.Count -eq 0) {
        $script:IndexRows.Clear()
        $script:IndexReady = $true
        $script:IndexFolder = $folderKey
        $script:IndexSignature = $signature
        $script:IndexFileCount = 0
        $script:IndexHeaderCount = 0
        $ProgressBar.Style = 'Continuous'
        $ProgressBar.Value = 0
        $OpenButton.Enabled = $false
        $StatusLabel.Text = 'Finished: no Excel files were found in the folder.'
        return
    }
    $script:IndexRows.Clear()
    $script:IndexReady = $false
    $script:IndexFolder = $folderKey
    $script:IndexSignature = $signature
    $script:IndexFileCount = $files.Count
    $script:ResultCount = 0
    $script:WorkerDone = $null
    $script:WorkerError = ''
    $script:WorkerCacheWarning = ''
    $script:WorkerWarnings.Clear()
    $script:CloseWhenDone = $false
    $SearchButton.Enabled = $false
    $BrowseButton.Enabled = $false
    $CancelButton.Enabled = $true
    $OpenButton.Enabled = $false
    $ProgressBar.Value = 0
    $ProgressBar.Style = 'Marquee'
    $ProgressBar.MarqueeAnimationSpeed = 25
    $StatusLabel.Text = 'Starting search...'

    try {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ClientSearch-' + [Guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        $script:ProgressFile = Join-Path $tempRoot 'progress.jsonl'
        $script:CancelFile = Join-Path $tempRoot 'cancel'
        [System.IO.File]::WriteAllText($script:ProgressFile, '', [System.Text.UTF8Encoding]::new($false))
        $script:ProgressReader = [System.IO.File]::Open($script:ProgressFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $script:ReadBuffer.Clear()

        $workerCommand = '& ' + (ConvertTo-PowerShellLiteral $script:WorkerPath) +
            ' -Folder ' + (ConvertTo-PowerShellLiteral $folder) +
            ' -Query ' + (ConvertTo-PowerShellLiteral $query) +
            ' -ProgressFile ' + (ConvertTo-PowerShellLiteral $script:ProgressFile) +
            ' -CancelFile ' + (ConvertTo-PowerShellLiteral $script:CancelFile) +
            ' -CachePath ' + (ConvertTo-PowerShellLiteral $script:IndexCachePath) +
            ' -FolderKey ' + (ConvertTo-PowerShellLiteral $folderKey) +
            ' -Signature ' + (ConvertTo-PowerShellLiteral $signature)
        $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($workerCommand))
        $powerShellExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $powerShellExe
        $startInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -EncodedCommand $encodedCommand"
        $startInfo.WorkingDirectory = $PSScriptRoot
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
        $script:WorkerProcess = [System.Diagnostics.Process]::Start($startInfo)
        if ($null -eq $script:WorkerProcess) { throw 'Could not start the background search process.' }
        $Timer.Start()
    } catch {
        $SearchButton.Enabled = $true
        $BrowseButton.Enabled = $true
        $CancelButton.Enabled = $false
        $ProgressBar.Style = 'Continuous'
        $ProgressBar.MarqueeAnimationSpeed = 0
        $StatusLabel.Text = 'Could not start the search.'
        Remove-WorkerFiles
        [System.Windows.Forms.MessageBox]::Show(('Could not start the background search.' + [Environment]::NewLine + $_.Exception.Message), 'Client Excel Search', 'OK', 'Error') | Out-Null
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Client Excel Search'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1120, 700)
$form.MinimumSize = New-Object System.Drawing.Size(900, 560)
$form.BackColor = [System.Drawing.Color]::FromArgb(247, 249, 252)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$form.Padding = New-Object System.Windows.Forms.Padding(16)

$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = 'Fill'
$layout.ColumnCount = 4
$layout.RowCount = 6
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 132)))
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 142)))
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 128)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 68)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 44)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$form.Controls.Add($layout)

$titlePanel = New-Object System.Windows.Forms.Panel
$titlePanel.Dock = 'Fill'
$title = New-Object System.Windows.Forms.Label
$title.Text = 'Find a Client'
$title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 18)
$title.ForeColor = [System.Drawing.Color]::FromArgb(28, 47, 75)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(0, 0)
$titlePanel.Controls.Add($title)
$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = 'Enter a Name or COF NO to find the workbook, sheet, and row.'
$subtitle.ForeColor = [System.Drawing.Color]::FromArgb(91, 105, 124)
$subtitle.AutoSize = $true
$subtitle.Location = New-Object System.Drawing.Point(2, 37)
$titlePanel.Controls.Add($subtitle)
$layout.Controls.Add($titlePanel, 0, 0)
$layout.SetColumnSpan($titlePanel, 4)

$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Text = 'Files folder'
$folderLabel.TextAlign = 'MiddleLeft'
$folderLabel.Dock = 'Fill'
$layout.Controls.Add($folderLabel, 0, 1)
$folderBox = New-Object System.Windows.Forms.TextBox
$folderBox.Text = $defaultFolder
$folderBox.Dock = 'Fill'
$folderBox.Margin = New-Object System.Windows.Forms.Padding(0, 7, 10, 7)
$layout.Controls.Add($folderBox, 1, 1)
$layout.SetColumnSpan($folderBox, 2)
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = 'Choose file...'
$browseButton.Dock = 'Fill'
$browseButton.Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 3)
$layout.Controls.Add($browseButton, 3, 1)
$browseTip = New-Object System.Windows.Forms.ToolTip
$browseTip.SetToolTip($browseButton, 'Choose any Excel workbook. The app will search all Excel files in its folder.')
$browseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Choose any Excel workbook in the folder to search'
$dialog.Filter = 'Excel workbooks (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm|All files (*.*)|*.*'
    $dialog.CheckFileExists = $true
    $dialog.Multiselect = $false
    if (Test-Path -LiteralPath $folderBox.Text -PathType Container) { $dialog.InitialDirectory = [System.IO.Path]::GetFullPath($folderBox.Text) }
    if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $folderBox.Text = [System.IO.Path]::GetDirectoryName($dialog.FileName)
    }
    $dialog.Dispose()
})

$queryLabel = New-Object System.Windows.Forms.Label
$queryLabel.Text = 'Name or COF NO'
$queryLabel.TextAlign = 'MiddleLeft'
$queryLabel.Dock = 'Fill'
$layout.Controls.Add($queryLabel, 0, 2)
$queryBox = New-Object System.Windows.Forms.TextBox
$queryBox.Dock = 'Fill'
$queryBox.Margin = New-Object System.Windows.Forms.Padding(0, 9, 10, 9)
$layout.Controls.Add($queryBox, 1, 2)
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = 'Search'
$searchButton.Dock = 'Fill'
$searchButton.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 4)
$searchButton.BackColor = [System.Drawing.Color]::FromArgb(35, 94, 160)
$searchButton.ForeColor = [System.Drawing.Color]::White
$searchButton.FlatStyle = 'Flat'
$layout.Controls.Add($searchButton, 2, 2)
$layout.SetColumnSpan($searchButton, 2)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Dock = 'Fill'
$progressBar.Minimum = 0
$progressBar.Maximum = 1
$progressBar.Value = 0
$progressBar.Style = 'Continuous'
$progressBar.MarqueeAnimationSpeed = 0
$layout.Controls.Add($progressBar, 0, 3)
$layout.SetColumnSpan($progressBar, 4)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = 'Ready. Enter a Name or COF NO, then click Search.'
$statusLabel.TextAlign = 'MiddleLeft'
$statusLabel.Dock = 'Fill'
$statusLabel.AutoEllipsis = $true
$layout.Controls.Add($statusLabel, 0, 4)
$layout.SetColumnSpan($statusLabel, 2)
$openButton = New-Object System.Windows.Forms.Button
$openButton.Text = 'Open Row'
$openButton.Dock = 'Fill'
$openButton.Margin = New-Object System.Windows.Forms.Padding(0, 3, 6, 3)
$openButton.Enabled = $false
$layout.Controls.Add($openButton, 2, 4)
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = 'Cancel'
$cancelButton.Dock = 'Fill'
$cancelButton.Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 3)
$cancelButton.Enabled = $false
$layout.Controls.Add($cancelButton, 3, 4)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.BorderStyle = 'Fixed3D'
$grid.ColumnHeadersHeight = 34
$grid.Columns.Add('ClientName', 'Name') | Out-Null
$grid.Columns.Add('CofNo', 'COF NO') | Out-Null
$grid.Columns.Add('MatchedIn', 'Matched in') | Out-Null
$grid.Columns.Add('File', 'Excel file') | Out-Null
$grid.Columns.Add('Sheet', 'Sheet') | Out-Null
$grid.Columns.Add('Row', 'Row') | Out-Null
$grid.Columns['ClientName'].FillWeight = 25
$grid.Columns['CofNo'].FillWeight = 15
$grid.Columns['MatchedIn'].FillWeight = 13
$grid.Columns['File'].FillWeight = 31
$grid.Columns['Sheet'].FillWeight = 12
$grid.Columns['Row'].FillWeight = 6
$layout.Controls.Add($grid, 0, 5)
$layout.SetColumnSpan($grid, 4)

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 150
$timer.Add_Tick({ Read-WorkerMessages $form $searchButton $browseButton $cancelButton $openButton $progressBar $statusLabel $grid $timer })
$searchButton.Add_Click({ Start-Search $folderBox $queryBox $grid $statusLabel $searchButton $browseButton $cancelButton $openButton $progressBar $timer })
$queryBox.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        Start-Search $folderBox $queryBox $grid $statusLabel $searchButton $browseButton $cancelButton $openButton $progressBar $timer
        $_.SuppressKeyPress = $true
    }
})
$openButton.Add_Click({
    if ($grid.SelectedRows.Count -gt 0) { Open-ResultRow $grid.SelectedRows[0].Tag }
    elseif ($grid.CurrentRow) { Open-ResultRow $grid.CurrentRow.Tag }
})
$grid.Add_CellDoubleClick({
    if ($_.RowIndex -ge 0) { Open-ResultRow $grid.Rows[$_.RowIndex].Tag }
})
$cancelButton.Add_Click({
    if ($script:CancelFile) {
        try {
            [System.IO.File]::WriteAllText($script:CancelFile, 'cancel')
            $cancelButton.Enabled = $false
            $statusLabel.Text = 'Cancellation requested. Closing the current workbook...'
        } catch {
            $statusLabel.Text = 'Could not send the cancellation request. The search may continue.'
        }
    }
})
$form.Add_FormClosing({
    if ($null -ne $script:WorkerProcess) {
        try { $script:WorkerProcess.Refresh() } catch { }
        if (-not $script:WorkerProcess.HasExited) {
            $_.Cancel = $true
            $script:CloseWhenDone = $true
            try { [System.IO.File]::WriteAllText($script:CancelFile, 'cancel') } catch { }
            $cancelButton.Enabled = $false
            $statusLabel.Text = 'Cancelling the search before closing the app...'
        } else {
            Remove-WorkerFiles
        }
    }
})
$form.Add_Shown({ $queryBox.Focus() })
[void]$form.ShowDialog()

