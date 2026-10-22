param(
    [string]$Folder,
    [string]$Query,
    [string]$ProgressFile,
    [string]$CancelFile,
    [string]$CachePath,
    [string]$FolderKey,
    [string]$Signature
)

$ErrorActionPreference = 'Stop'
$script:SearchWhitespaceChars = [char[]]@(9, 10, 13, 160, 8199, 8239, 8203, 65279)

function Send-Message {
    param($Writer, $Message)
    $Writer.WriteLine((ConvertTo-Json -InputObject $Message -Compress -Depth 5))
    $Writer.Flush()
}

function Write-CacheBatch {
    param($State, [object[]]$Rows)
    if ($null -eq $State.Writer -or $Rows.Count -eq 0) { return }
    try {
        $line = ConvertTo-Json -InputObject @{ type = 'batch'; rows = @($Rows) } -Compress -Depth 5
        $State.Writer.WriteLine($line)
    } catch {
        $State.Error = $_.Exception.Message
        try { $State.Writer.Dispose() } catch { }
        $State.Writer = $null
    }
}

function Get-FolderSignature {
    param($Files)
    $parts = foreach ($file in $Files) { "$($file.FullName)`t$($file.Length)`t$($file.LastWriteTimeUtc.Ticks)" }
    return [string]::Join([Environment]::NewLine, [string[]]$parts)
}

function Get-RangeValue {
    param($Matrix, [int]$Row, [int]$Column)
    if ($null -eq $Matrix) { return $null }
    if ($Matrix -isnot [Array] -or $Matrix.Rank -lt 2) { return $Matrix }
    $matrixRow = $Matrix.GetLowerBound(0) + $Row - 1
    $matrixColumn = $Matrix.GetLowerBound(1) + $Column - 1
    return $Matrix.GetValue($matrixRow, $matrixColumn)
}

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

function ConvertTo-HeaderKey {
    param($Value)
    $text = ConvertTo-CellString $Value
    return (($text -replace '[^A-Za-z0-9#]', '').ToUpperInvariant())
}

function Test-NameHeader {
    param([string]$Key)
    if ($Key -in @('NAME', 'CLIENT', 'CLIENTNAME', 'CLIENTSNAME', 'CUSTOMER', 'CUSTOMERNAME', 'CUSTOMERSNAME', 'FULLNAME', 'CLIENTFULLNAME', 'CLIENTNAMEFULL', 'CUSTOMERFULLNAME', 'NAMEOFCLIENT', 'NAMEOFTHECLIENT', 'NAMECLIENT', 'BORROWERNAME', 'MEMBERNAME', 'ACCOUNTNAME', 'CONTACTNAME')) { return $true }
    return $Key -match 'NAME' -and $Key -notmatch '(FILE|WORKBOOK|SHEET|PRODUCT|ITEM)NAME'
}

function Test-CofHeader {
    param([string]$Key)
    return $Key -match '^COF(?:NO|NUMBER|NUM|ID|#)(?:#)?$'
}

function Test-SearchMatch {
    param([string]$Value, [string]$SearchText, [switch]$CofNumber, [AllowNull()][string]$QueryDigits = $null, [AllowNull()][string]$NormalizedQuery = $null)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    if ($null -eq $NormalizedQuery) { $NormalizedQuery = Normalize-SearchText $SearchText }
    $normalizedValue = Normalize-SearchText $Value
    if ($normalizedValue.IndexOf($NormalizedQuery, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    if ($CofNumber) {
        if ($null -eq $QueryDigits) { $QueryDigits = [regex]::Replace($SearchText, '\D', '').TrimStart([char]'0') }
        $valueDigits = [regex]::Replace($Value, '\D', '').TrimStart([char]'0')
        if ($QueryDigits.Length -gt 0 -and $valueDigits.Length -gt 0 -and $valueDigits.Contains($QueryDigits)) { return $true }
    }
    return $false
}

function Get-DisplayCellText {
    param($Worksheet, [int]$Row, [int]$Column, $Fallback)
    if ($Column -le 0) { return '' }
    $cell = $null
    try {
        $cell = $Worksheet.Cells.Item($Row, $Column)
        $text = ([string]$cell.Text).Trim()
        if ($text -and $text -notmatch '^#+$') { return $text }
    } catch {
    } finally {
        if ($null -ne $cell) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell) }
    }
    return (ConvertTo-CellString $Fallback)
}

function Release-ComObject {
    param($Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

function Find-HeaderColumns {
    param($Matrix, [int]$RowCount, [int]$ColumnCount)
    $bestRow = 0
    $bestNameColumn = 0
    $bestCofColumn = 0
    $scanRows = [Math]::Min($RowCount, 100)
    for ($row = 1; $row -le $scanRows; $row++) {
        $nameColumn = 0
        $cofColumn = 0
        for ($column = 1; $column -le $ColumnCount; $column++) {
            $key = ConvertTo-HeaderKey (Get-RangeValue $Matrix $row $column)
            if ($nameColumn -eq 0 -and (Test-NameHeader $key)) { $nameColumn = $column }
            if ($cofColumn -eq 0 -and (Test-CofHeader $key)) { $cofColumn = $column }
        }
        $score = [int]($nameColumn -gt 0) + [int]($cofColumn -gt 0)
        $bestScore = [int]($bestNameColumn -gt 0) + [int]($bestCofColumn -gt 0)
        if ($score -gt $bestScore) {
            $bestRow = $row
            $bestNameColumn = $nameColumn
            $bestCofColumn = $cofColumn
        }
        if ($score -eq 2) { break }
    }
    if ($bestRow -eq 0) { return $null }
    return [pscustomobject]@{ HeaderRow = $bestRow; NameColumn = $bestNameColumn; CofColumn = $bestCofColumn }
}

$writerStream = $null
$writer = $null
$cacheState = @{ Writer = $null; Error = ''; TempPath = "$CachePath.$PID.tmp" }
$cacheSaved = $false
$indexValid = $false
$excel = $null
$errors = New-Object 'System.Collections.Generic.List[string]'
$matchCount = 0
$indexRowCount = 0
$headersCount = 0
$cancelled = $false
$scanFailed = $false
$queryDigits = [regex]::Replace($Query, '\D', '').TrimStart([char]'0')
$normalizedQuery = Normalize-SearchText $Query
$files = @()
try {
    $writerStream = [System.IO.File]::Open($ProgressFile, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
    $writer = [System.IO.StreamWriter]::new($writerStream, [System.Text.UTF8Encoding]::new($false))
    $files = @(Get-ChildItem -LiteralPath $Folder -File -Recurse -ErrorAction SilentlyContinue | Where-Object {
        $_.Extension -in @('.xls', '.xlsx', '.xlsm') -and $_.Name -notlike '~$*'
    } | Sort-Object FullName)
    Send-Message $writer @{ type = 'start'; total = $files.Count }

    if ($files.Count -gt 0) {
        try {
            $cacheDirectory = Split-Path -Parent $CachePath
            [System.IO.Directory]::CreateDirectory($cacheDirectory) | Out-Null
            $cacheState.Writer = [System.IO.StreamWriter]::new($cacheState.TempPath, $false, [System.Text.UTF8Encoding]::new($false))
            $cacheState.Writer.WriteLine((ConvertTo-Json -InputObject @{ type = 'meta'; version = 3; folder = $FolderKey; signature = $Signature; fileCount = $files.Count } -Compress -Depth 5))
        } catch {
            $cacheState.Error = $_.Exception.Message
            if ($null -ne $cacheState.Writer) { try { $cacheState.Writer.Dispose() } catch { }; $cacheState.Writer = $null }
        }
    }

    if ($files.Count -gt 0) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = 3
        try { $excel.Calculation = -4135 } catch { }

        for ($fileIndex = 0; $fileIndex -lt $files.Count; $fileIndex++) {
            if (Test-Path -LiteralPath $CancelFile) { $cancelled = $true; break }
            $file = $files[$fileIndex]
            $fileNumber = $fileIndex + 1
            Send-Message $writer @{ type = 'progress'; current = $fileNumber; total = $files.Count; file = $file.Name }
            $book = $null
            $worksheets = $null
            try {
                $book = $excel.Workbooks.Open($file.FullName, 0, $true)
                $worksheets = $book.Worksheets
                for ($sheetIndex = 1; $sheetIndex -le $worksheets.Count; $sheetIndex++) {
                    $sheet = $null
                    $used = $null
                    try {
                        $sheet = $worksheets.Item($sheetIndex)
                        $used = $sheet.UsedRange
                        $rowCount = [int]$used.Rows.Count
                        $columnCount = [int]$used.Columns.Count
                        $headerRows = [Math]::Min($rowCount, 100)
                        $headerRange = $null
                        try {
                            $headerRange = $used.Resize($headerRows, $columnCount)
                            $headerMatrix = $headerRange.Value2
                            $header = Find-HeaderColumns $headerMatrix $headerRows $columnCount
                        } finally {
                            Release-ComObject $headerRange
                        }
                        if ($null -eq $header) { continue }
                        $headersCount++
                        $firstSheetRow = [int]$used.Row
                        $firstSheetColumn = [int]$used.Column
                        $nameColumn = if ($header.NameColumn -gt 0) { $firstSheetColumn + $header.NameColumn - 1 } else { 0 }
                        $cofColumn = if ($header.CofColumn -gt 0) { $firstSheetColumn + $header.CofColumn - 1 } else { 0 }
                        $indexBatch = New-Object 'System.Collections.Generic.List[object]'
                        $dataRows = $rowCount - $header.HeaderRow
                        $dataStartRow = $firstSheetRow + $header.HeaderRow
                        $nameRange = $null
                        $cofRange = $null
                        $nameStartCell = $null
                        $cofStartCell = $null
                        try {
                            if ($dataRows -gt 0 -and $nameColumn -gt 0) {
                                $nameStartCell = $sheet.Cells.Item($dataStartRow, $nameColumn)
                                $nameRange = $nameStartCell.Resize($dataRows, 1)
                                $nameValues = $nameRange.Value2
                            }
                            if ($dataRows -gt 0 -and $cofColumn -gt 0) {
                                $cofStartCell = $sheet.Cells.Item($dataStartRow, $cofColumn)
                                $cofRange = $cofStartCell.Resize($dataRows, 1)
                                $cofValues = $cofRange.Value2
                            }
                            $nameIsMatrix = $null -ne $nameValues -and $nameValues -is [Array] -and $nameValues.Rank -ge 2
                            $cofIsMatrix = $null -ne $cofValues -and $cofValues -is [Array] -and $cofValues.Rank -ge 2
                            $nameRowBase = 0
                            $nameColumnBase = 0
                            $cofRowBase = 0
                            $cofColumnBase = 0
                            if ($nameIsMatrix) {
                                $nameRowBase = $nameValues.GetLowerBound(0)
                                $nameColumnBase = $nameValues.GetLowerBound(1)
                            }
                            if ($cofIsMatrix) {
                                $cofRowBase = $cofValues.GetLowerBound(0)
                                $cofColumnBase = $cofValues.GetLowerBound(1)
                            }
                            for ($row = 1; $row -le $dataRows; $row++) {
                                if (($row % 500) -eq 0 -and (Test-Path -LiteralPath $CancelFile)) { $cancelled = $true; break }
                                if (($row % 2000) -eq 0) {
                                    Send-Message $writer @{ type = 'detail'; text = "Reading $($file.Name), row $row of $dataRows..." }
                                }
                                if ($nameColumn -gt 0) {
                                    $nameRaw = if ($nameIsMatrix) { $nameValues.GetValue($nameRowBase + $row - 1, $nameColumnBase) } else { $nameValues }
                                    $nameValue = ConvertTo-CellString $nameRaw
                                } else { $nameValue = '' }
                                if ($cofColumn -gt 0) {
                                    $cofRaw = if ($cofIsMatrix) { $cofValues.GetValue($cofRowBase + $row - 1, $cofColumnBase) } else { $cofValues }
                                    $cofValue = ConvertTo-CellString $cofRaw
                                } else { $cofValue = '' }
                                if ($nameValue -or $cofValue) {
                                    $indexBatch.Add([pscustomobject]@{
                                        Name = $nameValue
                                        CofNo = $cofValue
                                        File = $file.Name
                                        Sheet = [string]$sheet.Name
                                        Row = $dataStartRow + $row - 1
                                        Path = $file.FullName
                                        NameColumn = $nameColumn
                                        CofColumn = $cofColumn
                                    })
                                    $indexRowCount++
                                    if ($indexBatch.Count -ge 500) {
                                        Write-CacheBatch $cacheState @($indexBatch.ToArray())
                                        $indexBatch.Clear()
                                    }
                                }
                                $nameMatch = $nameColumn -gt 0 -and (Test-SearchMatch $nameValue $Query -QueryDigits $queryDigits -NormalizedQuery $normalizedQuery)
                                $cofMatch = $cofColumn -gt 0 -and (Test-SearchMatch $cofValue $Query -CofNumber -QueryDigits $queryDigits -NormalizedQuery $normalizedQuery)
                                if (-not $nameMatch -and -not $cofMatch) { continue }

                                $sheetRow = $dataStartRow + $row - 1
                                $sheetName = [string]$sheet.Name
                                $nameText = if ($nameColumn -gt 0) { Get-DisplayCellText $sheet $sheetRow $nameColumn $nameValue } else { '' }
                                $cofText = if ($cofColumn -gt 0) { Get-DisplayCellText $sheet $sheetRow $cofColumn $cofValue } else { '' }
                                $matchedIn = if ($nameMatch -and $cofMatch) { 'Name, COF NO' } elseif ($nameMatch) { 'Name' } else { 'COF NO' }
                                Send-Message $writer @{
                                    type = 'result'
                                    Name = $nameText
                                    CofNo = $cofText
                                    MatchedIn = $matchedIn
                                    File = $file.Name
                                    Sheet = $sheetName
                                    Row = $sheetRow
                                    Path = $file.FullName
                                    NameColumn = $nameColumn
                                    CofColumn = $cofColumn
                                }
                                $matchCount++
                            }
                            if ($indexBatch.Count -gt 0) {
                                Write-CacheBatch $cacheState @($indexBatch.ToArray())
                            }
                        } finally {
                            Release-ComObject $nameRange
                            Release-ComObject $cofRange
                            Release-ComObject $nameStartCell
                            Release-ComObject $cofStartCell
                        }
                        if ($cancelled) { break }
                    } finally {
                        Release-ComObject $used
                        Release-ComObject $sheet
                    }
                }
            } catch {
                $errors.Add("$($file.Name): $($_.Exception.Message)")
                Send-Message $writer @{ type = 'warning'; text = "$($file.Name): $($_.Exception.Message)" }
            } finally {
                if ($null -ne $book) { try { $book.Close($false) } catch { } }
                Release-ComObject $worksheets
                Release-ComObject $book
            }
            Send-Message $writer @{ type = 'fileDone'; current = $fileNumber; total = $files.Count }
            if ($cancelled) { break }
        }
    }
} catch {
    $scanFailed = $true
    Send-Message $writer @{ type = 'error'; text = "The Microsoft Excel scan could not start or finish. $($_.Exception.Message)" }
} finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        Release-ComObject $excel
    }

    if (-not $cancelled -and -not $scanFailed -and $errors.Count -eq 0 -and $files.Count -gt 0) {
        try {
            $currentFiles = @(Get-ChildItem -LiteralPath $Folder -File -Recurse -ErrorAction SilentlyContinue | Where-Object {
                $_.Extension -in @('.xls', '.xlsx', '.xlsm') -and $_.Name -notlike '~$*'
            } | Sort-Object FullName)
            $indexValid = (Get-FolderSignature $currentFiles) -ceq $Signature
        } catch {
            $indexValid = $false
        }
    }

    if ($null -ne $cacheState.Writer) {
        try {
            if ($indexValid -and -not $cacheState.Error) {
                $cacheState.Writer.WriteLine((ConvertTo-Json -InputObject @{ type = 'stats'; headers = $headersCount; rows = $indexRowCount } -Compress -Depth 5))
                $cacheState.Writer.Flush()
                $cacheState.Writer.Dispose()
                $cacheState.Writer = $null
                if ([System.IO.File]::Exists($CachePath)) {
                    [System.IO.File]::Replace($cacheState.TempPath, $CachePath, $null)
                } else {
                    [System.IO.File]::Move($cacheState.TempPath, $CachePath)
                }
                $cacheSaved = $true
            } else {
                $cacheState.Writer.Dispose()
                $cacheState.Writer = $null
            }
        } catch {
            $cacheState.Error = $_.Exception.Message
            if ($null -ne $cacheState.Writer) { try { $cacheState.Writer.Dispose() } catch { }; $cacheState.Writer = $null }
        }
    }
    if (-not $cacheSaved -and $cacheState.TempPath -and [System.IO.File]::Exists($cacheState.TempPath)) {
        try { [System.IO.File]::Delete($cacheState.TempPath) } catch { }
    }
    if ($cacheState.Error) {
        Send-Message $writer @{ type = 'cacheWarning'; text = "The local search index could not be saved. Search will still work, but the full scan may be needed the next time the app starts. $($cacheState.Error)" }
    } elseif (-not $indexValid -and -not $cancelled -and -not $scanFailed -and $errors.Count -eq 0 -and $files.Count -gt 0) {
        Send-Message $writer @{ type = 'cacheWarning'; text = 'An Excel file changed while the index was being created. Search again to build a fresh index.' }
    }
    if ($null -ne $writer) {
        Send-Message $writer @{ type = 'done'; files = $files.Count; matches = $matchCount; headers = $headersCount; errors = $errors.Count; cancelled = $cancelled; cacheSaved = $cacheSaved; indexValid = $indexValid }
        $writer.Dispose()
    } elseif ($null -ne $writerStream) {
        $writerStream.Dispose()
    }
}

