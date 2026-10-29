param(
    [string]$ArchivePath = '\\molina.mhc\shares\Diplomat\Operations\BatchResults\all\M360DI\Archive',
    [string]$WorkbookPath = '',
    [string]$SheetName = 'Extracts Outbound & Inbound',
    [int]$StartRow = 68,
    [int]$EndRow = 124,
    [string]$Date = ''
)

$ErrorActionPreference = 'Stop'

if ($StartRow -lt 2 -or $EndRow -lt $StartRow) {
    throw 'The row range is invalid. StartRow must be at least 2 and EndRow must not be less than StartRow.'
}

if ([string]::IsNullOrWhiteSpace($Date)) {
    $Date = (Get-Date).Date.AddDays(-1).ToString('yyyyMMdd')
}
elseif ($Date -notmatch '^\d{8}$') {
    throw 'Date must be in YYYYMMDD format.'
}

if (-not (Test-Path -LiteralPath $ArchivePath -PathType Container)) {
    throw "Archive folder was not found or is not accessible: $ArchivePath"
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    Add-Type -AssemblyName System.Windows.Forms
    $picker = New-Object System.Windows.Forms.OpenFileDialog
    $picker.Title = 'Select the AutoSys daily status workbook'
    $picker.Filter = 'Excel workbooks (*.xlsx)|*.xlsx'
    $picker.Multiselect = $false

    if ($picker.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host 'No workbook selected; nothing was changed.'
        exit 0
    }

    $WorkbookPath = $picker.FileName
}

if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
    throw "Workbook was not found: $WorkbookPath"
}
$inputPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$archiveFiles = @(Get-ChildItem -LiteralPath $ArchivePath -Filter '*.zip' -File)

function Test-IgnoredFile {
    param([string]$FileName)

    $name = $FileName.ToLowerInvariant()
    return $name.Contains('.pdf.') -or $name.Contains('report') -or $name.Contains('summary')
}

function Get-EntryLineCount {
    param(
        [System.IO.Compression.ZipArchiveEntry]$Entry,
        [ValidateSet('All', 'NonBlank')][string]$Mode
    )

    $stream = $null
    $reader = $null
    try {
        $stream = $Entry.Open()
        $reader = New-Object System.IO.StreamReader($stream)
        $count = 0
        while ($null -ne ($line = $reader.ReadLine())) {
            if ($Mode -eq 'All' -or -not [string]::IsNullOrWhiteSpace($line)) {
                $count++
            }
        }
        return $count
    }
    finally {
        if ($null -ne $reader) {
            $reader.Dispose()
        }
        elseif ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Get-ZipTransactionTotal {
    param(
        [System.IO.FileInfo]$ZipFile,
        [ValidateSet('CountFiles', 'CountLines', 'CountLinesWithHeader')][string]$Rule
    )

    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipFile.FullName)
        $total = 0
        foreach ($entry in $zip.Entries) {
            if ([string]::IsNullOrEmpty($entry.Name)) {
                continue
            }

            if ($Rule -eq 'CountFiles') {
                $total++
                continue
            }

            if (Test-IgnoredFile -FileName $entry.Name) {
                continue
            }

            if ($Rule -eq 'CountLinesWithHeader') {
                $lines = Get-EntryLineCount -Entry $entry -Mode NonBlank
                $total += [Math]::Max($lines - 1, 0)
            }
            else {
                $lines = Get-EntryLineCount -Entry $entry -Mode All
                if ($lines -le 3) {
                    $total++
                }
                else {
                    $total += [Math]::Max($lines - 1, 0)
                }
            }
        }
        return $total
    }
    catch {
        Write-Warning "Could not process archive '$($ZipFile.Name)': $($_.Exception.Message)"
        return 0
    }
    finally {
        if ($null -ne $zip) {
            $zip.Dispose()
        }
    }
}

function Get-JobTransactionTotal {
    param([string]$JobName)

    $fixedTotals = @{
        'elig_d_d_ftp_ib_all_mmp_m360_member_exception' = 0
        'elig_d_d_ftp_ob_all_mapd_m360_connecture_dental_oec' = 0
        'elig_d_d_ftp_ib_all_mmp_m360_mbr_comments' = 3
        'elig_d_d_ftp_ib_mmp_rejectinputfile' = 1
    }

    if ($fixedTotals.ContainsKey($JobName)) {
        return [int]$fixedTotals[$JobName]
    }

    $matches = @(
        $script:archiveFiles |
            Where-Object {
                $_.Name.IndexOf($JobName, [StringComparison]::OrdinalIgnoreCase) -ge 0 -and
                $_.Name.IndexOf($script:Date, [StringComparison]::OrdinalIgnoreCase) -ge 0
            } |
            Sort-Object Name
    )

    $fileCountJobs = @(
        'elig_d_d_ftp_ib_all_mmp_m360_834_inc',
        'elig_d_d_ftp_ib_all_mmp_m360_marxcob',
        'elig_d_d_move_ob_all_mmp_m360_marxcob',
        'elig_d_d_ftp_ib_all_mmp_m360_lockbox04',
        'elig_d_d_ftp_ib_all_mmp_m360_dental_response'
    )
    $noHeaderJobs = @(
        'elig_d_d_ib_ftp_mapd_DTRRD',
        'oth_d_d_ib_ftp_infoxing_MARXCOB'
    )
    $lastArchiveOnlyJobs = @(
        'elig_d_d_ftp_ib_all_mmp_m360_extract1',
        'elig_d_d_ftp_ib_all_mmp_m360_extract2',
        'elig_d_d_ftp_ib_all_mmp_m360_agent_exception'
    )

    if ($fileCountJobs -contains $JobName) {
        $rule = 'CountFiles'
    }
    elseif ($JobName -eq 'elig_d_d_ftp_ib_all_mmp_m360_prov_response') {
        $rule = 'CountLinesWithHeader'
    }
    elseif ($noHeaderJobs -contains $JobName) {
        $rule = 'CountWrittenLines'
    }
    else {
        $rule = 'CountLines'
    }

    if ($matches.Count -eq 0) {
        $total = 0
    }
    else {
        if ($lastArchiveOnlyJobs -contains $JobName) {
            $matches = @($matches[-1])
        }

        $total = 0
        foreach ($zipFile in $matches) {
            if ($rule -eq 'CountWrittenLines') {
                $zip = $null
                try {
                    $zip = [System.IO.Compression.ZipFile]::OpenRead($zipFile.FullName)
                    foreach ($entry in $zip.Entries) {
                        if ([string]::IsNullOrEmpty($entry.Name) -or (Test-IgnoredFile -FileName $entry.Name)) {
                            continue
                        }
                        $total += Get-EntryLineCount -Entry $entry -Mode All
                    }
                }
                catch {
                    Write-Warning "Could not process archive '$($zipFile.Name)': $($_.Exception.Message)"
                }
                finally {
                    if ($null -ne $zip) {
                        $zip.Dispose()
                    }
                }
            }
            else {
                $total += Get-ZipTransactionTotal -ZipFile $zipFile -Rule $rule
            }
        }
    }

    if ($JobName -eq 'elig_d_d_ftp_ob_all_mmp_m360_lockbox05' -or
        $JobName -eq 'elig_d_d_ftp_ib_all_mmp_m360_compliance_tracker') {
        $total--
    }

    return [Math]::Max($total, 0)
}

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($inputPath)
    $worksheet = $workbook.Worksheets.Item($SheetName)
    $usedRange = $worksheet.UsedRange
    $lastRow = $usedRange.Row + $usedRange.Rows.Count - 1
    $lastColumn = $usedRange.Column + $usedRange.Columns.Count - 1

    $jobColumn = 0
    $transactionColumn = 0
    for ($row = 1; $row -le [Math]::Min(9, $lastRow); $row++) {
        for ($column = 1; $column -le $lastColumn; $column++) {
            $value = [string]$worksheet.Cells.Item($row, $column).Value2
            if ($value.Trim() -eq 'Job Name') {
                $jobColumn = $column
            }
            if ($value.Contains('No. of Transactions for Yesterday')) {
                $transactionColumn = $column
            }
        }
    }

    if ($jobColumn -eq 0) {
        throw "Could not find the 'Job Name' header in rows 1-9 of worksheet '$SheetName'."
    }
    if ($transactionColumn -eq 0) {
        throw "Could not find the transaction-count header in rows 1-9 of worksheet '$SheetName'."
    }

    for ($row = 2; $row -le $lastRow; $row++) {
        if ($row -lt $StartRow -or $row -gt $EndRow) {
            $worksheet.Cells.Item($row, $transactionColumn).ClearContents()
        }
    }

    Write-Host "`nPROCESSING DATE: $Date"
    for ($row = $StartRow; $row -le $EndRow; $row++) {
        $jobName = [string]$worksheet.Cells.Item($row, $jobColumn).Value2
        if ([string]::IsNullOrWhiteSpace($jobName)) {
            continue
        }

        $jobName = $jobName.Trim()
        $transactionTotal = Get-JobTransactionTotal -JobName $jobName
        $worksheet.Cells.Item($row, $transactionColumn).Value2 = $transactionTotal
        Write-Host ('{0,-60} {1,8}' -f $jobName, $transactionTotal)
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
    $outputPath = Join-Path (Split-Path -Parent $inputPath) ("{0}_transactions_{1}.xlsx" -f $baseName, $Date)
    $workbook.SaveAs($outputPath, 51)
    Write-Host "`nSaved updated workbook to:`n$outputPath"
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
