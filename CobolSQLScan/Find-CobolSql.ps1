#Requires -Version 5.1
# Scans a folder of COBOL sources and writes an MD table: which files contain EXEC SQL.

$ErrorActionPreference = 'Stop'

# ---- settings ----
$Recurse    = $false                 # set to $true to include subfolders
$OutputName = 'SQL_SCAN_RESULT.md'
$SkipExt    = @('.md')               # never scan these (keeps the result file out of the scan)
# ------------------

function Test-CommentLine([string]$Line) {
    # Fixed format: '*' or '/' in column 7 makes the whole line a comment
    if ($Line.Length -ge 7 -and ($Line[6] -eq '*' -or $Line[6] -eq '/')) { return $true }
    # Free format: line starting with '*>'
    if ($Line.TrimStart().StartsWith('*>')) { return $true }
    return $false
}

function Test-HasExecSql([string]$Path) {
    $lines = [System.IO.File]::ReadAllLines($Path)
    $sb = New-Object System.Text.StringBuilder
    foreach ($l in $lines) {
        if (-not (Test-CommentLine $l)) { [void]$sb.AppendLine($l) }
    }
    # \s+ also matches a line break, so "EXEC" and "SQL" on separate lines are found
    return [regex]::IsMatch($sb.ToString(), '\bEXEC\s+SQL\b', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

try {
    Add-Type -AssemblyName System.Windows.Forms

    # Hidden topmost owner so the dialog never opens behind the console window
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost = $true
    $owner.ShowInTaskbar = $false
    $owner.WindowState = 'Minimized'
    $owner.Show()

    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Select the folder that contains the COBOL source files'
    $dlg.ShowNewFolderButton = $false
    $result = $dlg.ShowDialog($owner)
    $owner.Close()

    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host 'No folder selected. Nothing to do.'
        exit 0
    }

    $folder = $dlg.SelectedPath
    Write-Host "Scanning: $folder"

    $files = @(Get-ChildItem -LiteralPath $folder -File -Recurse:$Recurse |
        Where-Object { $SkipExt -notcontains $_.Extension.ToLowerInvariant() } |
        Sort-Object FullName)

    if ($files.Count -eq 0) {
        Write-Host 'No files found in that folder.'
        exit 0
    }

    $rows = New-Object System.Collections.Generic.List[string]
    $rows.Add('| SRC_FILE | SQL_PRESENT? |')
    $rows.Add('|----------|--------------|')

    $yes = 0; $errs = 0
    foreach ($f in $files) {
        $name = $f.FullName.Substring($folder.Length).TrimStart('\', '/')
        try {
            if (Test-HasExecSql $f.FullName) { $flag = 'Y'; $yes++ } else { $flag = 'N' }
        } catch {
            $flag = 'ERR'; $errs++
            Write-Warning "Could not read $name : $($_.Exception.Message)"
        }
        $rows.Add('| ' + $name.Replace('|', '\|') + ' | ' + $flag + ' |')
    }

    $out = Join-Path $folder $OutputName
    $enc = New-Object System.Text.UTF8Encoding($false)
    try {
        [System.IO.File]::WriteAllLines($out, $rows, $enc)
    } catch {
        # Selected folder not writable: fall back to the Desktop
        $out = Join-Path ([Environment]::GetFolderPath('Desktop')) $OutputName
        [System.IO.File]::WriteAllLines($out, $rows, $enc)
    }

    Write-Host ''
    Write-Host "Files scanned : $($files.Count)"
    Write-Host "With EXEC SQL : $yes"
    if ($errs -gt 0) { Write-Host "Unreadable    : $errs" }
    Write-Host "Result file   : $out"

    Start-Process explorer.exe -ArgumentList "/select,`"$out`""
    exit 0
}
catch {
    Write-Host ''
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
