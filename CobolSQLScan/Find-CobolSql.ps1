# Scans a folder of COBOL sources and writes an MD table: which files contain EXEC SQL.
# Start it via Start-CobolSqlScan.bat. In the window:
#   1. pick the source folder (Browse button or drag and drop)
#   2. optionally pick the output folder (default: the source folder)
#   3. click Scan

$ErrorActionPreference = 'Stop'
$global:ErrorActionPreference = 'Stop'

# ---- settings ----
$global:Recurse    = $false              # set to $true to include subfolders
$global:Extensions = @('.txt', '.cbl')   # only files with these extensions are scanned
# ------------------

# Everything the GUI handlers use is global: handlers run later, and global scope is always reachable.

function global:Test-CommentLine([string]$Line) {
    # Fixed format: '*' or '/' in column 7 makes the whole line a comment
    if ($Line.Length -ge 7 -and ($Line[6] -eq '*' -or $Line[6] -eq '/')) { return $true }
    # Free format: line starting with '*>'
    if ($Line.TrimStart().StartsWith('*>')) { return $true }
    return $false
}

function global:Test-HasExecSql([string]$Path) {
    $lines = [System.IO.File]::ReadAllLines($Path)
    $sb = New-Object System.Text.StringBuilder
    foreach ($l in $lines) {
        if (-not (Test-CommentLine $l)) { [void]$sb.AppendLine($l) }
    }
    # \s+ also matches a line break, so "EXEC" and "SQL" on separate lines are found
    return [regex]::IsMatch($sb.ToString(), '\bEXEC\s+SQL\b', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

# Core: scan the folder and write the MD file. No GUI code in here.
# Returns NoFiles = $true when no matching files exist, otherwise the result details.
function global:Invoke-CobolScan {
    param(
        [string]$Folder,
        [string]$OutDir
    )

    if (-not $Folder -or -not (Test-Path -LiteralPath $Folder -PathType Container)) {
        throw "The source folder does not exist:`n$Folder"
    }
    if ([string]::IsNullOrEmpty($OutDir)) { $OutDir = $Folder }
    if (-not (Test-Path -LiteralPath $OutDir -PathType Container)) {
        throw "The output folder does not exist:`n$OutDir"
    }

    $recurse = [bool]$global:Recurse
    $exts    = @($global:Extensions)

    $files = @(Get-ChildItem -LiteralPath $Folder -File -Recurse:$recurse |
        Where-Object { $exts -contains $_.Extension.ToLowerInvariant() } |
        Sort-Object FullName)

    if ($files.Count -eq 0) {
        return [pscustomobject]@{ NoFiles = $true; Exts = ($exts -join ' or '); Recurse = $recurse }
    }

    # One row per base name (extension removed). If the same base name occurs more than once
    # (e.g. A.cbl and A.txt), the rows merge: Y wins, then ERR, then N.
    $status = @{}
    $errs = 0
    foreach ($f in $files) {
        $rel  = $f.FullName.Substring($Folder.Length).TrimStart('\', '/')
        $name = $rel.Substring(0, $rel.Length - $f.Extension.Length)
        if ($name -eq '') { $name = $rel }
        try {
            if (Test-HasExecSql $f.FullName) { $flag = 'Y' } else { $flag = 'N' }
        } catch {
            $flag = 'ERR'
            $errs++
            Write-Warning "Could not read $rel : $($_.Exception.Message)"
        }
        if (-not $status.ContainsKey($name)) {
            $status[$name] = $flag
        } elseif ($flag -eq 'Y' -or ($flag -eq 'ERR' -and $status[$name] -eq 'N')) {
            $status[$name] = $flag
        }
    }

    $rows = New-Object System.Collections.Generic.List[string]
    $rows.Add('| SRC_FILE | SQL_PRESENT? |')
    $rows.Add('|----------|--------------|')
    $yes = 0
    foreach ($name in ($status.Keys | Sort-Object)) {
        $flag = $status[$name]
        if ($flag -eq 'Y') { $yes++ }
        $rows.Add('| ' + $name.Replace('|', '\|') + ' | ' + $flag + ' |')
    }

    $outName = 'SQL_SCAN_RESULT_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.md'
    $out = Join-Path $OutDir $outName
    try {
        [System.IO.File]::WriteAllLines($out, $rows, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        throw "Could not save the result file in:`n$OutDir`n$($_.Exception.Message)`nChoose another output folder."
    }

    return [pscustomobject]@{ NoFiles = $false; OutPath = $out; Files = $files.Count; Unique = $status.Count; Yes = $yes; Errors = $errs }
}

# ==== GUI ====

function global:Set-ScanStatus([string]$Text, [string]$Kind, [switch]$Pump) {
    $lbl = $global:UI.Status
    if ($Kind -eq 'error') { $lbl.ForeColor = [System.Drawing.Color]::Firebrick }
    elseif ($Kind -eq 'ok') { $lbl.ForeColor = [System.Drawing.Color]::DarkGreen }
    else { $lbl.ForeColor = [System.Drawing.Color]::DimGray }
    $lbl.Text = $Text
    if ($Pump) { [System.Windows.Forms.Application]::DoEvents() }
}

function global:Update-OutDirBox {
    $ui = $global:UI
    if ([string]::IsNullOrEmpty($ui.OutDir)) { $ui.OutBox.Text = '(same folder as the source folder)' }
    else { $ui.OutBox.Text = $ui.OutDir }
}

function global:Set-InputFolder([string]$Path) {
    $ui = $global:UI
    $ui.InputFolder = $Path
    $name = Split-Path -Leaf $Path
    if (-not $name) { $name = $Path }
    $ui.DropText = $name + "`n(drop another folder here to replace it)"
    $ui.Drop.Invalidate()
    $ui.Scan.Enabled = $true
    Set-ScanStatus ("Selected: $Path`nNow check the output folder, then click Scan.") 'info'
}

function global:Start-Scan {
    $ui = $global:UI
    $ui.Scan.Enabled = $false
    $ui.Browse.Enabled = $false
    $ui.Drop.AllowDrop = $false
    $ui.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Set-ScanStatus 'Scanning...' 'info' -Pump
        $res = Invoke-CobolScan -Folder $ui.InputFolder -OutDir $ui.OutDir
        if ($res.NoFiles) {
            $msg = "No $($res.Exts) files found in:`n$($ui.InputFolder)"
            if (-not $res.Recurse) { $msg += "`n(Subfolders are not searched.)" }
            Set-ScanStatus $msg 'error'
        } else {
            $msg = "Done. Files scanned: $($res.Files), unique names: $($res.Unique), with EXEC SQL: $($res.Yes)"
            if ($res.Errors -gt 0) { $msg += ", unreadable: $($res.Errors)" }
            Set-ScanStatus ($msg + "`nSaved: " + $res.OutPath) 'ok'
            Start-Process explorer.exe -ArgumentList "/select,`"$($res.OutPath)`""
        }
    }
    catch {
        Set-ScanStatus ('ERROR: ' + $_.Exception.Message) 'error'
    }
    finally {
        $ui.Form.Cursor = [System.Windows.Forms.Cursors]::Default
        $ui.Browse.Enabled = $true
        $ui.Drop.AllowDrop = $true
        $ui.Scan.Enabled = $true
    }
}

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $global:UI = @{}
    $global:UI.PanelColor  = [System.Drawing.Color]::FromArgb(248, 248, 248)
    $global:UI.HoverColor  = [System.Drawing.Color]::FromArgb(221, 235, 247)
    $global:UI.InputFolder = ''
    $global:UI.OutDir      = ''
    $global:UI.DropText    = 'Drop a folder here'

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'COBOL SQL Scan'
    $form.ClientSize = New-Object System.Drawing.Size(540, 500)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.TopMost = $true
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Select the folder that contains your COBOL source files (.txt / .cbl)'
    $title.Location = New-Object System.Drawing.Point(20, 15)
    $title.Size = New-Object System.Drawing.Size(500, 22)
    $title.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($title)

    $opt1 = New-Object System.Windows.Forms.Label
    $opt1.Text = 'Option 1: click Browse and pick the folder'
    $opt1.Location = New-Object System.Drawing.Point(20, 50)
    $opt1.Size = New-Object System.Drawing.Size(500, 20)
    $form.Controls.Add($opt1)

    $browse = New-Object System.Windows.Forms.Button
    $browse.Text = 'Browse...'
    $browse.Location = New-Object System.Drawing.Point(20, 74)
    $browse.Size = New-Object System.Drawing.Size(160, 32)
    $form.Controls.Add($browse)

    $opt2 = New-Object System.Windows.Forms.Label
    $opt2.Text = 'Option 2: drag and drop the folder into the box below'
    $opt2.Location = New-Object System.Drawing.Point(20, 125)
    $opt2.Size = New-Object System.Drawing.Size(500, 20)
    $form.Controls.Add($opt2)

    $drop = New-Object System.Windows.Forms.Panel
    $drop.Location = New-Object System.Drawing.Point(20, 149)
    $drop.Size = New-Object System.Drawing.Size(500, 110)
    $drop.BackColor = $global:UI.PanelColor
    $drop.AllowDrop = $true
    $form.Controls.Add($drop)

    $outLbl = New-Object System.Windows.Forms.Label
    $outLbl.Text = 'Output folder for the result file (default: the source folder)'
    $outLbl.Location = New-Object System.Drawing.Point(20, 272)
    $outLbl.Size = New-Object System.Drawing.Size(500, 20)
    $form.Controls.Add($outLbl)

    $outBox = New-Object System.Windows.Forms.TextBox
    $outBox.ReadOnly = $true
    $outBox.Location = New-Object System.Drawing.Point(20, 297)
    $outBox.Size = New-Object System.Drawing.Size(300, 24)
    $form.Controls.Add($outBox)

    $outBrowse = New-Object System.Windows.Forms.Button
    $outBrowse.Text = 'Choose...'
    $outBrowse.Location = New-Object System.Drawing.Point(328, 295)
    $outBrowse.Size = New-Object System.Drawing.Size(90, 28)
    $form.Controls.Add($outBrowse)

    $outReset = New-Object System.Windows.Forms.Button
    $outReset.Text = 'Same as input'
    $outReset.Location = New-Object System.Drawing.Point(426, 295)
    $outReset.Size = New-Object System.Drawing.Size(94, 28)
    $form.Controls.Add($outReset)

    $scan = New-Object System.Windows.Forms.Button
    $scan.Text = 'Scan'
    $scan.Enabled = $false
    $scan.Location = New-Object System.Drawing.Point(20, 338)
    $scan.Size = New-Object System.Drawing.Size(500, 38)
    $scan.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($scan)

    $status = New-Object System.Windows.Forms.Label
    $status.Text = 'Step 1: choose the source folder with Browse or drag and drop.'
    $status.ForeColor = [System.Drawing.Color]::DimGray
    $status.Location = New-Object System.Drawing.Point(20, 388)
    $status.Size = New-Object System.Drawing.Size(500, 95)
    $form.Controls.Add($status)

    $global:UI.Form      = $form
    $global:UI.Browse    = $browse
    $global:UI.Drop      = $drop
    $global:UI.OutBox    = $outBox
    $global:UI.Scan      = $scan
    $global:UI.Status    = $status

    $drop.Add_Paint({
        param($s, $e)
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Gray, 2)
        $pen.DashStyle = [System.Drawing.Drawing2D.DashStyle]::Dash
        $rect = New-Object System.Drawing.Rectangle(4, 4, ($s.Width - 9), ($s.Height - 9))
        $e.Graphics.DrawRectangle($pen, $rect)
        $pen.Dispose()
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = [System.Drawing.StringAlignment]::Center
        $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
        $sf.Trimming = [System.Drawing.StringTrimming]::EllipsisPath
        $area = New-Object System.Drawing.RectangleF(8, 8, ($s.Width - 16), ($s.Height - 16))
        if ($global:UI.InputFolder) { $brush = [System.Drawing.Brushes]::DarkSlateGray } else { $brush = [System.Drawing.Brushes]::DimGray }
        $e.Graphics.DrawString($global:UI.DropText, $s.Font, $brush, $area, $sf)
        $sf.Dispose()
    })

    $drop.Add_DragEnter({
        param($s, $e)
        if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
            $s.BackColor = $global:UI.HoverColor
        } else {
            $e.Effect = [System.Windows.Forms.DragDropEffects]::None
        }
    })

    $drop.Add_DragOver({
        param($s, $e)
        if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        } else {
            $e.Effect = [System.Windows.Forms.DragDropEffects]::None
        }
    })

    $drop.Add_DragLeave({
        param($s, $e)
        $s.BackColor = $global:UI.PanelColor
    })

    $drop.Add_DragDrop({
        param($s, $e)
        $s.BackColor = $global:UI.PanelColor
        $items = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        $dir = $null
        foreach ($i in $items) {
            if (Test-Path -LiteralPath $i -PathType Container) { $dir = $i; break }
        }
        if ($dir) {
            Set-InputFolder $dir
        } else {
            Set-ScanStatus 'That is a file, not a folder. Please drop a folder.' 'error'
        }
    })

    $browse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Select the folder that contains the COBOL source files'
        $dlg.ShowNewFolderButton = $false
        if ($dlg.ShowDialog($global:UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) {
            Set-InputFolder $dlg.SelectedPath
        }
        $dlg.Dispose()
    })

    $outBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Select the folder where the result file will be saved'
        $dlg.ShowNewFolderButton = $true
        if ($global:UI.OutDir) { $dlg.SelectedPath = $global:UI.OutDir }
        elseif ($global:UI.InputFolder) { $dlg.SelectedPath = $global:UI.InputFolder }
        if ($dlg.ShowDialog($global:UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) {
            $global:UI.OutDir = $dlg.SelectedPath
            Update-OutDirBox
        }
        $dlg.Dispose()
    })

    $outReset.Add_Click({
        $global:UI.OutDir = ''
        Update-OutDirBox
    })

    $scan.Add_Click({ Start-Scan })

    $form.Add_Shown({
        param($s, $e)
        $s.Activate()
    })

    Update-OutDirBox

    [void]$form.ShowDialog()
    exit 0
}
catch {
    Write-Host ''
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
