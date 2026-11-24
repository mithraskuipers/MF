# Scans a folder of COBOL sources and writes an MD table: which files contain EXEC SQL.
# Start it via Start-CobolSqlScan.bat. A window offers two ways to pick the folder:
# a Browse button, or drag and drop of a folder onto the box.

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

function global:Set-ScanStatus([string]$Text, [string]$Kind) {
    $lbl = $global:UI.Status
    if ($Kind -eq 'error') { $lbl.ForeColor = [System.Drawing.Color]::Firebrick }
    elseif ($Kind -eq 'ok') { $lbl.ForeColor = [System.Drawing.Color]::DarkGreen }
    else { $lbl.ForeColor = [System.Drawing.Color]::DimGray }
    $lbl.Text = $Text
    [System.Windows.Forms.Application]::DoEvents()
}

function global:Invoke-FolderScan([string]$Folder) {
    $ui = $global:UI
    $ui.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $ui.Browse.Enabled = $false
    $ui.Drop.AllowDrop = $false
    try {
        $recurse = [bool]$global:Recurse
        $exts    = @($global:Extensions)
        $extText = $exts -join ' or '

        $files = @(Get-ChildItem -LiteralPath $Folder -File -Recurse:$recurse |
            Where-Object { $exts -contains $_.Extension.ToLowerInvariant() } |
            Sort-Object FullName)

        if ($files.Count -eq 0) {
            $msg = "No $extText files found in:`n$Folder"
            if (-not $recurse) { $msg += "`n(Subfolders are not searched.)" }
            Set-ScanStatus $msg 'error'
            return
        }

        $rows = New-Object System.Collections.Generic.List[string]
        $rows.Add('| SRC_FILE | SQL_PRESENT? |')
        $rows.Add('|----------|--------------|')

        # One row per base name (extension removed). If the same base name occurs more than once
        # (e.g. A.cbl and A.txt), the rows merge: Y wins, then ERR, then N.
        $status = @{}
        foreach ($f in $files) {
            $rel  = $f.FullName.Substring($Folder.Length).TrimStart('\', '/')
            $name = $rel.Substring(0, $rel.Length - $f.Extension.Length)
            if ($name -eq '') { $name = $rel }
            try {
                if (Test-HasExecSql $f.FullName) { $flag = 'Y' } else { $flag = 'N' }
            } catch {
                $flag = 'ERR'
                Write-Warning "Could not read $rel : $($_.Exception.Message)"
            }
            if (-not $status.ContainsKey($name)) {
                $status[$name] = $flag
            } elseif ($flag -eq 'Y' -or ($flag -eq 'ERR' -and $status[$name] -eq 'N')) {
                $status[$name] = $flag
            }
        }

        $yes = 0
        foreach ($name in ($status.Keys | Sort-Object)) {
            $flag = $status[$name]
            if ($flag -eq 'Y') { $yes++ }
            $rows.Add('| ' + $name.Replace('|', '\|') + ' | ' + $flag + ' |')
        }

        $outName = 'SQL_SCAN_RESULT_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.md'
        $out = Join-Path $Folder $outName
        $enc = New-Object System.Text.UTF8Encoding($false)
        try {
            [System.IO.File]::WriteAllLines($out, $rows, $enc)
        } catch {
            # Selected folder not writable: fall back to the Desktop
            $out = Join-Path ([Environment]::GetFolderPath('Desktop')) $outName
            [System.IO.File]::WriteAllLines($out, $rows, $enc)
        }

        Set-ScanStatus ("Done. Files scanned: $($files.Count), unique names: $($status.Count), with EXEC SQL: $yes`nSaved: $out") 'ok'
        Start-Process explorer.exe -ArgumentList "/select,`"$out`""
    }
    catch {
        Set-ScanStatus ('ERROR: ' + $_.Exception.Message) 'error'
    }
    finally {
        $ui.Form.Cursor = [System.Windows.Forms.Cursors]::Default
        $ui.Browse.Enabled = $true
        $ui.Drop.AllowDrop = $true
    }
}

function global:Request-ScanFolder([string]$Folder) {
    # Deferred via a short timer so a drop does not keep the dragging app (Explorer) waiting
    $global:UI.Pending = $Folder
    Set-ScanStatus "Scanning: $Folder" 'info'
    $global:UI.Timer.Start()
}

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $global:UI = @{}
    $global:UI.PanelColor = [System.Drawing.Color]::FromArgb(248, 248, 248)
    $global:UI.HoverColor = [System.Drawing.Color]::FromArgb(221, 235, 247)
    $global:UI.Pending = ''

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'COBOL SQL Scan'
    $form.ClientSize = New-Object System.Drawing.Size(540, 400)
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
    $drop.Size = New-Object System.Drawing.Size(500, 140)
    $drop.BackColor = $global:UI.PanelColor
    $drop.AllowDrop = $true
    $form.Controls.Add($drop)

    $status = New-Object System.Windows.Forms.Label
    $status.Text = 'Waiting for a folder.'
    $status.ForeColor = [System.Drawing.Color]::DimGray
    $status.Location = New-Object System.Drawing.Point(20, 300)
    $status.Size = New-Object System.Drawing.Size(500, 85)
    $form.Controls.Add($status)

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 100

    $global:UI.Form   = $form
    $global:UI.Browse = $browse
    $global:UI.Drop   = $drop
    $global:UI.Status = $status
    $global:UI.Timer  = $timer

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
        $area = New-Object System.Drawing.RectangleF(0, 0, $s.Width, $s.Height)
        $e.Graphics.DrawString('Drop a folder here', $s.Font, [System.Drawing.Brushes]::DimGray, $area, $sf)
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
            Request-ScanFolder $dir
        } else {
            Set-ScanStatus 'That is a file, not a folder. Please drop a folder.' 'error'
        }
    })

    $browse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Select the folder that contains the COBOL source files'
        $dlg.ShowNewFolderButton = $false
        if ($dlg.ShowDialog($global:UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) {
            Request-ScanFolder $dlg.SelectedPath
        }
        $dlg.Dispose()
    })

    $timer.Add_Tick({
        param($s, $e)
        $s.Stop()
        try { Invoke-FolderScan $global:UI.Pending }
        catch { Set-ScanStatus ('ERROR: ' + $_.Exception.Message) 'error' }
    })

    $form.Add_Shown({
        param($s, $e)
        $s.Activate()
    })

    [void]$form.ShowDialog()
    exit 0
}
catch {
    Write-Host ''
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
