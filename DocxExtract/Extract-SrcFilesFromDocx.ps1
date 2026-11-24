# Extracts mentions of source file names (or any pattern) from a Word .docx document
# and lists them in an MD table (column SRC_FILE).
# Start it via Start-DocxExtract.bat. Pick the .docx via Browse or drag and drop,
# set the search pattern and output folder, then click Extract.

$ErrorActionPreference = 'Stop'
$global:ErrorActionPreference = 'Stop'

# Everything the GUI handlers use is global: handlers run later, and global scope is always reachable.

# ---------------------------------------------------------------------------
# Simple pattern language
#   #      one digit
#   @      one letter
#   *      letters/digits/_ (1 or more)
#   ?      previous item is optional          e.g. @?
#   {3}    previous item exactly 3 times      e.g. #{3}
#   {1-3}  previous item 1 to 3 times         {2-} = 2 or more
#   \      next character is taken literally  e.g. \# or \*
#   ,      separates several patterns         e.g. CT###,AB###
#   spaces are ignored (use \  for a literal space)
# ---------------------------------------------------------------------------
function global:ConvertTo-SimpleRegex([string]$Text) {
    $alts = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    $canQuant = $false
    $n = $Text.Length
    $i = 0
    while ($i -lt $n) {
        $c = [string]$Text[$i]
        if ([char]::IsWhiteSpace($Text[$i])) { $i++; continue }
        if ($c -eq '\') {
            if ($i + 1 -ge $n) { throw 'The pattern ends with a lone backslash (\).' }
            [void]$sb.Append([regex]::Escape([string]$Text[$i + 1]))
            $canQuant = $true
            $i += 2
            continue
        }
        if ($c -eq '#') { [void]$sb.Append('[0-9]'); $canQuant = $true }
        elseif ($c -eq '@') { [void]$sb.Append('[A-Za-z]'); $canQuant = $true }
        elseif ($c -eq '*') { [void]$sb.Append('(?:[A-Za-z0-9_]+)'); $canQuant = $true }
        elseif ($c -eq '?') {
            if (-not $canQuant) { throw "A ? must come right after the item it makes optional, for example @?" }
            [void]$sb.Append('?')
            $canQuant = $false
        }
        elseif ($c -eq '{') {
            if (-not $canQuant) { throw "A count like {3} or {1-3} must come right after an item, for example #{3}" }
            $close = $Text.IndexOf('}', $i)
            if ($close -lt 0) { throw 'Missing closing } in the pattern.' }
            $inner = $Text.Substring($i + 1, $close - $i - 1).Trim()
            if ($inner -match '^\d{1,2}$') {
                $q = '{' + [int]$inner + '}'
            }
            elseif ($inner -match '^(\d{1,2})-(\d{1,2})$') {
                $lo = [int]$Matches[1]
                $hi = [int]$Matches[2]
                if ($hi -lt $lo) { throw "In {$inner} the second number must not be smaller than the first." }
                $q = '{' + $lo + ',' + $hi + '}'
            }
            elseif ($inner -match '^(\d{1,2})-$') {
                $q = '{' + [int]$Matches[1] + ',}'
            }
            else { throw 'Use {3} (exactly 3), {1-3} (1 to 3) or {2-} (2 or more) for counts.' }
            [void]$sb.Append($q)
            $canQuant = $false
            $i = $close + 1
            continue
        }
        elseif ($c -eq '}') { throw 'Unexpected } in the pattern.' }
        elseif ($c -eq ',') {
            if ($sb.Length -gt 0) { $alts.Add($sb.ToString()); [void]$sb.Clear() }
            $canQuant = $false
        }
        else { [void]$sb.Append([regex]::Escape($c)); $canQuant = $true }
        $i++
    }
    if ($sb.Length -gt 0) { $alts.Add($sb.ToString()) }
    if ($alts.Count -eq 0) { throw 'Type a search pattern first, for example CT###' }
    return '(?:' + ($alts -join '|') + ')'
}

function global:Get-SearchRegex([string]$Pattern, [bool]$Raw, [bool]$Whole, [bool]$Case) {
    if ($Raw) {
        if ([string]::IsNullOrWhiteSpace($Pattern)) { throw 'Type a regular expression first.' }
        $rx = $Pattern
    } else {
        $rx = ConvertTo-SimpleRegex $Pattern
        if ($Whole) { $rx = '(?<![A-Za-z0-9_])' + $rx + '(?![A-Za-z0-9_])' }
    }
    $opts = [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    if (-not $Case) {
        $opts = [System.Text.RegularExpressions.RegexOptions]($opts -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }
    try {
        $regex = New-Object -TypeName System.Text.RegularExpressions.Regex -ArgumentList @($rx, $opts, [TimeSpan]::FromSeconds(10))
    } catch {
        $m = $_.Exception.Message
        if ($_.Exception.InnerException) { $m = $_.Exception.InnerException.Message }
        throw "Invalid regular expression: $m"
    }
    return [pscustomobject]@{ Text = $rx; Regex = $regex }
}

function global:Format-Code([string]$Value) {
    $bt = [string][char]96
    if ($Value.Contains($bt)) { return $bt + $bt + ' ' + $Value + ' ' + $bt + $bt }
    return $bt + $Value + $bt
}

# Core: read the .docx, find matches, write the MD file. Returns a result object. No GUI code in here.
function global:Invoke-DocxExtraction {
    param(
        [string]$InFile,
        [string]$OutDir,
        [string]$Pattern,
        [bool]$Raw,
        [bool]$Whole,
        [bool]$Case
    )

    if (-not $InFile -or -not (Test-Path -LiteralPath $InFile -PathType Leaf)) {
        throw 'The selected .docx file no longer exists. Please choose it again.'
    }
    $search = Get-SearchRegex $Pattern $Raw $Whole $Case
    $regex = $search.Regex

    if ([string]::IsNullOrEmpty($OutDir)) { $OutDir = Split-Path -Parent $InFile }
    if (-not (Test-Path -LiteralPath $OutDir -PathType Container)) {
        throw "The output folder does not exist:`n$OutDir"
    }

    Add-Type -AssemblyName System.IO.Compression

    $fs = $null
    $zip = $null
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([System.StringComparer]::Ordinal)
    $found = New-Object System.Collections.Generic.List[string]
    $total = 0
    $partCount = 0
    try {
        # FileShare.ReadWrite so a document that is still open in Word can be read
        $fs = New-Object System.IO.FileStream($InFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        } catch {
            throw 'This file could not be opened as a .docx (it may be corrupted, password protected, or an old .doc file).'
        }

        $parts = @($zip.Entries | Where-Object { $_.FullName -match '^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$' })
        $hasBody = $false
        foreach ($p in $parts) { if ($p.FullName -ieq 'word/document.xml') { $hasBody = $true } }
        if (-not $hasBody) { throw 'No word/document.xml found. This is not a regular Word (.docx) document.' }
        $partCount = $parts.Count

        $sb = New-Object System.Text.StringBuilder
        foreach ($part in $parts) {
            $xml = New-Object System.Xml.XmlDocument
            $st = $part.Open()
            try {
                $rs = New-Object System.Xml.XmlReaderSettings
                $rs.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
                $reader = [System.Xml.XmlReader]::Create($st, $rs)
                try { $xml.Load($reader) } finally { $reader.Dispose() }
            } finally { $st.Dispose() }

            $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

            foreach ($para in $xml.SelectNodes('//w:p', $ns)) {
                [void]$sb.Clear()
                # Runs of one paragraph are joined, so a name split over several runs is still found
                foreach ($node in $para.SelectNodes('.//w:t | .//w:tab | .//w:br | .//w:cr | .//w:noBreakHyphen', $ns)) {
                    $ln = $node.LocalName
                    if ($ln -eq 't') { [void]$sb.Append($node.InnerText) }
                    elseif ($ln -eq 'noBreakHyphen') { [void]$sb.Append('-') }
                    else { [void]$sb.Append(' ') }
                }
                if ($sb.Length -eq 0) { continue }
                try {
                    $hits = $regex.Matches($sb.ToString())
                } catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
                    throw 'The search took too long. Please simplify the regular expression.'
                }
                foreach ($m in $hits) {
                    $total++
                    $key = $m.Value
                    if (-not $Case) { $key = $key.ToUpperInvariant() }
                    if ($seen.Add($key)) { $found.Add($m.Value) }
                }
            }
        }
    }
    finally {
        if ($zip) { $zip.Dispose() }
        if ($fs) { $fs.Dispose() }
    }

    # Natural sort: CT9 comes before CT394
    $pad = [System.Text.RegularExpressions.MatchEvaluator]{ param($mm) $mm.Value.PadLeft(12, '0') }
    $sorted = @($found | Sort-Object { [regex]::Replace($_, '\d+', $pad) })

    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $now = Get-Date
    $stamp = $now.ToString('yyyyMMdd_HHmmss', $inv)
    $nowText = $now.ToString('yyyy-MM-dd HH:mm:ss', $inv)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($InFile)
    $out = Join-Path $OutDir ($base + '_SRC_FILE_' + $stamp + '.md')

    if ($Raw) { $modeText = 'Regular expression (used exactly as typed)' } else { $modeText = 'Simple pattern' }
    if ($Raw) { $wholeText = 'Not applicable (regular expression mode)' } elseif ($Whole) { $wholeText = 'Yes' } else { $wholeText = 'No' }
    if ($Case) { $caseText = 'Yes' } else { $caseText = 'No' }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('# SRC_FILE extraction')
    $lines.Add('')
    $lines.Add('## Search settings')
    $lines.Add('')
    $lines.Add('- Source document: ' + (Format-Code ([System.IO.Path]::GetFileName($InFile))))
    $lines.Add('- Source path: ' + (Format-Code $InFile))
    $lines.Add('- Run date and time: ' + $nowText)
    $lines.Add('- Pattern entered: ' + (Format-Code $Pattern))
    $lines.Add('- Pattern mode: ' + $modeText)
    $lines.Add('- Regular expression used: ' + (Format-Code $search.Text))
    $lines.Add('- Case sensitive: ' + $caseText)
    $lines.Add('- Whole words only: ' + $wholeText)
    $lines.Add('- Document parts scanned: ' + $partCount + ' (body, plus headers, footers, footnotes and endnotes where present)')
    $lines.Add('- Unique matches: ' + $sorted.Count + ' (' + $total + ' occurrences in total)')
    $lines.Add('')
    $lines.Add('## Matches')
    $lines.Add('')
    $lines.Add('| SRC_FILE |')
    $lines.Add('|----------|')
    foreach ($s in $sorted) { $lines.Add('| ' + $s.Replace('|', '\|') + ' |') }

    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($out, $lines, $enc)

    return [pscustomobject]@{ OutPath = $out; Unique = $sorted.Count; Total = $total; Parts = $partCount; Matches = $sorted }
}

# ==== GUI ====

function global:Set-DocStatus([string]$Text, [string]$Kind, [switch]$Pump) {
    $lbl = $global:UI.Status
    if ($Kind -eq 'error') { $lbl.ForeColor = [System.Drawing.Color]::Firebrick }
    elseif ($Kind -eq 'ok') { $lbl.ForeColor = [System.Drawing.Color]::DarkGreen }
    elseif ($Kind -eq 'warn') { $lbl.ForeColor = [System.Drawing.Color]::DarkOrange }
    else { $lbl.ForeColor = [System.Drawing.Color]::DimGray }
    $lbl.Text = $Text
    if ($Pump) { [System.Windows.Forms.Application]::DoEvents() }
}

function global:Update-RegexPreview {
    $ui = $global:UI
    try {
        $r = Get-SearchRegex $ui.Combo.Text $ui.Raw.Checked $ui.Whole.Checked $ui.Case.Checked
        $ui.Preview.ForeColor = [System.Drawing.Color]::DimGray
        $ui.Preview.Text = 'Regex used: ' + $r.Text
    } catch {
        $ui.Preview.ForeColor = [System.Drawing.Color]::Firebrick
        $ui.Preview.Text = 'Pattern problem: ' + $_.Exception.Message
    }
}

function global:Update-OutDirBox {
    $ui = $global:UI
    if ([string]::IsNullOrEmpty($ui.OutDir)) { $ui.OutBox.Text = '(same folder as the input file)' }
    else { $ui.OutBox.Text = $ui.OutDir }
}

function global:Set-InputFile([string]$Path) {
    $ui = $global:UI
    $ui.InputFile = $Path
    $ui.DropText = [System.IO.Path]::GetFileName($Path) + "`n(drop another .docx here to replace it)"
    $ui.Drop.Invalidate()
    $ui.Extract.Enabled = $true
    Set-DocStatus ("Loaded: $Path`nNow check the pattern and output folder, then click Extract.") 'info'
}

function global:Start-Extraction {
    $ui = $global:UI
    $ui.Extract.Enabled = $false
    $ui.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Set-DocStatus 'Extracting...' 'info' -Pump
        $res = Invoke-DocxExtraction -InFile $ui.InputFile -OutDir $ui.OutDir -Pattern $ui.Combo.Text -Raw $ui.Raw.Checked -Whole $ui.Whole.Checked -Case $ui.Case.Checked
        if ($res.Unique -eq 0) {
            Set-DocStatus ("No matches found. A result file with the search settings was still saved:`n" + $res.OutPath) 'warn'
        } else {
            Set-DocStatus ("Done. $($res.Unique) unique matches ($($res.Total) occurrences in total).`nSaved: " + $res.OutPath) 'ok'
        }
        Start-Process explorer.exe -ArgumentList "/select,`"$($res.OutPath)`""
    }
    catch {
        Set-DocStatus ('ERROR: ' + $_.Exception.Message) 'error'
    }
    finally {
        $ui.Form.Cursor = [System.Windows.Forms.Cursors]::Default
        $ui.Extract.Enabled = $true
    }
}

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $global:UI = @{}
    $global:UI.PanelColor = [System.Drawing.Color]::FromArgb(248, 248, 248)
    $global:UI.HoverColor = [System.Drawing.Color]::FromArgb(221, 235, 247)
    $global:UI.InputFile = ''
    $global:UI.OutDir = ''
    $global:UI.DropText = 'Drop a .docx file here'
    $global:UI.HelpSimple = (@(
        'Symbols:   #  one digit      @  one letter      *  letters/digits/_ (1 or more)      \  next character literally',
        '                ?  item before it is optional      {3}  repeat 3 times      {1-3}  repeat 1 to 3 times      ,  several patterns',
        'Examples:  CT###  (CT394)     CT###@?  (CT393 or CT393B)     CT#{1-3}@  (CT7A, CT39B, CT393B)     CT*  (starts with CT)',
        'Spaces are ignored. Matching is done on whole words unless you untick that option.'
    ) -join "`n")
    $global:UI.HelpRaw = (@(
        'Regular expression mode: the text is used as a .NET regular expression, exactly as typed.',
        'Example:  \bCT\d{3}[A-Z]?\b  matches CT394 and CT393B.',
        'Whole word matching is not added automatically in this mode.'
    ) -join "`n")

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Extract SRC_FILE names from a Word document'
    $form.ClientSize = New-Object System.Drawing.Size(600, 640)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true
    $form.TopMost = $true
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Find source file names in application documentation (.docx)'
    $title.Location = New-Object System.Drawing.Point(20, 12)
    $title.Size = New-Object System.Drawing.Size(560, 22)
    $title.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($title)

    $step1 = New-Object System.Windows.Forms.Label
    $step1.Text = '1. Choose the Word document (.docx)'
    $step1.Location = New-Object System.Drawing.Point(20, 44)
    $step1.Size = New-Object System.Drawing.Size(560, 20)
    $form.Controls.Add($step1)

    $browse = New-Object System.Windows.Forms.Button
    $browse.Text = 'Browse...'
    $browse.Location = New-Object System.Drawing.Point(20, 68)
    $browse.Size = New-Object System.Drawing.Size(160, 30)
    $form.Controls.Add($browse)

    $or = New-Object System.Windows.Forms.Label
    $or.Text = 'or drag and drop the file into the box below'
    $or.Location = New-Object System.Drawing.Point(190, 75)
    $or.Size = New-Object System.Drawing.Size(390, 20)
    $form.Controls.Add($or)

    $drop = New-Object System.Windows.Forms.Panel
    $drop.Location = New-Object System.Drawing.Point(20, 104)
    $drop.Size = New-Object System.Drawing.Size(560, 70)
    $drop.BackColor = $global:UI.PanelColor
    $drop.AllowDrop = $true
    $form.Controls.Add($drop)

    $step2 = New-Object System.Windows.Forms.Label
    $step2.Text = '2. Search pattern (type your own or pick an example)'
    $step2.Location = New-Object System.Drawing.Point(20, 186)
    $step2.Size = New-Object System.Drawing.Size(560, 20)
    $form.Controls.Add($step2)

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.DropDownStyle = 'DropDown'
    $combo.Font = New-Object System.Drawing.Font('Consolas', 10)
    $combo.Location = New-Object System.Drawing.Point(20, 210)
    $combo.Size = New-Object System.Drawing.Size(560, 26)
    foreach ($ex in @('CT###', 'CT###@', 'CT###@?', 'CT#{1-3}@', 'CT#{1-3}@?', 'CT*')) { [void]$combo.Items.Add($ex) }
    $form.Controls.Add($combo)

    $help = New-Object System.Windows.Forms.Label
    $help.Location = New-Object System.Drawing.Point(20, 242)
    $help.Size = New-Object System.Drawing.Size(560, 84)
    $help.Font = New-Object System.Drawing.Font('Consolas', 7.5)
    $help.ForeColor = [System.Drawing.Color]::DimGray
    $help.Text = $global:UI.HelpSimple
    $form.Controls.Add($help)

    $cbCase = New-Object System.Windows.Forms.CheckBox
    $cbCase.Text = 'Case sensitive'
    $cbCase.Location = New-Object System.Drawing.Point(20, 334)
    $cbCase.Size = New-Object System.Drawing.Size(170, 24)
    $form.Controls.Add($cbCase)

    $cbWhole = New-Object System.Windows.Forms.CheckBox
    $cbWhole.Text = 'Whole words only'
    $cbWhole.Checked = $true
    $cbWhole.Location = New-Object System.Drawing.Point(200, 334)
    $cbWhole.Size = New-Object System.Drawing.Size(170, 24)
    $form.Controls.Add($cbWhole)

    $cbRaw = New-Object System.Windows.Forms.CheckBox
    $cbRaw.Text = 'Pattern is a regular expression'
    $cbRaw.Location = New-Object System.Drawing.Point(380, 334)
    $cbRaw.Size = New-Object System.Drawing.Size(200, 24)
    $form.Controls.Add($cbRaw)

    $preview = New-Object System.Windows.Forms.Label
    $preview.Location = New-Object System.Drawing.Point(20, 364)
    $preview.Size = New-Object System.Drawing.Size(560, 44)
    $preview.Font = New-Object System.Drawing.Font('Consolas', 8.5)
    $form.Controls.Add($preview)

    $step3 = New-Object System.Windows.Forms.Label
    $step3.Text = '3. Output folder (default: the folder of the .docx file)'
    $step3.Location = New-Object System.Drawing.Point(20, 420)
    $step3.Size = New-Object System.Drawing.Size(560, 20)
    $form.Controls.Add($step3)

    $outBox = New-Object System.Windows.Forms.TextBox
    $outBox.ReadOnly = $true
    $outBox.Location = New-Object System.Drawing.Point(20, 445)
    $outBox.Size = New-Object System.Drawing.Size(340, 24)
    $form.Controls.Add($outBox)

    $outBrowse = New-Object System.Windows.Forms.Button
    $outBrowse.Text = 'Choose...'
    $outBrowse.Location = New-Object System.Drawing.Point(368, 443)
    $outBrowse.Size = New-Object System.Drawing.Size(100, 28)
    $form.Controls.Add($outBrowse)

    $outReset = New-Object System.Windows.Forms.Button
    $outReset.Text = 'Same as input'
    $outReset.Location = New-Object System.Drawing.Point(476, 443)
    $outReset.Size = New-Object System.Drawing.Size(104, 28)
    $form.Controls.Add($outReset)

    $extract = New-Object System.Windows.Forms.Button
    $extract.Text = 'Extract'
    $extract.Enabled = $false
    $extract.Location = New-Object System.Drawing.Point(20, 486)
    $extract.Size = New-Object System.Drawing.Size(560, 38)
    $extract.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($extract)

    $status = New-Object System.Windows.Forms.Label
    $status.Text = 'Step 1: choose a .docx file with Browse or drag and drop.'
    $status.ForeColor = [System.Drawing.Color]::DimGray
    $status.Location = New-Object System.Drawing.Point(20, 534)
    $status.Size = New-Object System.Drawing.Size(560, 90)
    $form.Controls.Add($status)

    $global:UI.Form    = $form
    $global:UI.Browse  = $browse
    $global:UI.Drop    = $drop
    $global:UI.Combo   = $combo
    $global:UI.Case    = $cbCase
    $global:UI.Whole   = $cbWhole
    $global:UI.Raw     = $cbRaw
    $global:UI.Help    = $help
    $global:UI.Preview = $preview
    $global:UI.OutBox  = $outBox
    $global:UI.Extract = $extract
    $global:UI.Status  = $status

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
        $sf.Trimming = [System.Drawing.StringTrimming]::EllipsisCharacter
        $area = New-Object System.Drawing.RectangleF(8, 8, ($s.Width - 16), ($s.Height - 16))
        if ($global:UI.InputFile) { $brush = [System.Drawing.Brushes]::DarkSlateGray } else { $brush = [System.Drawing.Brushes]::DimGray }
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
        $file = $null
        foreach ($i in $items) {
            if ((Test-Path -LiteralPath $i -PathType Leaf) -and ([System.IO.Path]::GetExtension($i) -ieq '.docx')) { $file = $i; break }
        }
        if ($file) {
            Set-InputFile $file
        } else {
            Set-DocStatus 'Please drop a .docx file (Word 2007 or newer). Folders and old .doc files are not supported.' 'error'
        }
    })

    $browse.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = 'Select the Word document'
        $dlg.Filter = 'Word documents (*.docx)|*.docx'
        $dlg.CheckFileExists = $true
        $dlg.Multiselect = $false
        if ($dlg.ShowDialog($global:UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) {
            Set-InputFile $dlg.FileName
        }
        $dlg.Dispose()
    })

    $outBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Select the folder where the result file will be saved'
        $dlg.ShowNewFolderButton = $true
        if ($global:UI.OutDir) { $dlg.SelectedPath = $global:UI.OutDir }
        elseif ($global:UI.InputFile) { $dlg.SelectedPath = (Split-Path -Parent $global:UI.InputFile) }
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

    $combo.Add_TextChanged({ Update-RegexPreview })
    $cbCase.Add_CheckedChanged({ Update-RegexPreview })
    $cbWhole.Add_CheckedChanged({ Update-RegexPreview })
    $cbRaw.Add_CheckedChanged({
        $ui = $global:UI
        $ui.Whole.Enabled = -not $ui.Raw.Checked
        if ($ui.Raw.Checked) { $ui.Help.Text = $ui.HelpRaw } else { $ui.Help.Text = $ui.HelpSimple }
        Update-RegexPreview
    })

    $extract.Add_Click({ Start-Extraction })

    $form.Add_Shown({
        param($s, $e)
        $s.Activate()
    })

    $combo.Text = 'CT###'
    Update-OutDirBox
    Update-RegexPreview

    [void]$form.ShowDialog()
    exit 0
}
catch {
    Write-Host ''
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
