Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$root = Split-Path -Parent $PSScriptRoot
$script = Join-Path $PSScriptRoot "MR-Outlook.ps1"
$inputFile = Join-Path $root "Input\MR-Outlook\input.txt"
$inputRoot = Join-Path $root "Input"
$replyFolderFile = Join-Path $root "Config\reply_folder.txt"
$vendorDisplayToName = @{}

$bg = [Drawing.ColorTranslator]::FromHtml("#11181c")
$panel = [Drawing.ColorTranslator]::FromHtml("#172126")
$text = [Drawing.ColorTranslator]::FromHtml("#f4f7f8")
$muted = [Drawing.ColorTranslator]::FromHtml("#b8c7ce")
$accent = [Drawing.ColorTranslator]::FromHtml("#f2b632")
$teal = [Drawing.ColorTranslator]::FromHtml("#247b83")
$green = [Drawing.ColorTranslator]::FromHtml("#45a36b")
$red = [Drawing.ColorTranslator]::FromHtml("#a54850")

function quote($value) {
    '"' + ([string]$value -replace '"', '\"') + '"'
}

function psquote($value) {
    "'" + ([string]$value -replace "'", "''") + "'"
}

function label($textValue, $x, $y, $w, $h, $size, $color, $bold=$false) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $textValue
    $l.Location = New-Object Drawing.Point($x, $y)
    $l.Size = New-Object Drawing.Size($w, $h)
    $style = if ($bold) { [Drawing.FontStyle]::Bold } else { [Drawing.FontStyle]::Regular }
    $l.Font = New-Object Drawing.Font("Segoe UI", $size, $style)
    $l.ForeColor = $color
    $l.BackColor = [Drawing.Color]::Transparent
    $l
}

function button($textValue, $x, $y, $w, $h, $color) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $textValue
    $b.Location = New-Object Drawing.Point($x, $y)
    $b.Size = New-Object Drawing.Size($w, $h)
    $b.FlatStyle = "Flat"
    $b.FlatAppearance.BorderSize = 0
    $b.BackColor = $color
    $b.ForeColor = $text
    $b.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
    $b
}

function latest-input-folder {
    $latest = Get-ChildItem -LiteralPath $inputRoot -Directory | Where-Object { $_.Name -match '^\d{2}\.\d{2}\.\d{4}$' } | ForEach-Object { $d=[datetime]::MinValue; if([datetime]::TryParseExact($_.Name,'dd.MM.yyyy',$null,[Globalization.DateTimeStyles]::None,[ref]$d)){[pscustomobject]@{Folder=$_;Date=$d}} } | Sort-Object Date -Descending | Select-Object -First 1
    if($latest){$latest=$latest.Folder}
    $latest
}

function master-source {
    $latest = latest-input-folder
    if ($latest) { Join-Path $latest.FullName "MR_Master_Input.txt" } else { $inputFile }
}

function get-supplier-stats {
    $source = master-source
    if (!(Test-Path -LiteralPath $source)) { return @() }
    $delim = if ((Get-Content -LiteralPath $source -TotalCount 1) -like "*`t*") { "`t" } else { "," }
    $rows = @(Import-Csv -LiteralPath $source -Delimiter $delim)
    if (!$rows) { return @() }
    $vendorCol = @("Vendor Name", "Vendor name", "VendorName") | Where-Object { $rows[0].PSObject.Properties[$_] } | Select-Object -First 1
    if (!$vendorCol) { return @() }
    @($rows | ForEach-Object { ([string]$_.$vendorCol).Trim() } | Where-Object { $_ } | Group-Object | Sort-Object @{Expression="Count";Descending=$true}, @{Expression="Name";Ascending=$true} | ForEach-Object {
        [pscustomobject]@{ Vendor=$_.Name; Count=$_.Count }
    })
}

function refresh-suppliers {
    $supplierList.Items.Clear()
    $vendorDisplayToName.Clear()
    $stats = @(get-supplier-stats)
    $totalItems = 0
    foreach ($s in $stats) {
        $totalItems += [int]$s.Count
        $display = "{0} ({1} items)" -f $s.Vendor, $s.Count
        $vendorDisplayToName[$display] = $s.Vendor
        [void]$supplierList.Items.Add($display, $true)
    }
    $vendorCountLabel.Text = "($($stats.Count) suppliers / $totalItems items)"
    $source = master-source
    $masterPathLabel.Text = if (Test-Path -LiteralPath $source) { "Master: $source" } else { "Master: not prepared yet" }
}

function selected-vendors {
    @($supplierList.CheckedItems | ForEach-Object {
        $display = [string]$_
        if ($vendorDisplayToName.ContainsKey($display)) { $vendorDisplayToName[$display] } else { $display }
    })
}

function run-prepare {
    $prepareButton.Enabled = $false
    try {
        $cmd = "& $(psquote $script) -Mode prepare; Write-Host ''; Read-Host 'Press Enter to return to Launcher'"
        Start-Process -FilePath "powershell.exe" -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -Command " + (quote $cmd)) -WorkingDirectory $root -Wait | Out-Null
        refresh-suppliers
        [System.Windows.Forms.MessageBox]::Show("Prepare complete. Supplier list has been refreshed.", "MR Outlook") | Out-Null
    } finally {
        $prepareButton.Enabled = $true
    }
}

function run-one($mode, $vendor) {
    $args = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-NoExit",
        "-File", (quote $script),
        "-Mode", $mode
    )
    if ($vendor) { $args += @("-Vendor", (quote $vendor)) }
    if ($mode -eq "scan") { $args += @("-ReplyFolder", (quote $replyFolderBox.Text.Trim())) }
    if ($displayBox.Checked) { $args += "-Display" }

    Start-Process -FilePath "powershell.exe" -ArgumentList ($args -join " ") -WorkingDirectory $root
}

function run-mr($mode) {
    $selected = selected-vendors
    if ($mode -eq "send" -and $supplierList.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please run Prepare Input first to create the master data.", "MR Outlook") | Out-Null
        return
    }
    if (!$selected -and $supplierList.Items.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one supplier.", "MR Outlook") | Out-Null
        return
    }
    if ($mode -eq "scan") {
        if (!$replyFolderBox.Text.Trim()) {
            [System.Windows.Forms.MessageBox]::Show("Please enter the Inbox folder name for Scan Replies.", "MR Outlook") | Out-Null
            return
        }
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $replyFolderFile) | Out-Null
        Set-Content -LiteralPath $replyFolderFile -Value $replyFolderBox.Text.Trim() -Encoding UTF8
    }

    if ($supplierList.Items.Count -eq 0 -or $selected.Count -eq $supplierList.Items.Count) {
        run-one $mode $null
    } else {
        $selected | ForEach-Object { run-one $mode $_ }
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "MR Outlook Auto Tools"
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.ClientSize = New-Object Drawing.Size(760, 620)
$form.BackColor = $bg

$topBar = New-Object System.Windows.Forms.Panel
$topBar.Location = New-Object Drawing.Point(0, 0)
$topBar.Size = New-Object Drawing.Size(760, 7)
$topBar.BackColor = $accent
$form.Controls.Add($topBar)

$form.Controls.Add((label "MR OUTLOOK" 30 28 200 22 9 $accent $true))
$form.Controls.Add((label "Material Request Auto Tools" 30 52 430 38 21 $text $true))
$form.Controls.Add((label "Send, scan supplier replies, and remind pending MR requests." 32 92 500 24 9 $muted))
$pathLabel = label $root 32 113 690 20 8 $muted
$pathLabel.AutoEllipsis = $true
$form.Controls.Add($pathLabel)

$clockLabel = label "" 460 52 270 26 10 $text $true
$clockLabel.TextAlign = "MiddleRight"
$form.Controls.Add($clockLabel)
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$timer.Add_Tick({ $clockLabel.Text = (Get-Date -Format "dddd, MMM dd yyyy  HH:mm:ss") })
$timer.Start()
$clockLabel.Text = (Get-Date -Format "dddd, MMM dd yyyy  HH:mm:ss")

$configPanel = New-Object System.Windows.Forms.Panel
$configPanel.Location = New-Object Drawing.Point(30, 150)
$configPanel.Size = New-Object Drawing.Size(700, 74)
$configPanel.BackColor = $panel
$form.Controls.Add($configPanel)

$configPanel.Controls.Add((label "Inbox scan folder" 18 14 130 22 9 $muted))
$replyFolderBox = New-Object System.Windows.Forms.TextBox
$replyFolderBox.Text = if (Test-Path -LiteralPath $replyFolderFile) { (Get-Content -LiteralPath $replyFolderFile -TotalCount 1).Trim() } else { "MR_REQUEST" }
$replyFolderBox.Location = New-Object Drawing.Point(150, 12)
$replyFolderBox.Size = New-Object Drawing.Size(260, 24)
$replyFolderBox.Font = New-Object Drawing.Font("Segoe UI", 9)
$configPanel.Controls.Add($replyFolderBox)

$displayBox = New-Object System.Windows.Forms.CheckBox
$displayBox.Text = "Display only, do not send"
$displayBox.Checked = $true
$displayBox.Location = New-Object Drawing.Point(445, 13)
$displayBox.Size = New-Object Drawing.Size(210, 24)
$displayBox.ForeColor = $text
$displayBox.BackColor = $panel
$displayBox.Font = New-Object Drawing.Font("Segoe UI", 9)
$configPanel.Controls.Add($displayBox)

$configPanel.Controls.Add((label "Selected suppliers open in separate PowerShell output windows." 18 46 520 20 8 $muted))

$supplierPanel = New-Object System.Windows.Forms.Panel
$supplierPanel.Location = New-Object Drawing.Point(30, 245)
$supplierPanel.Size = New-Object Drawing.Size(700, 260)
$supplierPanel.BackColor = $panel
$form.Controls.Add($supplierPanel)

$supplierPanel.Controls.Add((label "Suppliers in input" 18 12 180 24 11 $text $true))
$vendorCountLabel = label "" 200 14 130 22 9 $muted
$supplierPanel.Controls.Add($vendorCountLabel)

$refreshButton = button "Refresh" 488 10 70 28 $accent
$refreshButton.ForeColor = $bg
$refreshButton.Add_Click({ refresh-suppliers })
$supplierPanel.Controls.Add($refreshButton)

$allButton = button "All" 565 10 55 28 $teal
$allButton.Add_Click({ for ($i=0; $i -lt $supplierList.Items.Count; $i++) { $supplierList.SetItemChecked($i, $true) } })
$supplierPanel.Controls.Add($allButton)

$noneButton = button "None" 628 10 55 28 $red
$noneButton.Add_Click({ for ($i=0; $i -lt $supplierList.Items.Count; $i++) { $supplierList.SetItemChecked($i, $false) } })
$supplierPanel.Controls.Add($noneButton)

$supplierList = New-Object System.Windows.Forms.CheckedListBox
$supplierList.CheckOnClick = $true
$supplierList.Location = New-Object Drawing.Point(18, 48)
$supplierList.Size = New-Object Drawing.Size(665, 178)
$supplierList.BackColor = [Drawing.ColorTranslator]::FromHtml("#0f171b")
$supplierList.ForeColor = $text
$supplierList.BorderStyle = "FixedSingle"
$supplierList.Font = New-Object Drawing.Font("Segoe UI", 9)
$supplierPanel.Controls.Add($supplierList)

$masterPathLabel = label "" 18 230 665 20 8 $muted
$masterPathLabel.AutoEllipsis = $true
$supplierPanel.Controls.Add($masterPathLabel)

$prepareButton = button "Prepare Input" 30 535 160 58 $accent
$prepareButton.ForeColor = $bg
$prepareButton.Add_Click({ run-prepare })
$form.Controls.Add($prepareButton)

$sendButton = button "Send MR" 210 535 160 58 $teal
$sendButton.Add_Click({ run-mr "send" })
$form.Controls.Add($sendButton)

$scanButton = button "Scan Replies" 390 535 160 58 $green
$scanButton.Add_Click({ run-mr "scan" })
$form.Controls.Add($scanButton)

$remindButton = button "Remind" 570 535 160 58 $red
$remindButton.Add_Click({ run-mr "remind" })
$form.Controls.Add($remindButton)

refresh-suppliers

[void]$form.ShowDialog()
