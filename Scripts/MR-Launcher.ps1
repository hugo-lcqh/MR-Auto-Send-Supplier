Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$root = Split-Path -Parent $PSScriptRoot
$script = Join-Path $PSScriptRoot "MR-Outlook.ps1"
$inputFile = Join-Path $root "Input\MR-Outlook\input.txt"
$inputRoot = Join-Path $root "Input"
$replyFolderFile = Join-Path $root "Config\reply_folder.txt"
$vendorDisplayToName = @{}
$appVersion = "1.3.0"
$uninstallKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{8D183A16-6C3F-4EF1-B296-502FB04CF3C6}_is1"
try {
    $installedApp = Get-ItemProperty -LiteralPath $uninstallKey -ErrorAction Stop
    $installedRoot = [IO.Path]::GetFullPath([string]$installedApp.InstallLocation).TrimEnd('\')
    if ($installedRoot -eq [IO.Path]::GetFullPath($root).TrimEnd('\') -and $installedApp.DisplayVersion) {
        $appVersion = [string]$installedApp.DisplayVersion
    }
} catch {}

$bg = [Drawing.ColorTranslator]::FromHtml("#EEF2F5")
$panel = [Drawing.ColorTranslator]::FromHtml("#FFFFFF")
$text = [Drawing.ColorTranslator]::FromHtml("#14232E")
$muted = [Drawing.ColorTranslator]::FromHtml("#5B6B76")
$accent = [Drawing.ColorTranslator]::FromHtml("#0E6F9F")
$teal = [Drawing.ColorTranslator]::FromHtml("#0A6F75")
$green = [Drawing.ColorTranslator]::FromHtml("#167A55")
$red = [Drawing.ColorTranslator]::FromHtml("#995600")
$navy = [Drawing.ColorTranslator]::FromHtml("#0B2239")
$line = [Drawing.ColorTranslator]::FromHtml("#D6DEE3")
$subtle = [Drawing.ColorTranslator]::FromHtml("#F7F9FA")
$white = [Drawing.Color]::White
$focus = [Drawing.ColorTranslator]::FromHtml("#F4B41A")

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
    $l.Margin = New-Object System.Windows.Forms.Padding(0)
    $l
}

function button($textValue, $x, $y, $w, $h, $color) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $textValue
    $b.Location = New-Object Drawing.Point($x, $y)
    $b.Size = New-Object Drawing.Size($w, $h)
    $b.FlatStyle = "Flat"
    $b.FlatAppearance.BorderSize = 1
    $b.FlatAppearance.BorderColor = $color
    $b.FlatAppearance.MouseOverBackColor = [System.Windows.Forms.ControlPaint]::Light($color, 0.12)
    $b.FlatAppearance.MouseDownBackColor = [System.Windows.Forms.ControlPaint]::Dark($color, 0.08)
    $b.BackColor = $color
    $b.ForeColor = $white
    $b.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    $b.TextAlign = "MiddleLeft"
    $b.Padding = New-Object System.Windows.Forms.Padding(14, 0, 10, 0)
    $b.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
    $b.UseVisualStyleBackColor = $false
    $b.AccessibleName = $textValue
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

function set-status($message, $color = $green) {
    if ($statusLabel) {
        $statusLabel.Text = $message
        $statusLabel.ForeColor = $color
    }
    if ($readyBadge) {
        $readyBadge.Text = $message
        $readyBadge.BackColor = $color
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function run-prepare {
    $prepareComplete = $false
    $prepareButton.Enabled = $false
    $form.UseWaitCursor = $true
    set-status "PREPARE RUNNING" $accent
    try {
        $cmd = "& $(psquote $script) -Mode prepare; Write-Host ''; Read-Host 'Press Enter to return to Launcher'"
        Start-Process -FilePath "powershell.exe" -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -Command " + (quote $cmd)) -WorkingDirectory $root -Wait | Out-Null
        refresh-suppliers
        $prepareComplete = $true
        set-status "PREPARE COMPLETE" $green
        [System.Windows.Forms.MessageBox]::Show("Prepare complete. Supplier list has been refreshed.", "MR Outlook") | Out-Null
    } finally {
        if (!$prepareComplete) { set-status "PREPARE FAILED" $red }
        $form.UseWaitCursor = $false
        $prepareButton.Enabled = $true
    }
}

function encode-vendor-selection($vendors) {
    $json = ConvertTo-Json -InputObject @($vendors) -Compress
    [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
}

function run-one($mode, $vendor, $vendorSelection) {
    $args = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-NoExit",
        "-File", (quote $script),
        "-Mode", $mode
    )
    if ($vendor) { $args += @("-Vendor", (quote $vendor)) }
    if ($vendorSelection) { $args += @("-VendorSelection", $vendorSelection) }
    if ($mode -eq "scan") { $args += @("-ReplyFolder", (quote $replyFolderBox.Text.Trim())) }
    if ($displayBox.Checked) { $args += "-Display" }
    if ($mode -eq "send" -and !$emailTableBox.Checked) { $args += "-HideEmailTable" }

    Start-Process -FilePath "powershell.exe" -ArgumentList ($args -join " ") -WorkingDirectory $root
    set-status ("{0} STARTED" -f $mode.ToUpperInvariant()) $green
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
        run-one $mode $null $null
    } else {
        run-one $mode $null (encode-vendor-selection $selected)
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "MR Outlook | Material Request Control Center"
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true
$form.MinimumSize = New-Object Drawing.Size(960, 740)
$form.ClientSize = New-Object Drawing.Size(1080, 740)
$form.AutoScaleMode = "Dpi"
$form.Font = New-Object Drawing.Font("Segoe UI", 9)
$form.BackColor = $bg
$form.KeyPreview = $true

$rootLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rootLayout.Dock = "Fill"
$rootLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$rootLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$rootLayout.ColumnCount = 1
$rootLayout.RowCount = 3
[void]$rootLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
[void]$rootLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 118))
[void]$rootLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
[void]$rootLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 42))
$form.Controls.Add($rootLayout)

$header = New-Object System.Windows.Forms.Panel
$header.Dock = "Fill"
$header.Margin = New-Object System.Windows.Forms.Padding(0)
$header.BackColor = $navy
$rootLayout.Controls.Add($header, 0, 0)

$headerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$headerLayout.Dock = "Fill"
$headerLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$headerLayout.Padding = New-Object System.Windows.Forms.Padding(32, 12, 28, 10)
$headerLayout.ColumnCount = 2
$headerLayout.RowCount = 3
[void]$headerLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 66))
[void]$headerLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 34))
[void]$headerLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 24))
[void]$headerLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 40))
[void]$headerLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$header.Controls.Add($headerLayout)

$headerAccent = New-Object System.Windows.Forms.Panel
$headerAccent.Dock = "Left"
$headerAccent.Width = 6
$headerAccent.BackColor = $focus
$header.Controls.Add($headerAccent)
$headerAccent.BringToFront()

$brandLabel = label "MR OUTLOOK  /  SUPPLIER OPERATIONS" 0 0 0 0 9 $focus $true
$brandLabel.Dock = "Fill"
$headerLayout.Controls.Add($brandLabel, 0, 0)

$titleLabel = label "MATERIAL REQUEST CONTROL CENTER" 0 0 0 0 21 $white $true
$titleLabel.Dock = "Fill"
$headerLayout.Controls.Add($titleLabel, 0, 1)

$subtitleLabel = label "Prepare input, send requests, scan replies, and follow up from one controlled workspace." 0 0 0 0 9 ([Drawing.ColorTranslator]::FromHtml("#C9D6E0"))
$subtitleLabel.Dock = "Fill"
$headerLayout.Controls.Add($subtitleLabel, 0, 2)

$headerMetaLayout = New-Object System.Windows.Forms.TableLayoutPanel
$headerMetaLayout.Dock = "Fill"
$headerMetaLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$headerMetaLayout.ColumnCount = 1
$headerMetaLayout.RowCount = 3
[void]$headerMetaLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
[void]$headerMetaLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 30))
[void]$headerMetaLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 30))
[void]$headerMetaLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$headerLayout.Controls.Add($headerMetaLayout, 1, 0)
$headerLayout.SetRowSpan($headerMetaLayout, 3)

$readyBadge = label "SYSTEM READY" 0 0 140 24 8 $white $true
$readyBadge.TextAlign = "MiddleCenter"
$readyBadge.BackColor = $green
$readyBadge.Anchor = "Top,Right"
$headerMetaLayout.Controls.Add($readyBadge, 0, 0)

$clockLabel = label "" 0 0 0 0 10 $white $true
$clockLabel.Dock = "Fill"
$clockLabel.TextAlign = "MiddleRight"
$headerMetaLayout.Controls.Add($clockLabel, 0, 1)

$pathLabel = label $root 0 0 0 0 8 ([Drawing.ColorTranslator]::FromHtml("#AFC0CC"))
$pathLabel.Dock = "Fill"
$pathLabel.AutoEllipsis = $true
$pathLabel.TextAlign = "MiddleRight"
$headerMetaLayout.Controls.Add($pathLabel, 0, 2)

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$timer.Add_Tick({ $clockLabel.Text = (Get-Date -Format "ddd, dd MMM yyyy  |  HH:mm:ss") })
$timer.Start()
$clockLabel.Text = (Get-Date -Format "ddd, dd MMM yyyy  |  HH:mm:ss")

$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = "Fill"
$mainLayout.BackColor = $bg
$mainLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$mainLayout.Padding = New-Object System.Windows.Forms.Padding(24, 22, 24, 22)
$mainLayout.ColumnCount = 2
$mainLayout.RowCount = 1
[void]$mainLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 64))
[void]$mainLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 36))
[void]$mainLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$rootLayout.Controls.Add($mainLayout, 0, 1)

$supplierPanel = New-Object System.Windows.Forms.Panel
$supplierPanel.Dock = "Fill"
$supplierPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 9, 0)
$supplierPanel.BackColor = $panel
$supplierPanel.BorderStyle = "FixedSingle"
$mainLayout.Controls.Add($supplierPanel, 0, 0)

$supplierLayout = New-Object System.Windows.Forms.TableLayoutPanel
$supplierLayout.Dock = "Fill"
$supplierLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$supplierLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$supplierLayout.ColumnCount = 1
$supplierLayout.RowCount = 3
[void]$supplierLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
[void]$supplierLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 70))
[void]$supplierLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
[void]$supplierLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 42))
$supplierPanel.Controls.Add($supplierLayout)

$queueHeader = New-Object System.Windows.Forms.Panel
$queueHeader.Dock = "Fill"
$queueHeader.BackColor = $panel
$supplierLayout.Controls.Add($queueHeader, 0, 0)
$queueHeader.Controls.Add((label "SUPPLIER QUEUE" 16 14 220 24 12 $text $true))
$vendorCountLabel = label "" 17 39 260 20 9 $muted
$queueHeader.Controls.Add($vendorCountLabel)

$queueTools = New-Object System.Windows.Forms.FlowLayoutPanel
$queueTools.Dock = "Right"
$queueTools.Width = 296
$queueTools.Padding = New-Object System.Windows.Forms.Padding(0, 13, 12, 0)
$queueTools.FlowDirection = "RightToLeft"
$queueTools.WrapContents = $false
$queueTools.BackColor = $panel
$queueHeader.Controls.Add($queueTools)

$noneButton = button "Clear" 0 0 72 40 $red
$noneButton.TextAlign = "MiddleCenter"
$noneButton.Padding = New-Object System.Windows.Forms.Padding(0)
$noneButton.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
$noneButton.TabIndex = 2
$noneButton.AccessibleDescription = "Clear every supplier selection"
$noneButton.Add_Click({ for ($i=0; $i -lt $supplierList.Items.Count; $i++) { $supplierList.SetItemChecked($i, $false) } })
$queueTools.Controls.Add($noneButton)

$allButton = button "Select all" 0 0 88 40 $teal
$allButton.TextAlign = "MiddleCenter"
$allButton.Padding = New-Object System.Windows.Forms.Padding(0)
$allButton.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
$allButton.TabIndex = 1
$allButton.AccessibleDescription = "Select every supplier in the queue"
$allButton.Add_Click({ for ($i=0; $i -lt $supplierList.Items.Count; $i++) { $supplierList.SetItemChecked($i, $true) } })
$queueTools.Controls.Add($allButton)

$refreshButton = button "Refresh" 0 0 82 40 $accent
$refreshButton.TextAlign = "MiddleCenter"
$refreshButton.Padding = New-Object System.Windows.Forms.Padding(0)
$refreshButton.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
$refreshButton.TabIndex = 0
$refreshButton.AccessibleDescription = "Reload suppliers from the latest master input"
$refreshButton.Add_Click({ refresh-suppliers; set-status "QUEUE REFRESHED" $green })
$queueTools.Controls.Add($refreshButton)

$supplierList = New-Object System.Windows.Forms.CheckedListBox
$supplierList.CheckOnClick = $true
$supplierList.Dock = "Fill"
$supplierList.Margin = New-Object System.Windows.Forms.Padding(16, 0, 16, 0)
$supplierList.BackColor = $subtle
$supplierList.ForeColor = $text
$supplierList.BorderStyle = "FixedSingle"
$supplierList.Font = New-Object Drawing.Font("Segoe UI", 10)
$supplierList.IntegralHeight = $false
$supplierList.TabIndex = 3
$supplierList.AccessibleName = "Suppliers selected for this operation"
$supplierList.AccessibleDescription = "Checked suppliers run together in one PowerShell window"
$supplierLayout.Controls.Add($supplierList, 0, 1)

$masterPathLabel = label "" 0 0 0 0 8 $muted
$masterPathLabel.Dock = "Fill"
$masterPathLabel.Padding = New-Object System.Windows.Forms.Padding(15, 3, 15, 0)
$masterPathLabel.TextAlign = "MiddleLeft"
$masterPathLabel.AutoEllipsis = $true
$supplierLayout.Controls.Add($masterPathLabel, 0, 2)

$rightLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rightLayout.Dock = "Fill"
$rightLayout.Margin = New-Object System.Windows.Forms.Padding(9, 0, 0, 0)
$rightLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$rightLayout.ColumnCount = 1
$rightLayout.RowCount = 2
[void]$rightLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
[void]$rightLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 204))
[void]$rightLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$mainLayout.Controls.Add($rightLayout, 1, 0)

$configPanel = New-Object System.Windows.Forms.Panel
$configPanel.Dock = "Fill"
$configPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 16)
$configPanel.BackColor = $panel
$configPanel.BorderStyle = "FixedSingle"
$rightLayout.Controls.Add($configPanel, 0, 0)

$configLayout = New-Object System.Windows.Forms.TableLayoutPanel
$configLayout.Dock = "Fill"
$configLayout.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 10)
$configLayout.ColumnCount = 1
$configLayout.RowCount = 4
[void]$configLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
[void]$configLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 30))
[void]$configLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 62))
[void]$configLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 36))
[void]$configLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 36))
$configPanel.Controls.Add($configLayout)

$settingsTitle = label "RUN SETTINGS" 0 0 0 0 11 $text $true
$settingsTitle.Dock = "Fill"
$configLayout.Controls.Add($settingsTitle, 0, 0)

$folderPanel = New-Object System.Windows.Forms.Panel
$folderPanel.Dock = "Fill"
$folderPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$folderPanel.BackColor = $panel
$configLayout.Controls.Add($folderPanel, 0, 1)
$folderLabel = label "Inbox folder used by Scan Replies" 0 0 300 20 8 $muted
$folderLabel.Dock = "Top"
$folderPanel.Controls.Add($folderLabel)
$replyFolderBox = New-Object System.Windows.Forms.TextBox
$replyFolderBox.Text = if (Test-Path -LiteralPath $replyFolderFile) { (Get-Content -LiteralPath $replyFolderFile -TotalCount 1).Trim() } else { "MR_REQUEST" }
$replyFolderBox.Dock = "Bottom"
$replyFolderBox.Height = 30
$replyFolderBox.Font = New-Object Drawing.Font("Segoe UI", 10)
$replyFolderBox.BackColor = $subtle
$replyFolderBox.ForeColor = $text
$replyFolderBox.BorderStyle = "FixedSingle"
$replyFolderBox.TabIndex = 4
$replyFolderBox.AccessibleName = "Inbox folder for scanning replies"
$folderPanel.Controls.Add($replyFolderBox)

$displayBox = New-Object System.Windows.Forms.CheckBox
$displayBox.Text = "Display only, do not send"
$displayBox.Checked = $true
$displayBox.Dock = "Fill"
$displayBox.ForeColor = $text
$displayBox.BackColor = $panel
$displayBox.Font = New-Object Drawing.Font("Segoe UI", 9)
$displayBox.FlatStyle = "Flat"
$displayBox.TabIndex = 5
$displayBox.AccessibleDescription = "Open draft emails for review instead of sending them"
$configLayout.Controls.Add($displayBox, 0, 2)

$emailTableBox = New-Object System.Windows.Forms.CheckBox
$emailTableBox.Text = "Show MR table in email"
$emailTableBox.Checked = $true
$emailTableBox.Dock = "Fill"
$emailTableBox.ForeColor = $text
$emailTableBox.BackColor = $panel
$emailTableBox.Font = New-Object Drawing.Font("Segoe UI", 9)
$emailTableBox.FlatStyle = "Flat"
$emailTableBox.TabIndex = 6
$emailTableBox.AccessibleDescription = "Include the material request summary table in supplier email bodies"
$configLayout.Controls.Add($emailTableBox, 0, 3)

$workflowPanel = New-Object System.Windows.Forms.Panel
$workflowPanel.Dock = "Fill"
$workflowPanel.Margin = New-Object System.Windows.Forms.Padding(0)
$workflowPanel.BackColor = $panel
$workflowPanel.BorderStyle = "FixedSingle"
$rightLayout.Controls.Add($workflowPanel, 0, 1)

$workflowLayout = New-Object System.Windows.Forms.TableLayoutPanel
$workflowLayout.Dock = "Fill"
$workflowLayout.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 6)
$workflowLayout.ColumnCount = 1
$workflowLayout.RowCount = 5
[void]$workflowLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
[void]$workflowLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 52))
for ($rowIndex = 1; $rowIndex -lt 5; $rowIndex++) {
    [void]$workflowLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 25))
}
$workflowPanel.Controls.Add($workflowLayout)

$workflowHeader = New-Object System.Windows.Forms.Panel
$workflowHeader.Dock = "Fill"
$workflowHeader.BackColor = $panel
$workflowLayout.Controls.Add($workflowHeader, 0, 0)
$workflowHeader.Controls.Add((label "OPERATION WORKFLOW" 0 0 310 23 11 $text $true))
$workflowHeader.Controls.Add((label "Select suppliers, then run one controlled task." 1 24 320 20 8 $muted))

$prepareButton = button "01   PREPARE INPUT" 0 0 0 0 $navy
$prepareButton.Dock = "Fill"
$prepareButton.TabIndex = 7
$prepareButton.AccessibleDescription = "Build the latest master input and refresh the supplier queue"
$prepareButton.Add_Click({ run-prepare })
$workflowLayout.Controls.Add($prepareButton, 0, 1)

$sendButton = button "02   SEND MR" 0 0 0 0 $accent
$sendButton.Dock = "Fill"
$sendButton.TabIndex = 8
$sendButton.AccessibleDescription = "Create or send material request emails for selected suppliers"
$sendButton.Add_Click({ run-mr "send" })
$workflowLayout.Controls.Add($sendButton, 0, 2)

$scanButton = button "03   SCAN REPLIES" 0 0 0 0 $green
$scanButton.Dock = "Fill"
$scanButton.TabIndex = 9
$scanButton.AccessibleDescription = "Scan the configured Inbox folder for supplier reply attachments"
$scanButton.Add_Click({ run-mr "scan" })
$workflowLayout.Controls.Add($scanButton, 0, 3)

$remindButton = button "04   REMIND PENDING" 0 0 0 0 $red
$remindButton.Dock = "Fill"
$remindButton.TabIndex = 10
$remindButton.AccessibleDescription = "Create or send reminders for pending material requests"
$remindButton.Add_Click({ run-mr "remind" })
$workflowLayout.Controls.Add($remindButton, 0, 4)

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 8000
$toolTip.InitialDelay = 450
$toolTip.ReshowDelay = 100
$toolTip.SetToolTip($prepareButton, $prepareButton.AccessibleDescription)
$toolTip.SetToolTip($sendButton, $sendButton.AccessibleDescription)
$toolTip.SetToolTip($scanButton, $scanButton.AccessibleDescription)
$toolTip.SetToolTip($remindButton, $remindButton.AccessibleDescription)

$footer = New-Object System.Windows.Forms.TableLayoutPanel
$footer.Dock = "Fill"
$footer.Margin = New-Object System.Windows.Forms.Padding(0)
$footer.Padding = New-Object System.Windows.Forms.Padding(24, 4, 16, 4)
$footer.BackColor = $panel
$footer.ColumnCount = 3
$footer.RowCount = 1
[void]$footer.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Absolute, 130))
[void]$footer.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
[void]$footer.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Absolute, 84))
[void]$footer.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
$rootLayout.Controls.Add($footer, 0, 2)

$statusLabel = label "SYSTEM READY" 0 0 0 0 8 $green $true
$statusLabel.Dock = "Fill"
$statusLabel.TextAlign = "MiddleLeft"
$footer.Controls.Add($statusLabel, 0, 0)

$footerLabel = label "Version $appVersion | Developed by Hugo Le Chi Quoc Hung | Phone: +84 39 5656 909" 0 0 0 0 8 $muted
$footerLabel.Dock = "Fill"
$footerLabel.TextAlign = "MiddleRight"
$footer.Controls.Add($footerLabel, 1, 0)

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Close"
$closeButton.Dock = "Fill"
$closeButton.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$closeButton.FlatStyle = "Flat"
$closeButton.FlatAppearance.BorderColor = $line
$closeButton.BackColor = $subtle
$closeButton.ForeColor = $text
$closeButton.Font = New-Object Drawing.Font("Segoe UI", 8, [Drawing.FontStyle]::Bold)
$closeButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$closeButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$closeButton.TabIndex = 11
$closeButton.AccessibleDescription = "Close the MR Outlook launcher"
$footer.Controls.Add($closeButton, 2, 0)
$form.CancelButton = $closeButton

$form.Add_Shown({ $supplierList.Focus() })

refresh-suppliers

[void]$form.ShowDialog()
