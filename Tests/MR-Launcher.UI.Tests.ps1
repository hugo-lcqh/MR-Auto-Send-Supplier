$root = Split-Path -Parent $PSScriptRoot
$launcherPath = Join-Path $root "Scripts\MR-Launcher.ps1"

$tokens = $null
$parserErrors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($launcherPath, [ref]$tokens, [ref]$parserErrors)
$source = Get-Content -LiteralPath $launcherPath -Raw
$failures = [System.Collections.Generic.List[string]]::new()

function Assert-LauncherContains([string]$text, [string]$message) {
    if (!$source.Contains($text)) { $failures.Add($message) }
}

if ($parserErrors.Count) {
    $failures.Add("Launcher must parse without errors")
}

Assert-LauncherContains '$form.AutoScaleMode = "Dpi"' "Launcher must scale cleanly on factory displays with non-default DPI"
Assert-LauncherContains '$form.FormBorderStyle = "Sizable"' "Launcher must support different desktop resolutions"
Assert-LauncherContains '$form.MinimumSize' "Resizable launcher must define a usable minimum size"
Assert-LauncherContains 'System.Windows.Forms.TableLayoutPanel' "Launcher must use layout containers instead of a fixed coordinate-only canvas"
Assert-LauncherContains '$headerLayout = New-Object System.Windows.Forms.TableLayoutPanel' "Header content must adapt without overlapping at minimum window width"
Assert-LauncherContains 'MATERIAL REQUEST CONTROL CENTER' "Header must identify the manufacturing operations workspace"
Assert-LauncherContains 'SUPPLIER QUEUE' "Supplier selection must have a clear operations label"
Assert-LauncherContains 'RUN SETTINGS' "Runtime options must be grouped separately from actions"
Assert-LauncherContains 'OPERATION WORKFLOW' "Launcher actions must be presented as an ordered workflow"
Assert-LauncherContains 'SYSTEM READY' "Launcher must expose a readable system status"
Assert-LauncherContains '$supplierList.AccessibleName' "Supplier list must expose an accessibility name"
Assert-LauncherContains '$replyFolderBox.AccessibleName' "Inbox folder input must expose an accessibility name"
Assert-LauncherContains '$prepareButton.AccessibleDescription' "Workflow actions must explain their purpose to assistive technology"
Assert-LauncherContains '$form.CancelButton = $closeButton' "Launcher must provide a predictable Escape-key exit"
Assert-LauncherContains '[System.Windows.Forms.ControlPaint]' "Button interaction colors must use the WinForms ControlPaint runtime type"
Assert-LauncherContains 'function set-status' "Launcher must provide operational feedback while tasks start and complete"
Assert-LauncherContains 'PREPARE RUNNING' "Prepare must expose an in-progress status"
if ($source.Contains('[Drawing.ControlPaint]')) {
    $failures.Add("Launcher must not reference the non-existent Drawing.ControlPaint type")
}

if ($failures.Count -eq 0) {
    $scriptDirectory = Split-Path -Parent $launcherPath
    $quotedScriptDirectory = "'" + ($scriptDirectory -replace "'", "''") + "'"
    $runtimeSource = $source.Replace('$PSScriptRoot', $quotedScriptDirectory).Replace('[void]$form.ShowDialog()', '')
    try {
        & {
            param($runtimeSource)
            Invoke-Expression $runtimeSource
            try {
                $form.Size = $form.MinimumSize
                $form.PerformLayout()
                $rootLayout.PerformLayout()
                $headerLayout.PerformLayout()
                $mainLayout.PerformLayout()

                if (!$form -or !$supplierList -or !$replyFolderBox -or !$prepareButton) {
                    throw "Required launcher controls were not created"
                }
                if ($supplierList.Width -lt 350 -or $replyFolderBox.Width -lt 180 -or $prepareButton.Height -lt 44) {
                    throw "Core controls are clipped at the minimum supported window size"
                }
                if ($titleLabel.Right -gt $headerMetaLayout.Left) {
                    throw "Header title overlaps system metadata at the minimum supported window size"
                }
            } finally {
                if ($timer) { $timer.Stop() }
                if ($form) { $form.Dispose() }
            }
        } $runtimeSource
    } catch {
        $failures.Add("Launcher must initialize and lay out without runtime errors: $($_.Exception.Message)")
    }
}

if ($failures.Count) {
    $failures | ForEach-Object { Write-Host "FAIL: $_" -ForegroundColor Red }
    throw "$($failures.Count) launcher UI contract test(s) failed"
}

Write-Host "PASS: MR Launcher professional UI contract" -ForegroundColor Green
