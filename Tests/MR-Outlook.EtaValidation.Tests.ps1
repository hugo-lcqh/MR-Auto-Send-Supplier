param(
    [string]$ScriptPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "Scripts\MR-Outlook.ps1")
)

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw "MR-Outlook.ps1 has parser errors" }

foreach ($name in @("norm", "headerIndex", "etaDateValue", "validEtaDate", "normalizeEtaDate", "configureEtaInput", "invalidEtaRows")) {
    $functionAst = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    }, $true)
    if (!$functionAst) { throw "Function not found: $name" }
    Invoke-Expression $functionAst.Extent.Text
}

foreach ($value in @("31/12/2026", "29/02/2028", "22/2/2026", "2/2/2026")) {
    if (!(validEtaDate $value)) { throw "Valid ETA was rejected: $value" }
}
foreach ($value in @("31/02/2026", "2/22/2026", "2026-12-31", "")) {
    if (validEtaDate $value) { throw "Invalid ETA was accepted: $value" }
}
if ((normalizeEtaDate "22/2/2026") -ne "22/02/2026") {
    throw "Day-first ETA was not normalized to DD/MM/YYYY"
}

$xl = $null
$wb = $null
$ws = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $headers = @("Part No.", "ETA Vendor can Supply")
    $ws.Cells.Item(1, 1).Value2 = $headers[0]
    $ws.Cells.Item(1, 2).Value2 = $headers[1]

    configureEtaInput $ws $headers 6
    $etaRange = $ws.Range("B2:B6")
    if ($etaRange.NumberFormat -ne "@") { throw "ETA cells must preserve day-first input as text" }
    if ($etaRange.Validation.Type -ne 7) { throw "ETA validation is not Custom type" }
    if ($etaRange.Validation.AlertStyle -ne 1) { throw "ETA validation does not use Stop alert" }
    if (!$etaRange.Validation.IgnoreBlank) { throw "Blank ETA should remain allowed" }
    if (!$etaRange.Validation.ShowError) { throw "ETA validation error alert is disabled" }

    $ws.Cells.Item(2, 2).Value2 = "22/2/2026"
    $ws.Cells.Item(3, 2).Value2 = "22/02/2026"
    $ws.Cells.Item(4, 2).Value2 = "2/22/2026"
    $ws.Cells.Item(5, 2).Value2 = "31/2/2026"
    $ws.Cells.Item(6, 2).Value2 = ""
    $xl.CalculateFull()
    if (!$ws.Cells.Item(2, 2).Validation.Value) { throw "Excel rejected valid day-first ETA 22/2/2026" }
    if ($ws.Cells.Item(4, 2).Validation.Value) { throw "Excel accepted invalid month-first ETA 2/22/2026" }
    $badRows = @(invalidEtaRows $ws 2 6)
    if ($badRows.Count -ne 2 -or $badRows[0] -ne 4 -or $badRows[1] -ne 5) {
        throw "Expected rows 4 and 5 to have invalid ETA; got: $($badRows -join ', ')"
    }
} finally {
    if ($wb) { $wb.Close($false) }
    if ($xl) { $xl.Quit() }
    if ($ws) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) }
    if ($wb) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) }
    if ($xl) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) }
}

Write-Host "PASS: ETA accepts day-first dates independent of Excel regional settings"
