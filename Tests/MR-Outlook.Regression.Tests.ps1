param([string]$SupplierPath)

$root = Split-Path -Parent $PSScriptRoot
$outlookPath = Join-Path $root "Scripts\MR-Outlook.ps1"
$launcherPath = Join-Path $root "Scripts\MR-Launcher.ps1"
$localSupplierPath = Join-Path $root "Config\suppliers.csv"
$sampleSupplierPath = Join-Path $root "Config\suppliers.example.csv"
$supplierPath = if ($SupplierPath) {
    if ([IO.Path]::IsPathRooted($SupplierPath)) { $SupplierPath } else { Join-Path $root $SupplierPath }
} elseif (Test-Path -LiteralPath $localSupplierPath) {
    $localSupplierPath
} else {
    $sampleSupplierPath
}

$tokens = $null
$parserErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($outlookPath, [ref]$tokens, [ref]$parserErrors)
$source = Get-Content -Raw -LiteralPath $outlookPath
$launcher = Get-Content -Raw -LiteralPath $launcherPath
$launcherTokens = $null
$launcherErrors = $null
$launcherAst = [System.Management.Automation.Language.Parser]::ParseFile($launcherPath, [ref]$launcherTokens, [ref]$launcherErrors)
$failures = [System.Collections.Generic.List[string]]::new()

function Assert-True($condition, [string]$message) {
    if (!$condition) { $failures.Add($message) }
}

function Get-AssignedStrings([string]$variableName) {
    $assignment = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $node.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
        $node.Left.VariablePath.UserPath -eq $variableName
    }, $true) | Select-Object -First 1

    if (!$assignment) { return @() }
    @($assignment.Right.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.StringConstantExpressionAst]
    }, $true) | ForEach-Object { $_.Value })
}

Assert-True ($parserErrors.Count -eq 0) "MR-Outlook.ps1 must parse without errors"

$columnFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "excelColumnName"
}, $true) | Select-Object -First 1
Assert-True ($null -ne $columnFunction) "excelColumnName must generate dynamic Excel column labels"
if ($columnFunction) {
    Invoke-Expression $columnFunction.Extent.Text
    Assert-True ((excelColumnName 1) -eq "A") "Column 1 must be A"
    Assert-True ((excelColumnName 38) -eq "AL") "Column 38 must be AL"
    Assert-True ((excelColumnName 68) -eq "BP") "Column 68 must be BP"
}

$visibleColumns = @(Get-AssignedStrings "visibleSupplierCols")
$importantColumns = @(Get-AssignedStrings "importantCols")
$emailTableColumns = @(Get-AssignedStrings "emailTableColumns")
$requiredVisibleColumns = @(
    "Change Type","Item No","Order Remark","Request Qty","Part No.","Shortage Qty",
    "Material need by date","ETA Vendor can Supply","MC Remarks for MR","Part lead time",
    "Usage","Part Desc","Vendor Name","_SourceFile","_SourceSheet","_SourceRow","IF CAN NOT"
)
Assert-True ((@($visibleColumns | Sort-Object) -join "|") -eq (@($requiredVisibleColumns | Sort-Object) -join "|")) "Supplier attachment must show exactly the requested 17 columns"
Assert-True ($source -match '\$visibleSupplierIndexes\.Values\s+-notcontains\s+\$columnIndex[\s\S]*?\.Hidden\s*=\s*\$true') "Supplier attachment must hide every column outside the requested 17 columns"
Assert-True ($importantColumns -contains "Request Qty") "Supplier attachment must highlight Request Qty"
Assert-True ($importantColumns -notcontains "Shortage Qty") "Supplier attachment must not highlight hidden Shortage Qty"
Assert-True ($emailTableColumns -contains "Request Qty") "Email summary table must include Request Qty"
Assert-True ($emailTableColumns -notcontains "Shortage Qty") "Email summary table must not include Shortage Qty"
Assert-True ($source -match '\$etaRange\.NumberFormat\s*=\s*"dd-mmm-yyyy"') "ETA column AL must display an unambiguous date such as 29-Jul-2026"
Assert-True ($source -match '\$materialNeedRange\.NumberFormat\s*=\s*"d-mmm-yy"') "Supplier attachment must display Material need by date as 6-Sep-26"
Assert-True ($source -match 'if\(isMaterialNeedDateHeader \$h\)\{return "d-mmm-yy"\}') "Master workbook must display Material need by date as 6-Sep-26"
Assert-True ($source -match 'MATERIAL NEED BY DATE"\)\{\$value=materialNeedDateText \$value;\$value=excelDateValue \$value\}') "Supplier attachment must write Material need by date as a real Excel date"
Assert-True ($source -match 'supplierEmailTable -rows @\(\$group\.Group\) -columns \$emailTableColumns') "Send mode must build an HTML table from supplier rows"
Assert-True ($source -match '\$emailTable\s*\r?\n') "Generated email body must include the supplier data table"
Assert-True ($source -match '\[switch\]\$HideEmailTable') "Send mode must accept the option to hide the email summary table"
Assert-True ($source -match '\$emailTable\s*=\s*if\(\$HideEmailTable\)\s*\{\s*""\s*\}\s*else\s*\{\s*supplierEmailTable') "Send mode must omit the email table when the user disables it"
Assert-True ($source -match 'foreach\(\$mcAddress in \$mcAddresses\)' -and $source -match 'Recipients\.Add\(\$mcAddress\)') "Each supplier loop must resolve every configured MC recipient"
Assert-True ($source -match "href='mailto:") "MC display tag must link to the resolved MC address"
Assert-True ($source -match 'MC tag resolved:') "Each supplier loop must report the resolved MC tag"
Assert-True ($source -match 'GetExchangeUser\(\)') "MC display tag should prefer the compact Outlook directory name"
Assert-True ($source -match '<b>Action required:</b>') "Supplier email must highlight the required attachment action"
Assert-True ($source -match 'reply to this same email thread with the completed file attached') "Supplier email must ask for the completed file in the same mail loop"

$compactMcNameFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "compactMcName"
}, $true) | Select-Object -First 1
Assert-True ($compactMcNameFunction) "MC compact display-name helper must be available"
if ($compactMcNameFunction) {
    Invoke-Expression $compactMcNameFunction.Extent.Text
    Assert-True ((compactMcName "ChiQuocHung.Le@ttigroup.com.vn" "ChiQuocHung.Le@ttigroup.com.vn") -eq "Chi Quoc Hung Le") "Email fallback must display a compact MC name instead of the full address"
    Assert-True ((compactMcName "Hugo Le Chi Quoc Hung (VN.OP-SUC)" "ChiQuocHung.Le@ttigroup.com.vn") -eq "Hugo Le Chi Quoc Hung (VN.OP-SUC)") "Resolved Outlook MC display name must be preserved"
}

$columnLookupFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "col"
}, $true) | Select-Object -First 1
Assert-True ($columnLookupFunction) "Column lookup helper must be available"
if ($columnLookupFunction) {
    Invoke-Expression $columnLookupFunction.Extent.Text
    $supplierWithSpacedHeader = [pscustomobject]@{ "MC " = "mc@example.com" }
    Assert-True ((col $supplierWithSpacedHeader @("MC") $false) -eq "MC ") "Column lookup must accept trailing spaces in suppliers.csv headers"
}

$etaDateFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "etaDateValue"
}, $true) | Select-Object -First 1
$normalizeEtaFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "normalizeEtaDate"
}, $true) | Select-Object -First 1
$excelEtaFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "excelEtaValue"
}, $true) | Select-Object -First 1
Assert-True ($etaDateFunction -and $normalizeEtaFunction -and $excelEtaFunction) "ETA date parsing helpers must be available"
if ($etaDateFunction -and $normalizeEtaFunction -and $excelEtaFunction) {
    Invoke-Expression $etaDateFunction.Extent.Text
    Invoke-Expression $normalizeEtaFunction.Extent.Text
    Invoke-Expression $excelEtaFunction.Extent.Text
    $etaOaDate = ([datetime]"2026-07-29").ToOADate()
    Assert-True ((normalizeEtaDate $etaOaDate) -eq "29-Jul-2026") "Excel numeric ETA dates must normalize as DD-MMM-YYYY"
    Assert-True ((normalizeEtaDate "29/07/2026") -eq "29-Jul-2026") "Day-first ETA text must normalize as DD-MMM-YYYY"
    Assert-True ((excelEtaValue "29/07/2026") -eq $etaOaDate) "Existing ETA text must be written back to Excel as a real date"
}

$excelDateFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "excelDateValue"
}, $true) | Select-Object -First 1
$materialNeedDateFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "materialNeedDateText"
}, $true) | Select-Object -First 1
Assert-True ($excelDateFunction -and $materialNeedDateFunction) "Material need by date formatter must be available"
if ($excelDateFunction -and $materialNeedDateFunction) {
    Invoke-Expression $excelDateFunction.Extent.Text
    Invoke-Expression $materialNeedDateFunction.Extent.Text
    $materialNeedOaDate = ([datetime]"2026-09-06").ToOADate()
    Assert-True ((materialNeedDateText $materialNeedOaDate) -eq "6-Sep-26") "Prepare must normalize Excel dates as d-MMM-yy"
    Assert-True ((materialNeedDateText "06/09/2026") -eq "6-Sep-26") "Prepare must normalize day-first date text as d-MMM-yy"
}

$supplierTableFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "supplierEmailTable"
}, $true) | Select-Object -First 1
$valueFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "val"
}, $true) | Select-Object -First 1
Assert-True ($supplierTableFunction -and $valueFunction) "Supplier email table helper must be available"
if ($supplierTableFunction -and $valueFunction) {
    Invoke-Expression $valueFunction.Extent.Text
    Invoke-Expression $supplierTableFunction.Extent.Text
    $tableRows = @([pscustomobject]@{ "Request Qty" = "900"; "Part No." = "A&B <tag>" })
    $tableColumns = @(
        [pscustomobject]@{ Header = "Request Qty"; Names = @("Request Qty") },
        [pscustomobject]@{ Header = "Part No."; Names = @("Part No.") }
    )
    $tableHtml = supplierEmailTable $tableRows $tableColumns
    Assert-True ($tableHtml -match '<table' -and $tableHtml -match 'Request Qty' -and $tableHtml -match '>900<') "Email summary must render supplier data as an HTML table"
    Assert-True ($tableHtml -match 'A&amp;B &lt;tag&gt;') "Email summary must HTML-encode workbook values"
}

Assert-True ($source -match '\$xl\.AutomationSecurity\s*=\s*3') "Excel automation must force-disable macros"
Assert-True ($source -match 'Workbooks\.Open\(\$path\)') "Reply attachments must open through the compatible Excel COM overload"
Assert-True ($source -match 'for\(\$attachmentIndex\s*=\s*1') "Attachment loop needs its own counter"
Assert-True ($source -match 'for\(\$headerIndex\s*=\s*0') "Header loop needs a separate counter"
Assert-True ($source -match '\$extension\s+-notmatch\s+\x27\^\(\\\.xls\[xm\]\?\|\\\.csv\)\$\x27') "Scan must accept CSV reply attachments as well as Excel files"
Assert-True ($source -notmatch '\$_\.Vendor\s*-eq\s*\$replyVendor\s*-or') "A reply must not close every request for the same vendor"

Assert-True ($source -match '\$range\.Value2\s*=\s*\$arr') "Supplier rows must be written to Excel as one range"
Assert-True ($source -match '\$arr\[\(\$rowIndex\+1\),\$columnIndex\]') "Two-dimensional supplier array indexes must parenthesize row arithmetic"
Assert-True ($source -notmatch '\$arr\[\$rowIndex\+1,\$columnIndex\]') "Two-dimensional supplier array indexes must not trigger PowerShell op_Addition"
Assert-True ($source -match '\$failedCount\s*=\s*0' -and $source -match '\$failedCount\+\+') "Send mode must count supplier processing failures"
Assert-True ($source -match 'if\(\$failedCount -gt 0\)[\s\S]*?"WARNING"') "Send summary must warn instead of reporting OK when suppliers fail"
Assert-True ($source -match 'if\(\!\$Display\)\s*\{[\s\S]*?appendSentRecord') "Previewed mail must not be recorded as sent"
Assert-True ($source -match 'if\(\!\$Display\)\s*\{[\s\S]*?\$request\.Reminded') "Previewed reminder must not be recorded as reminded"

$masterWritableFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "assertMasterFileWritable"
}, $true) | Select-Object -First 1
$masterOutputFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "writeMasterOutputs"
}, $true) | Select-Object -First 1
Assert-True ($masterWritableFunction) "Scan must preflight master output files before reading replies"
Assert-True ($masterOutputFunction) "Scan must publish text and Excel master outputs through one writer"
Assert-True ($source -match 'writeMasterOutputs\s+\$data\s+\$InputFile') "Scan must use the atomic master writer after updating rows"
Assert-True ($source -match 'exportMasterXlsx\s+\$stagingTxt\s+\$stagingXlsx\s+\$templatePath\s+\$headers\s+\$displayHeaders\s+\$delimiter') "Master Excel staging must preserve the source delimiter"
Assert-True ($source -notmatch 'Remove-Item -LiteralPath \$xlsx -Force') "Master export must not delete the existing Excel file before publishing"
if ($masterWritableFunction) {
    Invoke-Expression $masterWritableFunction.Extent.Text
    $lockedMasterPath = Join-Path ([IO.Path]::GetTempPath()) ("mr-master-lock-{0}.txt" -f [guid]::NewGuid().ToString("N"))
    Set-Content -LiteralPath $lockedMasterPath -Value "master" -Encoding UTF8
    $lockStream = $null
    try {
        $lockStream = [IO.File]::Open($lockedMasterPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $lockDetected = $false
        try { assertMasterFileWritable $lockedMasterPath } catch { $lockDetected = $true }
        Assert-True ($lockDetected) "Scan must stop clearly when the master file is locked"
    } finally {
        if ($lockStream) { $lockStream.Dispose() }
        Remove-Item -LiteralPath $lockedMasterPath -Force -ErrorAction SilentlyContinue
    }
}
if ($masterOutputFunction) {
    $exportTsvFunction = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "exportTsv"
    }, $true) | Select-Object -First 1
    if ($exportTsvFunction) { Invoke-Expression $exportTsvFunction.Extent.Text }
    Invoke-Expression $masterOutputFunction.Extent.Text
    function log($message, $level = "INFO") { }
    function exportMasterXlsx($txtPath, $xlsxPath, $templatePath, $headers, $displayHeaders, $delimiter = "`t") {
        Copy-Item -LiteralPath $txtPath -Destination $xlsxPath -Force
    }
    $writerRoot = Join-Path ([IO.Path]::GetTempPath()) ("mr-master-writer-{0}" -f [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $writerRoot -Force | Out-Null
    $writerTxt = Join-Path $writerRoot "MR_Master_Input.txt"
    $writerXlsx = Join-Path $writerRoot "MR_Master_Input.xlsx"
    try {
        $writerRows = @([pscustomobject]@{ Vendor = "Test Supplier"; "Reply Qty" = "42" })
        writeMasterOutputs $writerRows $writerTxt $writerXlsx "" @("Vendor", "Reply Qty") @("Vendor", "Reply Qty") "`t"
        Assert-True ((Test-Path -LiteralPath $writerTxt) -and (Test-Path -LiteralPath $writerXlsx)) "Master writer must publish both output files"
        Assert-True ((Get-Content -Raw -LiteralPath $writerTxt) -match "Test Supplier" -and (Get-Content -Raw -LiteralPath $writerXlsx) -match "Test Supplier") "Master writer must publish the updated data to both outputs"
    } finally {
        Remove-Item -LiteralPath $writerRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Assert-True ($launcher -match 'VendorSelection') "Launcher must pass selected suppliers to one worker"
Assert-True ($launcher -notmatch '\$selected\s*\|\s*ForEach-Object\s*\{\s*run-one') "Launcher must not start one PowerShell process per supplier"
Assert-True ($launcher -notmatch 'function start-run-monitor|function new-run-monitor|RUN MONITOR') "Launcher must not use the popup run monitor"
Assert-True ($launcher -notmatch 'RedirectStandardOutput|RedirectStandardError') "Launcher must not redirect PowerShell output through background callbacks"
Assert-True ($launcher -match 'Start-Process -FilePath "powershell\.exe"') "Launcher tasks must run in a normal PowerShell terminal"
Assert-True ($source -match 'function terminalBanner') "Terminal output must show a task banner"
Assert-True ($source -match 'ForegroundColor') "Terminal output must use readable status colors"
Assert-True ($source -match 'function terminalProgress' -and $source -match 'Write-Progress') "Terminal output must show live task progress"
Assert-True ($source -match 'Write-Progress[\s\S]*-Completed') "Terminal progress must close cleanly after the task"
Assert-True ($source -match 'Duration:') "Terminal result must show elapsed duration"
Assert-True ($source -match 'MATERIAL REQUEST CONTROL CENTER') "Terminal banner must show the control-center identity"
Assert-True ($source -match 'Get-Date -Format "HH:mm:ss"') "Terminal logs must include compact timestamps"

$terminalTaskTitleFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "terminalTaskTitle"
}, $true) | Select-Object -First 1
$terminalProgressTextFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "terminalProgressText"
}, $true) | Select-Object -First 1
Assert-True ($terminalTaskTitleFunction -and $terminalProgressTextFunction) "Terminal presentation helpers must be available"
if ($terminalTaskTitleFunction -and $terminalProgressTextFunction) {
    Invoke-Expression $terminalTaskTitleFunction.Extent.Text
    Invoke-Expression $terminalProgressTextFunction.Extent.Text
    Assert-True ((terminalTaskTitle "send") -eq "SEND SUPPLIER MR") "Send mode must have a clear terminal mission title"
    Assert-True ((terminalTaskTitle "prepare") -eq "PREPARE MASTER DATA") "Prepare mode must have a clear terminal mission title"
    Assert-True ((terminalProgressText "Processing supplier" 5 10) -eq "[ 50% ] Processing supplier") "Terminal progress must show a readable percentage badge"
    Assert-True ((terminalProgressText "Done" 12 10) -eq "[100% ] Done") "Terminal progress percentage must clamp at 100"
}
Assert-True ($launcher -match '-Mode prepare') "Launcher must define the prepare task mode"
foreach ($mode in @('send', 'scan', 'remind')) {
    Assert-True ($launcher -match ('"' + $mode + '"')) "Launcher must define the $mode task mode"
}
Assert-True ($launcher -match 'function run-one' -and $launcher -match '"-NoExit"') "Non-prepare tasks must stay visible in the terminal after completion"
Assert-True ($launcher -match '\$emailTableBox\.Text\s*=\s*"Show MR table in email"') "Launcher must offer an email-table checkbox"
Assert-True ($launcher -match '\$emailTableBox\.Checked\s*=\s*\$true') "Email-table checkbox must preserve the current enabled behavior by default"
Assert-True ($launcher -match 'if\s*\(\$mode -eq "send" -and !\$emailTableBox\.Checked\)\s*\{\s*\$args \+= "-HideEmailTable"\s*\}') "Launcher must hide the table only when Send MR is run with the checkbox cleared"

$encodeFunction = $launcherAst.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "encode-vendor-selection"
}, $true) | Select-Object -First 1
$decodeFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "decodeVendorSelection"
}, $true) | Select-Object -First 1
Assert-True ($null -ne $encodeFunction -and $null -ne $decodeFunction) "Vendor selection encoding must have a tested decoder"
if ($encodeFunction -and $decodeFunction) {
    Invoke-Expression $encodeFunction.Extent.Text
    Invoke-Expression $decodeFunction.Extent.Text
    $vendorNames = @("SUPPLIER A", "SUPPLIER, B", "SUPPLIER (C)")
    $decodedVendorNames = @(decodeVendorSelection (encode-vendor-selection $vendorNames))
    Assert-True (($decodedVendorNames -join '|') -eq ($vendorNames -join '|')) "Vendor selection must round-trip without changing names"
}

$normFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "norm"
}, $true) | Select-Object -First 1
$addSupplierKeyFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "addSupplierKey"
}, $true) | Select-Object -First 1
$findSupplierFunction = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq "findSupplier"
}, $true) | Select-Object -First 1
Assert-True ($normFunction -and $addSupplierKeyFunction -and $findSupplierFunction) "Supplier matching helpers must be available for prefix matching"
if ($normFunction -and $addSupplierKeyFunction -and $findSupplierFunction) {
    Invoke-Expression $normFunction.Extent.Text
    Invoke-Expression $addSupplierKeyFunction.Extent.Text
    Invoke-Expression $findSupplierFunction.Extent.Text
    $supplierExact = @{}
    $supplierAliases = @{}
    $weidaRow = [pscustomobject]@{ VendorName = "WEIDA (VIETNAM) MANUFACTURING CO.,LTD"; "Email To" = "weida@example.com" }
    $yaoRow = [pscustomobject]@{ VendorName = "YAO-I VIETNAM COMPANY LIMITED"; "Email To" = "yao@example.com" }
    addSupplierKey $supplierExact $weidaRow.VendorName $weidaRow
    addSupplierKey $supplierExact $yaoRow.VendorName $yaoRow
    Assert-True ((findSupplier "WEIDA (VIETNAM)" $supplierExact $supplierAliases).VendorName -eq $weidaRow.VendorName) "Unique supplier prefix must resolve WEIDA"
    Assert-True ((findSupplier "YAO-I" $supplierExact $supplierAliases).VendorName -eq $yaoRow.VendorName) "Unique supplier prefix must resolve YAO-I"
    $ambiguousRow = [pscustomobject]@{ VendorName = "YAO-I OTHER COMPANY LIMITED"; "Email To" = "other@example.com" }
    addSupplierKey $supplierExact $ambiguousRow.VendorName $ambiguousRow
    Assert-True ($null -eq (findSupplier "YAO-I" $supplierExact $supplierAliases)) "Ambiguous supplier prefix must not auto-select a recipient"
}

$supplierRows = @(Import-Csv -LiteralPath $supplierPath)
$supplierHeaders = @($supplierRows[0].PSObject.Properties.Name)
$supplierMcColumn = col $supplierRows[0] @("MC", "MC in charge", "MC Email") $false
Assert-True ($supplierMcColumn) "Supplier configuration must expose an MC column even when its header has trailing spaces"
$emailColumns = @(
    (@("Email To", "EmailTo", "To") | Where-Object { $supplierHeaders -contains $_ } | Select-Object -First 1),
    (@("Email CC", "EmailCC", "CC") | Where-Object { $supplierHeaders -contains $_ } | Select-Object -First 1),
    $supplierMcColumn
) | Where-Object { $_ }
for ($supplierIndex = 0; $supplierIndex -lt $supplierRows.Count; $supplierIndex++) {
    $supplierRow = $supplierRows[$supplierIndex]
    foreach ($emailColumn in $emailColumns) {
        foreach ($address in ([string]$supplierRow.$emailColumn -split ';')) {
            $address = $address.Trim()
            if (!$address) { continue }
            try { [void][Net.Mail.MailAddress]::new($address) }
            catch { $failures.Add("Supplier $emailColumn must be valid at row $($supplierIndex + 2)") }
        }
    }
}

if ($failures.Count) {
    $failures | ForEach-Object { Write-Host "FAIL: $_" -ForegroundColor Red }
    throw "$($failures.Count) MR-Outlook regression test(s) failed"
}

Write-Host "PASS: MR-Outlook regression checks" -ForegroundColor Green
