$root = Split-Path -Parent $PSScriptRoot
$outlookPath = Join-Path $root "Scripts\MR-Outlook.ps1"

$tokens = $null
$parserErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($outlookPath, [ref]$tokens, [ref]$parserErrors)
if ($parserErrors.Count) { throw "MR-Outlook.ps1 has parser errors" }

$assignment = $ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
    $node.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
    $node.Left.VariablePath.UserPath -eq "visibleSupplierCols"
}, $true) | Select-Object -First 1

if (!$assignment) { throw "Visible supplier attachment column list is missing" }
$actualHeaders = @($assignment.Right.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.StringConstantExpressionAst]
}, $true) | ForEach-Object { $_.Value })

$expectedHeaders = @(
    "Change Type",
    "Item No",
    "Order Remark",
    "Request Qty",
    "Part No.",
    "Shortage Qty",
    "Material need by date",
    "ETA Vendor can Supply",
    "MC Remarks for MR",
    "Part lead time",
    "Usage",
    "Part Desc",
    "Vendor Name",
    "_SourceFile",
    "_SourceSheet",
    "_SourceRow",
    "IF CAN NOT"
)

if ((@($actualHeaders | Sort-Object) -join "|") -ne (@($expectedHeaders | Sort-Object) -join "|")) {
    throw "Visible attachment columns mismatch. Actual: $($actualHeaders -join ', ')"
}

$source = Get-Content -Raw -LiteralPath $outlookPath
if ($source -notmatch '\$visibleSupplierIndexes\s*=\s*columnIndexes\s+\$outHeaders\s+\$visibleSupplierCols') {
    throw "Visible attachment columns must be resolved from the generated headers"
}
if ($source -notmatch '\$visibleSupplierIndexes\.Values\s+-notcontains\s+\$columnIndex[\s\S]*?\.Hidden\s*=\s*\$true') {
    throw "Every attachment column outside the visible list must be hidden"
}

Write-Host "PASS: Supplier attachment shows the requested 17 columns and hides all remaining columns"
