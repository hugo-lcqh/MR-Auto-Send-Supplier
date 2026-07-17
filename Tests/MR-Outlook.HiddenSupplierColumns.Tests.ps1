param(
    [string]$ScriptPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "Scripts\MR-Outlook.ps1")
)

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw "MR-Outlook.ps1 has parser errors" }

$assignment = $ast.Find({
    param($node)
    $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $node.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
        $node.Left.VariablePath.UserPath -eq "hiddenSupplierCols"
}, $true)
if (!$assignment) { throw "Variable not found: hiddenSupplierCols" }

$traceCols = @("_SourceKey", "_SourceFile", "_SourceSheet", "_SourceRow")
$hiddenSupplierCols = Invoke-Expression $assignment.Right.Extent.Text
$required = @(
    "Item Seq", "Follow By Mps", "Change Type", "Adjust date", "Batch No", "Phase Code",
    "Item No", "Master pack", "Item Type", "Description", "Customer Item", "End Buyer",
    "Customer Name", "Region", "Customer PO Number", "Order Number",
    "Customer Due Date", "Original Qty", "Original Promise Date", "Original Schedule Ship Date",
    "Request Qty", "Request Date", "Order Priority", "New Request Qty", "New Request Date",
    "Reply Qty", "Reply Promise Date", "Reply Schedule Date", "New Priority", "Material Ready Date",
    "Part Remarks", "MPS Remark"
)

$missing = @($required + $traceCols | Where-Object { $hiddenSupplierCols -notcontains $_ })
if ($missing.Count) { throw "Hidden supplier columns missing: $($missing -join ', ')" }

$visible = @("Part Lead Time", "Usage", "Order Remark", "OFIS Line ID")
$incorrectlyHidden = @($visible | Where-Object { $hiddenSupplierCols -contains $_ })
if ($incorrectlyHidden.Count) { throw "Supplier columns should remain visible: $($incorrectlyHidden -join ', ')" }

$duplicates = @($hiddenSupplierCols | Group-Object | Where-Object Count -gt 1 | Select-Object -ExpandProperty Name)
if ($duplicates.Count) { throw "Duplicate hidden supplier columns: $($duplicates -join ', ')" }

Write-Host "PASS: Send MR hides all requested supplier columns and trace columns"
