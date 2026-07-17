param(
    [string]$ScriptPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "Scripts\MR-Outlook.ps1")
)

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw "MR-Outlook.ps1 has parser errors" }

foreach ($name in @("norm", "headerIndex", "columnIndexes")) {
    $functionAst = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    }, $true)
    if (!$functionAst) { throw "Function not found: $name" }
    Invoke-Expression $functionAst.Extent.Text
}

$templateHeaders = @(1..63 | ForEach-Object { "TemplateColumn$_" }) # A:BK
$traceHeaders = @("_SourceKey", "_SourceFile", "_SourceSheet", "_SourceRow")
$indexes = columnIndexes @($templateHeaders + $traceHeaders) $traceHeaders
$expected = @{ _SourceKey = 64; _SourceFile = 65; _SourceSheet = 66; _SourceRow = 67 } # BL:BO

foreach ($name in $traceHeaders) {
    if ($indexes[$name] -ne $expected[$name]) {
        throw "$name expected column $($expected[$name]), got $($indexes[$name])"
    }
}

try {
    columnIndexes $templateHeaders $traceHeaders | Out-Null
    throw "Missing trace columns were not rejected"
} catch {
    if ($_.Exception.Message -notlike "Missing column: _SourceKey*") { throw }
}

Write-Host "PASS: trace columns move to BL:BO after template columns A:BK"
