param(
    [string]$ScriptPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "Scripts\MR-Outlook.ps1")
)

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw "MR-Outlook.ps1 has parser errors" }

foreach ($name in @("findSentMailBySubject", "newReminderReply")) {
    $functionAst = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    }, $true)
    if (!$functionAst) { throw "Function not found: $name" }
    Invoke-Expression $functionAst.Extent.Text
}

$subject = "MR_REQUEST|SUPPLIER A|20260717_101500"
$script:replyAllCount = 0
$original = [pscustomobject]@{ Id="newest"; Subject=$subject; SentOn=[datetime]"2026-07-17T10:15:00" }
$original | Add-Member -MemberType ScriptMethod -Name ReplyAll -Value {
    $script:replyAllCount++
    [pscustomobject]@{ SourceId=$this.Id; Subject="RE: $($this.Subject)" }
}
$older = [pscustomobject]@{ Id="older"; Subject=$subject; SentOn=[datetime]"2026-07-16T10:15:00" }
$unrelated = [pscustomobject]@{ Id="other"; Subject="MR_REQUEST|OTHER|20260717_101500"; SentOn=[datetime]"2026-07-17T11:00:00" }
$items = @($unrelated, $original, $older)

$mail = newReminderReply $items $subject ([datetime]"2026-07-01")
if (!$mail -or $mail.SourceId -ne "newest" -or $script:replyAllCount -ne 1) {
    throw "Reminder was not created with ReplyAll on the exact tracked Subject"
}
if (newReminderReply $items "MR_REQUEST|MISSING|20260717_101500" ([datetime]"2026-07-01")) {
    throw "Missing original mail should not fall back to a new conversation"
}

$scriptText = Get-Content -LiteralPath $ScriptPath -Raw
$remindBlock = [regex]::Match($scriptText, '(?s)if\(\$Mode -eq "remind"\)\{.*\}\s*$').Value
if (!$remindBlock) { throw "MR Remind block not found" }
if ($remindBlock -match '\.CreateItem\(0\)') { throw "MR Remind still creates a new mail conversation" }
if ($remindBlock -notmatch 'newReminderReply\s+\$items') { throw "MR Remind does not use the tracked Subject reply helper" }

Write-Host "PASS: MR Remind replies in the original Sent Items thread by tracked Subject"
