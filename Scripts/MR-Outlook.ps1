param(
    [ValidateSet("send","scan","remind","prepare")][string]$Mode="send",
    [string]$Vendor,
    [Alias("Input")][string]$InputFile="Input\MR-Outlook\input.txt",
    [string]$InputRoot="Input",
    [string]$Template="Input\Template.xlsx",
    [string]$Supplier="Config\suppliers.csv",
    [string]$OutDir="MR_Out",
    [string]$ReplyFolder="MR_REQUEST",
    [string]$VendorSelection,
    [switch]$Display,
    [switch]$HideEmailTable
)

$ErrorActionPreference="Stop"
$Root=if($env:AUTOTOOLS_ROOT){$env:AUTOTOOLS_ROOT}else{Split-Path -Parent $PSScriptRoot}
if(![IO.Path]::IsPathRooted($InputFile)){$InputFile=Join-Path $Root $InputFile}
if(![IO.Path]::IsPathRooted($InputRoot)){$InputRoot=Join-Path $Root $InputRoot}
if(![IO.Path]::IsPathRooted($Template)){$Template=Join-Path $Root $Template}
if(![IO.Path]::IsPathRooted($Supplier)){$Supplier=Join-Path $Root $Supplier}
if(![IO.Path]::IsPathRooted($OutDir)){$OutDir=Join-Path $Root $OutDir}
function col($o,$a,$req=$true){
    foreach($n in $a){$p=$o.PSObject.Properties[$n];if($p){return $p.Name}}
    foreach($p in $o.PSObject.Properties){foreach($n in $a){if([string]::Equals(([string]$p.Name).Trim(),([string]$n).Trim(),[StringComparison]::OrdinalIgnoreCase)){return $p.Name}}}
    if($req){throw "Missing column: $($a -join ', ')"}
}
function val($o,$a){foreach($n in $a){$p=$o.PSObject.Properties[$n];if($p){return [string]$p.Value}};return ""}
function compactMcName($resolvedName,$address){
    $name=([string]$resolvedName).Trim()
    $email=([string]$address).Trim()
    if($name -and $name -notmatch '@' -and ![string]::Equals($name,$email,[StringComparison]::OrdinalIgnoreCase)){return $name}
    $local=($email -split '@')[0]
    $local=($local -creplace '(?<=[a-z])(?=[A-Z])',' ' -replace '[._-]+',' ').Trim()
    if(!$local){return "MC"}
    return [Globalization.CultureInfo]::InvariantCulture.TextInfo.ToTitleCase($local.ToLowerInvariant())
}
function supplierEmailTable($rows,$columns){
    $rows=@($rows)
    if(!$rows.Count){return ""}
    $html=[Text.StringBuilder]::new()
    [void]$html.Append('<div style="margin:14px 0 6px;"><b>MR details:</b></div><div style="overflow-x:auto;"><table style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:10pt;">')
    [void]$html.Append('<thead><tr>')
    foreach($column in @($columns)){[void]$html.Append(('<th style="background:#002060;color:#fff;border:1px solid #b7c3d0;padding:5px 7px;text-align:left;white-space:nowrap;">{0}</th>' -f [Net.WebUtility]::HtmlEncode([string]$column.Header)))}
    [void]$html.Append('</tr></thead><tbody>')
    foreach($row in $rows){
        [void]$html.Append('<tr>')
        foreach($column in @($columns)){
            $value=val $row @($column.Names)
            [void]$html.Append(('<td style="border:1px solid #d9d9d9;padding:4px 7px;vertical-align:top;">{0}</td>' -f [Net.WebUtility]::HtmlEncode($value)))
        }
        [void]$html.Append('</tr>')
    }
    [void]$html.Append('</tbody></table></div>')
    return $html.ToString()
}
function terminalTaskTitle($mode){
    switch($mode){
        "prepare"{"PREPARE MASTER DATA";break}
        "send"{"SEND SUPPLIER MR";break}
        "scan"{"SCAN SUPPLIER REPLIES";break}
        "remind"{"SEND FOLLOW-UP REMINDERS";break}
        default{([string]$mode).ToUpperInvariant()}
    }
}
function terminalProgressText($status,$current=0,$total=0){
    $percent=if([int]$total -gt 0){[int][Math]::Min(100,[Math]::Max(0,[Math]::Round(100*([double]$current/[double]$total))))}else{0}
    return ("[{0,3}% ] {1}" -f $percent,$status)
}
function terminalBanner($mode){
    $script:terminalStartedAt=Get-Date
    $title=terminalTaskTitle $mode
    try{$Host.UI.RawUI.WindowTitle="MR Outlook // $title"}catch{}
    $frame="+"+("-"*74)+"+"
    Write-Host ""
    Write-Host $frame -ForegroundColor DarkCyan
    Write-Host ("|  {0,-70}  |" -f "MR // OUTLOOK AUTOMATION") -ForegroundColor Yellow
    Write-Host ("|  {0,-70}  |" -f "MATERIAL REQUEST CONTROL CENTER") -ForegroundColor DarkGray
    Write-Host ("|{0}|" -f ("-"*74)) -ForegroundColor DarkCyan
    Write-Host ("|  {0,-10}{1,-60}  |" -f "MISSION",$title) -ForegroundColor White
    Write-Host ("|  {0,-10}{1,-60}  |" -f "STATUS","ONLINE / RUNNING") -ForegroundColor Green
    Write-Host ("|  {0,-10}{1,-60}  |" -f "STARTED",(Get-Date -Format "dd-MMM-yyyy  HH:mm:ss")) -ForegroundColor DarkGray
    Write-Host $frame -ForegroundColor DarkCyan
    Write-Host ""
}
function terminalProgress($activity,$status,$current=0,$total=0){
    if([int]$total -gt 0){
        $percent=[int][Math]::Min(100,[Math]::Max(0,[Math]::Round(100*([double]$current/[double]$total))))
        Write-Progress -Id 1 -Activity ("MR // {0}" -f ([string]$activity).ToUpperInvariant()) -Status (terminalProgressText $status $current $total) -CurrentOperation ("{0} / {1} complete" -f $current,$total) -PercentComplete $percent
    }else{
        Write-Progress -Id 1 -Activity ("MR // {0}" -f ([string]$activity).ToUpperInvariant()) -Status $status
    }
}
function terminalProgressDone($activity){Write-Progress -Id 1 -Activity $activity -Completed}
function terminalDuration(){
    if($script:terminalStartedAt){return ((Get-Date)-$script:terminalStartedAt).ToString("hh\:mm\:ss")}
    return "00:00:00"
}
function terminalSection($title){
    Write-Host ""
    $label=(" {0} " -f ([string]$title).ToUpperInvariant())
    Write-Host ("  +--{0}{1}+" -f $label,("-"*[Math]::Max(1,67-$label.Length))) -ForegroundColor Cyan
}
function log($m,$level="INFO"){
    $prefix="INFO";$color="Gray"
    if($level -eq "STEP"){$prefix="STEP";$color="Cyan"}
    elseif($level -eq "SUCCESS"){$prefix="PASS";$color="Green"}
    elseif($level -eq "WARNING"){$prefix="WARN";$color="Yellow"}
    Write-Host ("  {0} " -f (Get-Date -Format "HH:mm:ss")) -NoNewline -ForegroundColor DarkGray
    Write-Host ("[{0}]" -f $prefix) -NoNewline -ForegroundColor $color
    Write-Host (" {0}" -f $m)
}
function norm($s){(([string]$s) -replace '\s+',' ').Trim().ToUpperInvariant()}
function releaseCom($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}catch{}}}
function excelColumnName($index){
    $n=[int]$index
    if($n -lt 1){return ""}
    $name=""
    while($n -gt 0){$n--; $name=([char][int](65+($n%26)))+$name; $n=[math]::Floor($n/26)}
    return $name
}
function validEmailList($value){
    $addresses=@(([string]$value -split ';')|%{$_.Trim()}|?{$_})
    if(!$addresses.Count){return $false}
    foreach($address in $addresses){try{[void][Net.Mail.MailAddress]::new($address)}catch{return $false}}
    return $true
}
function appendSentRecord($record,$path){
    $mutex=[Threading.Mutex]::new($false,"Local\MR_Auto_Send_Supplier_Log")
    $locked=$false
    try{
        try{$locked=$mutex.WaitOne(30000)}catch [Threading.AbandonedMutexException]{$locked=$true}
        if(!$locked){throw "Timed out waiting to update $path"}
        if((Test-Path -LiteralPath $path) -and (Get-Item -LiteralPath $path).Length -gt 0){$record|Export-Csv -LiteralPath $path -Append -NoTypeInformation -Encoding UTF8}
        else{$record|Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8}
    }finally{
        if($locked){$mutex.ReleaseMutex()}
        $mutex.Dispose()
    }
}
function addSupplierKey($map,$key,$row){$k=norm $key;if($k -and !$map.ContainsKey($k)){$map[$k]=$row}}
function findSupplier($name,$exact,$aliases){
    $k=norm $name
    if($exact.ContainsKey($k)){return $exact[$k]}
    foreach($a in @($aliases.Keys|Sort-Object Length -Descending)){
        if($k -eq $a -or $k.StartsWith("$a ") -or $k.StartsWith("$a(") -or $k.StartsWith("$a-")){return $aliases[$a]}
    }
    $prefixMatches=@($exact.Keys|?{$_ -like "$k *" -or $_ -like "$k(" -or $_ -like "$k-*"})
    if($prefixMatches.Count -eq 1){return $exact[$prefixMatches[0]]}
    return $null
}
function findSentMailBySubject($items,$subject,$cutoff){
    $wanted=([string]$subject).Trim()
    if(!$wanted){return $null}
    foreach($m in $items){
        try{
            if($cutoff -and [datetime]$m.SentOn -lt [datetime]$cutoff){break}
            if([string]::Equals(([string]$m.Subject).Trim(),$wanted,[StringComparison]::OrdinalIgnoreCase)){return $m}
        }catch{}
    }
    return $null
}
function newReminderReply($items,$subject,$cutoff){
    $original=findSentMailBySubject $items $subject $cutoff
    if(!$original){return $null}
    return $original.ReplyAll()
}
function latestInputFolder($base){Get-ChildItem -LiteralPath $base -Directory|?{$_.Name -match '^\d{2}\.\d{2}\.\d{4}$'}|%{$d=[datetime]::MinValue;if([datetime]::TryParseExact($_.Name,'dd.MM.yyyy',$null,[Globalization.DateTimeStyles]::None,[ref]$d)){[pscustomobject]@{Path=$_.FullName;Date=$d}}}|Sort-Object Date -Descending|Select-Object -First 1 -ExpandProperty Path}
function exportTsv($rows,$path){$rows|Export-Csv -LiteralPath $path -Delimiter "`t" -NoTypeInformation -Encoding UTF8}
function assertMasterFileWritable($path){
    $fullPath=[IO.Path]::GetFullPath([string]$path);$parent=Split-Path -Parent $fullPath
    if(!(Test-Path -LiteralPath $parent)){throw "Master output folder not found: $parent"}
    if(!(Test-Path -LiteralPath $fullPath)){return}
    $stream=$null
    try{$stream=[IO.File]::Open($fullPath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)}catch{throw "Master output is locked or not writable: $fullPath. Close it and run Scan Reply again."}finally{if($stream){$stream.Dispose()}}
}
function decodeVendorSelection($encoded){
    if(!$encoded){return @()}
    try{
        $json=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($encoded))
        $decoded=$json|ConvertFrom-Json
        foreach($name in @($decoded)){[string]$name}
    }catch{throw "Invalid VendorSelection: $($_.Exception.Message)"}
}

$selectedVendorSet=$null
if($VendorSelection){
    $selectedVendorSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    @(decodeVendorSelection $VendorSelection)|%{$key=norm $_;if($key){[void]$selectedVendorSet.Add($key)}}
}
function vendorSelected($name){
    if($selectedVendorSet){return $selectedVendorSet.Contains((norm $name))}
    return (!$Vendor -or ([string]$name) -like "*$Vendor*")
}
function invokeRemind($trackingFile){
    $track=@();if(Test-Path -LiteralPath $trackingFile){$track=@(Import-Csv -LiteralPath $trackingFile)}
    log "Tracking rows: $($track.Count)"
    $ol=$null;$namespace=$null;$sent=$null;$items=$null
    try{
        $ol=New-Object -ComObject Outlook.Application
        $namespace=$ol.GetNamespace("MAPI");$sent=$namespace.GetDefaultFolder(5);$items=$sent.Items;$items.Sort("[SentOn]",$true)
        $known=@{};$track|%{$known[$_.Subject]=$true};$cut=(Get-Date).AddDays(-30)
        foreach($m in $items){
            try{
                if($m.SentOn -lt $cut){break}
                if(([string]$m.Subject) -notlike "MR_REQUEST|*" -or $known[$m.Subject]){continue}
                $vn=([string]$m.Subject) -replace '^MR_REQUEST\|(.+)\|\d{8}_\d{6}.*$','$1'
                $track+=[pscustomobject]@{Vendor=$vn;To=$m.To;CC=$m.CC;Subject=$m.Subject;Sent=([datetime]$m.SentOn).ToString("s");Attachment="";Replied="";Reminded=""}
            }catch{Write-Warning "Skipped unreadable Sent Items entry: $($_.Exception.Message)"}
        }
        $now=Get-Date
        $due=$track|?{!$_.Replied -and !$_.Reminded -and (vendorSelected $_.Vendor) -and ([datetime]$_.Sent).AddHours(24) -lt $now}
        log "Reminders due: $(@($due).Count)"
        terminalProgress "Send reminders" "Preparing reminder emails..." 0 @($due).Count
        $remindNo=0
        foreach($request in $due){
            $remindNo++
            terminalProgress "Send reminders" ("Processing {0}/{1}: {2}" -f $remindNo,@($due).Count,$request.Vendor) ($remindNo-1) @($due).Count
            $mail=$null;$mailAttachment=$null
            try{
                $mail=newReminderReply $items $request.Subject $cut
                if(!$mail){Write-Warning "Skipped reminder: original Sent Items mail not found for Subject: $($request.Subject)";continue}
                $mail.BodyFormat=2
                $mail.HTMLBody='<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;"><b>Dear Supplier,</b><br><br>Please help reply the MR file today.<br><br>Thanks.</div><br>'+$mail.HTMLBody
                if(Test-Path -LiteralPath $request.Attachment){$mailAttachment=$mail.Attachments.Add($request.Attachment)}
                if(!$Display){$mail.Send();$request.Reminded=Get-Date -Format s}else{$mail.Display()}
                log "$(if($Display){'Displayed reminder'}else{'Sent reminder'}): $($request.Vendor)" "SUCCESS"
            }catch{Write-Warning "Failed reminder $($request.Vendor): $($_.Exception.Message)"}
            finally{releaseCom $mailAttachment;releaseCom $mail}
        }
        terminalProgressDone "Send reminders"
        $track|Export-Csv -LiteralPath $trackingFile -NoTypeInformation -Encoding UTF8
        terminalSection "Result"
        log ("Duration: {0}" -f (terminalDuration))
        log "Remind complete" "SUCCESS"
    }finally{
        releaseCom $items;releaseCom $sent;releaseCom $namespace;releaseCom $ol
    }
}

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$sentFile=Join-Path $OutDir "mr_sent.csv"
terminalBanner $Mode
log "Working folder: $Root"
if($Mode -eq "remind"){invokeRemind $sentFile;return}
function readTemplateHeaders($path){
    if(!(Test-Path -LiteralPath $path)){return @()}
    $xl=$null;$wb=$null;$ws=$null;$used=$null
    try{
        $xl=New-Object -ComObject Excel.Application;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=3
        $wb=$xl.Workbooks.Open($path)
        try{$ws=$wb.Worksheets.Item("MR")}catch{$ws=$wb.Worksheets.Item(1)}
        $used=$ws.UsedRange;$lastCol=$used.Column+$used.Columns.Count-1;$headers=@()
        for($c=1;$c -le $lastCol;$c++){$h=([string]$ws.Cells.Item(1,$c).Text).Trim();if($h){$headers+=$h}}
        return $headers
    }finally{
        if($wb){try{$wb.Close($false)|Out-Null}catch{}}
        if($xl){try{$xl.Quit()|Out-Null}catch{}}
        releaseCom $used;releaseCom $ws;releaseCom $wb;releaseCom $xl
    }
}
function uniqueHeaders($headers){
    $seen=@{};$out=@()
    foreach($h in $headers){
        if($seen.ContainsKey($h)){$seen[$h]++;$out+=("{0}__DUP{1}" -f $h,$seen[$h])}
        else{$seen[$h]=1;$out+=$h}
    }
    return $out
}
$defaultMasterCols=@("Item Seq","Change Type","Batch No","Phase Code","Item No","Description","Customer Item","Order Remark","Original Schedule Ship Date","Request Qty","Request Date","Material Ready Date","Part No.","Shortage Qty","Usage","Material Need By Date","MC Remark","MC Remark with Detail Delivery","ASCP LT","Buyer","Part Desc","Vendor Name","Release Date","Bu Category","Reply Qty","SUPPLIER NEED CHECK")
$traceCols=@("_SourceKey","_SourceFile","_SourceSheet","_SourceRow")
$templateDisplayCols=@(readTemplateHeaders $Template)
if(!$templateDisplayCols -or $templateDisplayCols.Count -eq 0){$templateDisplayCols=$defaultMasterCols}
$templateInternalCols=@(uniqueHeaders $templateDisplayCols)
$masterCols=@($templateInternalCols+$traceCols)
$masterDisplayCols=@($templateDisplayCols+$traceCols)
$displayByInternal=@{};for($i=0;$i -lt $masterCols.Count;$i++){$displayByInternal[$masterCols[$i]]=$masterDisplayCols[$i]}
log "Master template columns: $($templateDisplayCols.Count) from $(if(Test-Path -LiteralPath $Template){$Template}else{'fallback schema'})"
$masterAliases=@{
    "Part No."=@("Part No.","Part No","PartNo")
    "Material Need By Date"=@("Material Need By Date","Material need by date","Material Need Date")
    "Vendor Name"=@("Vendor Name","Vendor name","VendorName")
    "ASCP LT"=@("ASCP LT","Part lead time","Part Lead Time","Lead Time")
    "SUPPLIER NEED CHECK"=@("SUPPLIER NEED CHECK","MC double check","MC Double Check")
}
function displayHeader($h){if($displayByInternal.ContainsKey($h)){$displayByInternal[$h]}else{$h}}
function supplierHeader($h){
    $d=displayHeader $h
    $n=norm $d
    if($n -in @("MC DOUBLE CHECK","SUPPLIER NEED CHECK")){return "ETA Vendor can Supply"}
    if($n -eq "REMARK"){return "IF CAN NOT"}
    return $d
}
function headerIndex($headers,$names){
    foreach($name in $names){
        for($i=0;$i -lt $headers.Count;$i++){if((norm $headers[$i]) -eq (norm $name)){return $i+1}}
    }
    return 0
}
function columnIndexes($headers,$names,$required=$true){
    $indexes=@{}
    foreach($name in $names){
        $index=headerIndex $headers @($name)
        if(!$index -and $required){throw "Missing column: $name"}
        $indexes[$name]=$index
    }
    return $indexes
}
$masterTraceIndexes=columnIndexes $masterCols $traceCols
log ("Trace column indexes: " + (($traceCols|%{"$_=$($masterTraceIndexes[$_])"}) -join ", "))
function addGuidelineSheet($wb,$afterSheet){
    $guide=$null
    try{
        $guide=$wb.Worksheets.Add([Type]::Missing,$afterSheet)
        $guide.Name="Guideline"
        $guide.Range("A1:B1").Merge()|Out-Null
        $guide.Cells.Item(1,1).Value2="Supplier Feedback Guideline"
        $guide.Cells.Item(3,1).Value2="Field";$guide.Cells.Item(3,2).Value2="What to fill"
        $guide.Cells.Item(4,1).Value2="Reply Qty";$guide.Cells.Item(4,2).Value2="Confirm the quantity you can supply."
        $guide.Cells.Item(5,1).Value2="ETA Vendor can Supply";$guide.Cells.Item(5,2).Value2="Fill the ETA/date you can supply for pull-in or add-in demand in day/month/year format (for example 22/2/2026 or 22/02/2026)."
        $guide.Cells.Item(6,1).Value2="IF CAN NOT";$guide.Cells.Item(6,2).Value2="If your provided ETA cannot meet our requested date, please give the detailed reason and recovery/action option."
        $guide.Cells.Item(7,1).Value2="Required details";$guide.Cells.Item(7,2).Value2="Capacity constraints; Material constraints (please specify the material); Manufacturing constraints; Available options (spot buy, truck, air shipment, etc.) and related cost; Any actions or options you can take to meet our requested ETA."
        $guide.Cells.Item(8,1).Value2="Feedback deadline";$guide.Cells.Item(8,2).Value2="Please feedback within 24 hours."
        $guide.Range("A1:B1").Interior.Color=6299648;$guide.Range("A1:B1").Font.Color=16777215;$guide.Range("A1:B1").Font.Bold=$true;$guide.Range("A1:B1").Font.Size=14
        $guide.Range("A3:B3").Interior.Color=15189684;$guide.Range("A3:B3").Font.Bold=$true
        $guide.Range("A3:B8").Borders.LineStyle=1;$guide.Range("A3:B8").Borders.Color=14277081
        $guide.Range("A4:A8").Font.Bold=$true;$guide.Range("B4:B8").WrapText=$true
        $guide.Columns.Item(1).ColumnWidth=24;$guide.Columns.Item(2).ColumnWidth=110
        $guide.Rows.Item(7).RowHeight=54
    }finally{releaseCom $guide}
}
function sourceNames($h){
    $d=displayHeader $h
    if($masterAliases[$d]){return $masterAliases[$d]}
    if($masterAliases[$h]){return $masterAliases[$h]}
    return @($d,$h)|Select-Object -Unique
}
function headerInfo($values,$firstRow,$firstCol,$cols,$lastRow){
    $bestRow=0;$bestScore=0;$bestMap=@{}
    for($r=1;$r -le [Math]::Min(100,$lastRow);$r++){
        $relativeRow=$r-$firstRow+1
        $map=@{};for($c=$firstCol;$c -lt $firstCol+$cols;$c++){$h=norm (cellValue $values $relativeRow ($c-$firstCol+1));if($h -and !$map.ContainsKey($h)){$map[$h]=$c}}
        $score=0
        foreach($h in $masterCols){if($traceCols -contains $h){continue};foreach($n in (sourceNames $h)){if($map.ContainsKey((norm $n))){$score++;break}}}
        if($score -gt $bestScore){$bestRow=$r;$bestScore=$score;$bestMap=$map}
    }
    [pscustomobject]@{Row=$bestRow;Score=$bestScore;Map=$bestMap}
}
function cellValue($values,$row,$col){
    try{$v=$values[$row,$col]}catch{return ""}
    if($null -eq $v){return ""}
    return [string]$v
}
function isDateHeader($h){(displayHeader $h) -match '(?i)(date|system update on)'}
function formatOaDate($n){
    $dt=[datetime]::FromOADate([double]$n)
    if([math]::Abs(([double]$n)-[math]::Floor([double]$n)) -gt 0.00001){return $dt.ToString("dd/MM/yyyy HH:mm")}
    return $dt.ToString("dd/MM/yyyy")
}
function dateTextValue($value){
    $s=([string]$value).Trim()
    if(!$s){return ""}
    $n=0.0
    if([double]::TryParse($s,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$n) -and $n -gt 20000 -and $n -lt 80000){return formatOaDate $n}
    return $s
}
function etaDateValue($value){
    $s=([string]$value).Trim()
    if(!$s){return $null}
    $n=0.0
    if([double]::TryParse($s,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$n) -and $n -gt 20000 -and $n -lt 80000){return [datetime]::FromOADate($n)}
    [string[]]$formats=@("d/M/yyyy","dd/M/yyyy","d/MM/yyyy","dd/MM/yyyy","d-MMM-yyyy","dd-MMM-yyyy")
    $dt=[datetime]::MinValue
    if([datetime]::TryParseExact($s,$formats,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$dt)){return $dt}
    return $null
}
function validEtaDate($value){return $null -ne (etaDateValue $value)}
function normalizeEtaDate($value){
    $dt=etaDateValue $value
    if($null -eq $dt){return ""}
    return $dt.ToString("dd-MMM-yyyy",[Globalization.CultureInfo]::InvariantCulture)
}
function excelEtaValue($value){
    $dt=etaDateValue $value
    if($null -ne $dt){return $dt.ToOADate()}
    return $value
}
function configureEtaInput($ws,$headers,$lastRow){
    $etaCol=headerIndex $headers @("ETA Vendor can Supply")
    if(!$etaCol -or $lastRow -lt 2){return}
    $etaRange=$null
    try{
        $etaRange=$ws.Range($ws.Cells.Item(2,$etaCol),$ws.Cells.Item($lastRow,$etaCol))
        $etaRange.NumberFormat="dd-mmm-yyyy"
        $etaCell=$ws.Cells.Item(2,$etaCol).Address($false,$false)
        $formula=('=OR({0}="",AND(ISNUMBER({0}),{0}>=DATE(2000,1,1),{0}<=DATE(2100,12,31)))' -f $etaCell)
        $etaRange.Validation.Delete()
        $etaRange.Validation.Add(7,1,1,$formula)
        $etaRange.Validation.IgnoreBlank=$true
        $etaRange.Validation.ShowInput=$true
        $etaRange.Validation.InputTitle="ETA Vendor can Supply"
        $etaRange.Validation.InputMessage="Enter a date, e.g. 29-Jul-2026."
        $etaRange.Validation.ShowError=$true
        $etaRange.Validation.ErrorTitle="Invalid ETA"
        $etaRange.Validation.ErrorMessage="Enter a valid date, e.g. 29-Jul-2026."
    }finally{releaseCom $etaRange}
}
function invalidEtaRows($values,$etaCol,$lastRow){
    if(!$etaCol -or $lastRow -lt 2){return}
    $invalid=@()
    for($row=2;$row -le $lastRow;$row++){
        $eta=([string](cellValue $values $row $etaCol)).Trim()
        if($eta -and !(validEtaDate $eta)){$invalid+=$row}
    }
    return $invalid
}
function replyMatchKey($vendorName,$partNumber,$quantity,$needDate){
    return @((norm $vendorName),(norm $partNumber),(norm $quantity),(norm (materialNeedDateText $needDate))) -join [char]31
}
function isMaterialNeedDateHeader($h){(norm (displayHeader $h)) -eq "MATERIAL NEED BY DATE"}
function materialNeedDateText($value){
    $s=([string]$value).Trim()
    if(!$s){return ""}
    $n=0.0;$dt=[datetime]::MinValue
    if([double]::TryParse($s,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$n) -and $n -gt 20000 -and $n -lt 80000){$dt=[datetime]::FromOADate($n)}
    else{
        [string[]]$formats=@("d/M/yyyy","dd/M/yyyy","d/MM/yyyy","dd/MM/yyyy","d-MMM-yy","dd-MMM-yy","d-MMM-yyyy","dd-MMM-yyyy","yyyy/M/d","yyyy-MM-dd")
        if(![datetime]::TryParseExact($s,$formats,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$dt)){return $s}
    }
    return $dt.ToString("d-MMM-yy",[Globalization.CultureInfo]::InvariantCulture)
}
function cellMasterValue($ws,$values,$row,$absRow,$col,$absCol,$header){
    if(isMaterialNeedDateHeader $header){return materialNeedDateText (cellValue $values $row $col)}
    if(isDateHeader $header){
        $raw=cellValue $values $row $col
        $rawDate=dateTextValue $raw
        if($rawDate -and $rawDate -ne ([string]$raw).Trim()){return $rawDate}
        $txt=([string]$ws.Cells.Item($absRow,$absCol).Text).Trim()
        if($txt -and $txt -notmatch '^#+$'){return dateTextValue $txt}
        return dateTextValue $raw
    }
    return cellValue $values $row $col
}
function dateNumberFormat($h){
    if(isMaterialNeedDateHeader $h){return "d-mmm-yy"}
    if((displayHeader $h) -match '(?i)(adjust date|release date|system update on)'){return "dd/mm/yyyy hh:mm"}
    return "dd/mm/yyyy"
}
function excelDateValue($value){
    $s=([string]$value).Trim()
    if(!$s){return ""}
    $n=0.0
    if([double]::TryParse($s,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$n) -and $n -gt 20000 -and $n -lt 80000){return $n}
    $dt=[datetime]::MinValue
    foreach($culture in @([Globalization.CultureInfo]::CurrentCulture,[Globalization.CultureInfo]::GetCultureInfo("vi-VN"),[Globalization.CultureInfo]::GetCultureInfo("en-GB"),[Globalization.CultureInfo]::GetCultureInfo("en-US"),[Globalization.CultureInfo]::InvariantCulture)){
        if([datetime]::TryParse($s,$culture,[Globalization.DateTimeStyles]::AssumeLocal,[ref]$dt)){return $dt.ToOADate()}
    }
    return $s
}
function mappedSourceCols($map,$headers){
    $source=@{}
    foreach($h in $headers){
        if($h -like '_Source*'){continue}
        $col=$null
        foreach($n in (sourceNames $h)){$k=norm $n;if($map.ContainsKey($k)){$col=$map[$k];break}}
        $source[$h]=$col
    }
    return $source
}
function exportMasterXlsx($txtPath,$xlsx,$templatePath,$headers,$displayHeaders,$delimiter="`t"){
    $xl=$null;$wb=$null;$ws=$null;$used=$null;$clearRange=$null;$dataRange=$null;$headerRange=$null
    try{
        $xl=New-Object -ComObject Excel.Application
        $xl.DisplayAlerts=$false;$xl.ScreenUpdating=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=3
        try{$xl.Calculation=-4135}catch{}
        $fromTemplate=Test-Path -LiteralPath $templatePath
        if($fromTemplate){Copy-Item -LiteralPath $templatePath -Destination $xlsx -Force;$wb=$xl.Workbooks.Open($xlsx)}else{$wb=$xl.Workbooks.Add()}
        try{$ws=$wb.Worksheets.Item("MR")}catch{$ws=$wb.Worksheets.Item(1);$ws.Name="MR"}
        $used=$ws.UsedRange;$lastUsedRow=$used.Row+$used.Rows.Count-1;$lastUsedCol=[Math]::Max($headers.Count,$used.Column+$used.Columns.Count-1)
        $clearRange=$ws.Range($ws.Cells.Item(1,1),$ws.Cells.Item($lastUsedRow,$lastUsedCol));$clearRange.ClearContents()|Out-Null
        $list=@(Import-Csv -LiteralPath $txtPath -Delimiter $delimiter)
        if($list.Count -gt 0){
            $dataHeaders=@($list[0].PSObject.Properties.Name)
            $arr=New-Object 'object[,]' $list.Count,$dataHeaders.Count
            for($r=0;$r -lt $list.Count;$r++){for($c=0;$c -lt $dataHeaders.Count;$c++){$v=[string]$list[$r].($dataHeaders[$c]);$arr[$r,$c]=if(isDateHeader $dataHeaders[$c]){excelDateValue $v}else{$v}}}
            $dataRange=$ws.Range($ws.Cells.Item(2,1),$ws.Cells.Item($list.Count+1,$dataHeaders.Count));$dataRange.Value2=$arr
        }
        for($c=0;$c -lt $displayHeaders.Count;$c++){$ws.Cells.Item(1,$c+1).Value2=$displayHeaders[$c]}
        for($c=0;$c -lt $displayHeaders.Count;$c++){if(isDateHeader $headers[$c]){$ws.Columns.Item($c+1).NumberFormat=dateNumberFormat $headers[$c]}}
        if($ws.AutoFilterMode){$ws.AutoFilterMode=$false}
        $headerRange=$ws.Range($ws.Cells.Item(1,1),$ws.Cells.Item(1,$headers.Count));$headerRange.AutoFilter()|Out-Null;$headerRange.Font.Bold=$true
        $ws.Columns.AutoFit()|Out-Null
        if($fromTemplate){$wb.Save()}else{$wb.SaveAs($xlsx,51)}
        log "Master xlsx staged: $xlsx"
    }finally{
        if($wb){try{$wb.Close($false)|Out-Null}catch{}}
        if($xl){try{$xl.Quit()|Out-Null}catch{}}
        releaseCom $headerRange;releaseCom $dataRange;releaseCom $clearRange;releaseCom $used;releaseCom $ws;releaseCom $wb;releaseCom $xl
    }
}
function writeMasterOutputs($rows,$txtPath,$xlsxPath,$templatePath,$headers,$displayHeaders,$delimiter){
    $folder=Split-Path -Parent $txtPath;$stamp=[guid]::NewGuid().ToString("N")
    $stagingTxt=Join-Path $folder (".{0}.{1}.txt" -f [IO.Path]::GetFileNameWithoutExtension($txtPath),$stamp)
    $stagingXlsx=Join-Path $folder (".{0}.{1}.xlsx" -f [IO.Path]::GetFileNameWithoutExtension($xlsxPath),$stamp)
    try{
        if($delimiter -eq "`t"){exportTsv $rows $stagingTxt}else{$rows|Export-Csv -LiteralPath $stagingTxt -NoTypeInformation -Encoding UTF8}
        if(!(Test-Path -LiteralPath $stagingTxt)){throw "Failed to stage master text output: $stagingTxt"}
        exportMasterXlsx $stagingTxt $stagingXlsx $templatePath $headers $displayHeaders $delimiter
        if(!(Test-Path -LiteralPath $stagingXlsx) -or (Get-Item -LiteralPath $stagingXlsx).Length -le 0){throw "Failed to stage master Excel output: $stagingXlsx"}
        assertMasterFileWritable $txtPath;assertMasterFileWritable $xlsxPath
        Move-Item -LiteralPath $stagingTxt -Destination $txtPath -Force
        Move-Item -LiteralPath $stagingXlsx -Destination $xlsxPath -Force
        if(!(Test-Path -LiteralPath $txtPath) -or !(Test-Path -LiteralPath $xlsxPath)){throw "Master outputs were not published"}
        log "Master outputs published: $txtPath and $xlsxPath" "SUCCESS"
    }finally{
        foreach($staging in @($stagingTxt,$stagingXlsx)){if(Test-Path -LiteralPath $staging){Remove-Item -LiteralPath $staging -Force -ErrorAction SilentlyContinue}}
    }
}
function buildMaster($folder,$masterFile){
    $rows=[System.Collections.ArrayList]::new();$xl=New-Object -ComObject Excel.Application
    $xl.DisplayAlerts=$false;$xl.ScreenUpdating=$false;$xl.EnableEvents=$false;try{$xl.AutomationSecurity=3}catch{};try{$xl.Calculation=-4135}catch{}
    try{
        $files=@(Get-ChildItem -LiteralPath $folder -File|?{$_.Name -notlike '~$*' -and $_.Name -notlike 'MR_Master_Input*' -and $_.Extension -match '^\.xls'})
        log "Input folder: $folder"
        log "Excel files to read: $($files.Count)"
        terminalProgress "Prepare input" "Waiting to read Excel files..." 0 $files.Count
        $fileNo=0
        foreach($file in $files){
            $fileNo++;$before=$rows.Count
            terminalProgress "Prepare input" ("Reading {0}/{1}: {2}" -f $fileNo,$files.Count,$file.Name) ($fileNo-1) $files.Count
            log "Reading $fileNo/$($files.Count): $($file.Name)" "STEP"
            $wb=$null;$ws=$null;$used=$null
            log "Opening workbook: $($file.Name)"
            try{$wb=$xl.Workbooks.Open($file.FullName)}catch{Write-Warning "Cannot open input workbook: $($file.Name) - $($_.Exception.Message)";continue}
            log "Opened workbook: $($file.Name)"
            try{
                try{$ws=$wb.Worksheets.Item("MR")}catch{Write-Warning "Sheet MR not found: $($file.Name)";continue}
                $used=$ws.UsedRange;$firstRow=$used.Row;$firstCol=$used.Column;$cols=$used.Columns.Count;$lastRow=$firstRow+$used.Rows.Count-1;$values=$used.Value2
                log "Sheet MR size: $lastRow rows x $cols columns"
                $hi=headerInfo $values $firstRow $firstCol $cols $lastRow;$map=$hi.Map
                if($hi.Score -lt 3){Write-Warning "Cannot detect header row in MR sheet: $($file.Name)";continue}
                log "Header row: $($hi.Row) ($($hi.Score) matched columns)"
                $sourceCols=mappedSourceCols $map $masterCols
                $vendorSourceCol=$sourceCols["Vendor Name"];$partSourceCol=$sourceCols["Part No."]
                $lastDataRow=$lastRow
                if($vendorSourceCol -or $partSourceCol){
                    $lastDataRow=$hi.Row
                    for($scan=$lastRow;$scan -gt $hi.Row;$scan--){
                        $scanRr=$scan-$firstRow+1
                        $vendorValue=if($vendorSourceCol){cellValue $values $scanRr ($vendorSourceCol-$firstCol+1)}else{""}
                        $partValue=if($partSourceCol){cellValue $values $scanRr ($partSourceCol-$firstCol+1)}else{""}
                        if(![string]::IsNullOrWhiteSpace($vendorValue) -or ![string]::IsNullOrWhiteSpace($partValue)){$lastDataRow=$scan;break}
                    }
                    log "Data rows to scan: $($lastDataRow-$hi.Row) of $($lastRow-$hi.Row) used rows"
                }
                for($r=$hi.Row+1;$r -le $lastDataRow;$r++){
                    if((($r-$hi.Row) % 500) -eq 0){log "Reading rows: $($r-$hi.Row)/$($lastDataRow-$hi.Row) from $($file.Name)" "STEP"}
                    $rr=$r-$firstRow+1
                    $vendorValue=if($vendorSourceCol){cellValue $values $rr ($vendorSourceCol-$firstCol+1)}else{""}
                    $partValue=if($partSourceCol){cellValue $values $rr ($partSourceCol-$firstCol+1)}else{""}
                    if([string]::IsNullOrWhiteSpace($vendorValue) -and [string]::IsNullOrWhiteSpace($partValue)){continue}
                    $o=[ordered]@{}
                    foreach($h in $masterCols){
                        if($traceCols -contains $h){$o[$h]=""}else{$col=$sourceCols[$h];$o[$h]=if($col){cellMasterValue $ws $values $rr $r ($col-$firstCol+1) $col $h}else{""}}
                    }
                    $o["_SourceFile"]=$file.Name;$o["_SourceSheet"]="MR";$o["_SourceRow"]=$r;$o["_SourceKey"]="$($file.Name)|MR|$r"
                    [void]$rows.Add([pscustomobject]$o)
                }
            }finally{
                if($wb){try{$wb.Close($false)|Out-Null}catch{}}
                releaseCom $used;releaseCom $ws;releaseCom $wb
            }
            log "Rows added from $($file.Name): $($rows.Count-$before)"
            terminalProgress "Prepare input" ("Completed {0}/{1}: {2}" -f $fileNo,$files.Count,$file.Name) $fileNo $files.Count
        }
    }finally{if($xl){try{$xl.Quit()|Out-Null}catch{}};releaseCom $xl}
    if($rows.Count -eq 0){throw "No rows found in sheet MR under $folder"}
    log "Master input created: $masterFile ($($rows.Count) rows)"
    writeMasterOutputs $rows $masterFile ([IO.Path]::ChangeExtension($masterFile,".xlsx")) $Template $masterCols $masterDisplayCols "`t"
    terminalProgressDone "Prepare input"
}

$latestFolder=latestInputFolder $InputRoot
if($latestFolder){
    $masterFile=Join-Path $latestFolder "MR_Master_Input.txt"
    if($Mode -eq "prepare" -or !(Test-Path -LiteralPath $masterFile)){buildMaster $latestFolder $masterFile}
    $InputFile=$masterFile
}
log "Input file: $InputFile"
$delim=if((Get-Content -LiteralPath $InputFile -TotalCount 1) -like "*`t*"){"`t"}else{","}
$data=Import-Csv -LiteralPath $InputFile -Delimiter $delim
if(!$data){throw "No rows in $InputFile"}
log "Input rows: $(@($data).Count)"
$vc=col $data[0] @("Vendor name","Vendor Name","VendorName")
$pc=col $data[0] @("Part No.","Part No","PartNo")
$qc=col $data[0] @("Request Qty","New Request Qty","Qty","Shortage Qty")
$dc=col $data[0] @("Material Need By Date","Material need by date","Material Ready Date","MR Date","Ready Date")
$rq=col $data[0] @("Reply Qty") $false;if(!$rq){$rq="Reply Qty";$data|%{$_|Add-Member NoteProperty $rq "" -Force}}
$mc=col $data[0] @("ETA Vendor can Supply","ETA Vendor can Suppy","SUPPLIER NEED CHECK","MC double check","MC Double Check") $false
if(!$mc){$mc="SUPPLIER NEED CHECK";$data|%{$_|Add-Member NoteProperty $mc "" -Force}}
$ifCannot=col $data[0] @("IF CAN NOT","Remark") $false
if(!$ifCannot){$ifCannot="Remark";$data|%{$_|Add-Member NoteProperty $ifCannot "" -Force}}
$sk=col $data[0] @("_SourceKey") $false

if($Mode -eq "prepare"){
    terminalSection "Result"
    log ("Duration: {0}" -f (terminalDuration))
    log "Prepare complete. Master files are ready under: $latestFolder" "SUCCESS"
    return
}

if($Mode -eq "scan"){
    assertMasterFileWritable $InputFile
    assertMasterFileWritable ([IO.Path]::ChangeExtension($InputFile,".xlsx"))
    log "Master outputs are writable" "SUCCESS"
}

if($Mode -eq "send"){
    $sup=@(Import-Csv -LiteralPath $Supplier)
    if(!$sup.Count){throw "No supplier rows in $Supplier"}
    $sv=col $sup[0] @("Vendor name","Vendor Name","VendorName")
    $to=col $sup[0] @("Email To","EmailTo","To")
    $cc=col $sup[0] @("Email CC","EmailCC","CC") $false
    $mcContact=col $sup[0] @("MC","MC in charge","MC In Charge","MC Email","MC email") $false
    $kw=col $sup[0] @("Keyword","Keywords","Alias","Aliases") $false
    $byVendor=@{};$supplierAliases=@{}
    foreach($srow in $sup){
        addSupplierKey $byVendor $srow.$sv $srow
        if($kw){foreach($alias in ([string]$srow.$kw -split ';')){addSupplierKey $supplierAliases $alias $srow}}
    }
    $ol=$null;$xl=$null;$sentCount=0;$displayedCount=0;$failedCount=0
    $outCols=@($data[0].PSObject.Properties.Name)
    $outHeaders=@($outCols|%{supplierHeader $_})
    $visibleSupplierCols=@(
        "Change Type","Item No","Order Remark","Request Qty","Part No.","Shortage Qty",
        "Material need by date","ETA Vendor can Supply","MC Remarks for MR","Part lead time",
        "Usage","Part Desc","Vendor Name","_SourceFile","_SourceSheet","_SourceRow","IF CAN NOT"
    )
    $visibleSupplierIndexes=columnIndexes $outHeaders $visibleSupplierCols
    $emailTableColumns=@(
        [pscustomobject]@{Header="OFIS Line ID";Names=@("OFIS Line ID")},
        [pscustomobject]@{Header="Order Remark";Names=@("Order Remark")},
        [pscustomobject]@{Header="Request Qty";Names=@("Request Qty","New Request Qty","Qty")},
        [pscustomobject]@{Header="Part No.";Names=@("Part No.","Part No","PartNo")},
        [pscustomobject]@{Header="Material need by date";Names=@("Material Need By Date","Material need by date","Material Ready Date","MR Date","Ready Date")},
        [pscustomobject]@{Header="ETA Vendor can Supply";Names=@("ETA Vendor can Supply","ETA Vendor can Suppy","SUPPLIER NEED CHECK","MC double check","MC Double Check")},
        [pscustomobject]@{Header="Usage";Names=@("Usage")},
        [pscustomobject]@{Header="Part Lead Time";Names=@("Part Lead Time","ASCP LT","Lead Time")},
        [pscustomobject]@{Header="IF CAN NOT";Names=@("IF CAN NOT","Remark")},
        [pscustomobject]@{Header="_SourceKey";Names=@("_SourceKey")}
    )
    $aliases=@{
        "Request Qty"=@("Request Qty","New Request Qty","Qty")
        "Part No."=@("Part No.","Part No","PartNo")
        "Material Need By Date"=@("Material Need By Date","Material need by date","Material Ready Date","MR Date","Ready Date")
        "MC double check"=@("ETA Vendor can Supply","ETA Vendor can Suppy","SUPPLIER NEED CHECK","MC double check","MC Double Check")
        "SUPPLIER NEED CHECK"=@("ETA Vendor can Supply","ETA Vendor can Suppy","SUPPLIER NEED CHECK","MC double check","MC Double Check")
        "Remark"=@("IF CAN NOT","Remark")
    }
    $etaInstructionColumn=excelColumnName (headerIndex $outHeaders @("ETA Vendor can Supply"))
    $cannotInstructionColumn=excelColumnName (headerIndex $outHeaders @("IF CAN NOT"))
    try{
        $ol=New-Object -ComObject Outlook.Application
        $xl=New-Object -ComObject Excel.Application
        $xl.DisplayAlerts=$false;$xl.ScreenUpdating=$false;$xl.EnableEvents=$false
        try{$xl.AutomationSecurity=3}catch{}
        try{$xl.Calculation=-4135}catch{}
        $groups=@($data|?{vendorSelected $_.$vc}|Group-Object $vc)
        log "Suppliers to process: $($groups.Count)"
        terminalProgress "Send MR" "Preparing supplier emails..." 0 $groups.Count
        $sendNo=0
        foreach($group in $groups){
            $sendNo++
            $vendorName=$group.Name.Trim();if(!$vendorName){continue}
            terminalProgress "Send MR" ("Processing {0}/{1}: {2}" -f $sendNo,$groups.Count,$vendorName) ($sendNo-1) $groups.Count
            log "Sending $sendNo/$($groups.Count): $vendorName ($(@($group.Group).Count) rows)" "STEP"
            $s=findSupplier $vendorName $byVendor $supplierAliases
            if(!$s){Write-Warning "Supplier config not found or ambiguous: $vendorName";continue}
            if([string]::IsNullOrWhiteSpace([string]$s.$to)){Write-Warning "Email To is empty: $vendorName";continue}
            $toValue=([string]$s.$to).Trim()
            if(!(validEmailList $toValue)){Write-Warning "Invalid Email To, skipped: $vendorName";continue}

            $wb=$null;$ws=$null;$range=$null;$header=$null;$materialNeedRange=$null;$mail=$null;$recipients=$null;$mailAttachment=$null;$mcRecipients=@()
            try{
                $stamp=Get-Date -Format "yyyyMMdd_HHmmss"
                $safe=$vendorName -replace '[\\/:*?"<>|]','_'
                $xlsx=Join-Path (Resolve-Path $OutDir) "MR_${safe}_$stamp.xlsx"
                $wb=$xl.Workbooks.Add();$ws=$wb.Worksheets.Item(1);$ws.Name="MR";$ws.Columns.NumberFormat="@"
                $rowCount=@($group.Group).Count
                $arr=New-Object 'object[,]' ($rowCount+1),$outCols.Count
                for($columnIndex=0;$columnIndex -lt $outCols.Count;$columnIndex++){$arr[0,$columnIndex]=$outHeaders[$columnIndex]}
                for($rowIndex=0;$rowIndex -lt $rowCount;$rowIndex++){
                    $sourceRow=$group.Group[$rowIndex]
                    for($columnIndex=0;$columnIndex -lt $outCols.Count;$columnIndex++){
                        $columnName=$outCols[$columnIndex]
                        $names=if($aliases[$columnName]){$aliases[$columnName]}else{@($columnName)}
                        $value=val $sourceRow $names
                        if((norm (supplierHeader $columnName)) -eq "ETA VENDOR CAN SUPPLY"){$value=excelEtaValue $value}
                        elseif((norm (supplierHeader $columnName)) -eq "MATERIAL NEED BY DATE"){$value=materialNeedDateText $value;$value=excelDateValue $value}
                        $arr[($rowIndex+1),$columnIndex]=$value
                    }
                }
                $lastOutputRow=$rowCount+1
                $range=$ws.Range($ws.Cells.Item(1,1),$ws.Cells.Item($lastOutputRow,$outCols.Count))
                $range.Value2=$arr
                $range.Borders.LineStyle=1;$range.Borders.Color=14277081
                $materialNeedCol=headerIndex $outHeaders @("Material need by date")
                if($materialNeedCol -and $lastOutputRow -gt 1){$materialNeedRange=$ws.Range($ws.Cells.Item(2,$materialNeedCol),$ws.Cells.Item($lastOutputRow,$materialNeedCol));$materialNeedRange.NumberFormat="d-mmm-yy"}
                $header=$ws.Range($ws.Cells.Item(1,1),$ws.Cells.Item(1,$outCols.Count));$header.Font.Bold=$true;$header.Font.Color=16777215;$header.Interior.Color=6299648;$header.AutoFilter()|Out-Null
                $importantCols=@("Part No.","Request Qty","Material need by date","ETA Vendor can Supply","IF CAN NOT")
                $importantCols|%{$ic=headerIndex $outHeaders @($_);if($ic -gt 0){$ws.Cells.Item(1,$ic).Interior.Color=255;$ws.Cells.Item(1,$ic).Font.Color=16777215;if($lastOutputRow -gt 1){$ws.Range($ws.Cells.Item(2,$ic),$ws.Cells.Item($lastOutputRow,$ic)).Interior.Color=13434879}}}
                addGuidelineSheet $wb $ws
                $ws.Activate()|Out-Null
                for($columnIndex=1;$columnIndex -le $outCols.Count;$columnIndex++){$ws.Columns.Item($columnIndex).ColumnWidth=14}
                $importantCols|%{$ic=headerIndex $outHeaders @($_);if($ic -gt 0){$ws.Columns.Item($ic).ColumnWidth=24}}
                configureEtaInput $ws $outHeaders $lastOutputRow
                for($columnIndex=1;$columnIndex -le $outCols.Count;$columnIndex++){
                    if($visibleSupplierIndexes.Values -notcontains $columnIndex){$ws.Columns.Item($columnIndex).Hidden=$true}
                }
                $ws.Rows.Item(1).RowHeight=24;$ws.Application.ActiveWindow.SplitRow=1;$ws.Application.ActiveWindow.FreezePanes=$true
                $wb.SaveAs($xlsx,51);$wb.Close($false)
                releaseCom $header;$header=$null;releaseCom $range;$range=$null;releaseCom $ws;$ws=$null;releaseCom $wb;$wb=$null
                log "Attachment created: $xlsx" "SUCCESS"

                $subject="MR_REQUEST|$vendorName|$stamp"
                $emailTable=if($HideEmailTable){""}else{supplierEmailTable -rows @($group.Group) -columns $emailTableColumns}
                $mcInCharge=if($mcContact){([string]$s.$mcContact).Trim()}else{""}
                if($mcInCharge -in @("TRUE","FALSE")){$mcInCharge=""}
                $mcAddresses=@(([string]$mcInCharge -split ';')|%{$_.Trim()}|?{$_}|Select-Object -Unique)
                if($mcAddresses.Count -and !(validEmailList ($mcAddresses -join ';'))){Write-Warning "Invalid MC email, skipped: $vendorName";continue}
                $ccList=@();if($cc){$ccList+=@(([string]$s.$cc -split ';')|%{$_.Trim()}|?{$_ -and $mcAddresses -notcontains $_})}
                $ccValue=(@($ccList|Select-Object -Unique) -join ";")
                if($ccValue -and !(validEmailList $ccValue)){Write-Warning "Invalid Email CC, skipped: $vendorName";continue}

                $mail=$ol.CreateItem(0);$mail.To=$toValue;$mail.CC=$ccValue
                foreach($mcAddress in $mcAddresses){$mcRecipient=$mail.Recipients.Add($mcAddress);$mcRecipient.Type=2;$mcRecipients+=@($mcRecipient)}
                $recipients=$mail.Recipients
                if(!$recipients.ResolveAll()){Write-Warning "Unresolved Outlook recipient, skipped: $vendorName";continue}
                $mcTag=""
                if($mcRecipients.Count){
                    $mcTags=@()
                    for($mcIndex=0;$mcIndex -lt $mcRecipients.Count;$mcIndex++){
                        $mcAddressEntry=$null;$mcExchangeUser=$null
                        try{
                            $mcAddress=$mcAddresses[$mcIndex]
                            $resolvedMcName=[string]$mcRecipients[$mcIndex].Name
                            try{
                                $mcAddressEntry=$mcRecipients[$mcIndex].AddressEntry
                                if($mcAddressEntry){
                                    try{$mcExchangeUser=$mcAddressEntry.GetExchangeUser()}catch{}
                                    if($mcExchangeUser -and $mcExchangeUser.Name){$resolvedMcName=[string]$mcExchangeUser.Name}
                                    elseif($mcAddressEntry.Name){$resolvedMcName=[string]$mcAddressEntry.Name}
                                }
                            }catch{}
                            $mcName=compactMcName $resolvedMcName $mcAddress
                            $mcTags+=("<a href='mailto:{0}' style='color:#0563c1;text-decoration:underline;'>{1}</a>" -f [Net.WebUtility]::HtmlEncode($mcAddress),[Net.WebUtility]::HtmlEncode("@$mcName"))
                            log "MC tag resolved: @$mcName for $vendorName" "SUCCESS"
                        }finally{
                            releaseCom $mcExchangeUser
                            releaseCom $mcAddressEntry
                        }
                    }
                    $mcTag=" "+($mcTags -join ", ")
                }else{
                    log "MC tag skipped: no MC configured for $vendorName" "WARNING"
                }
                $mail.Subject=$subject;$mail.BodyFormat=2;$mail.HTMLBody=@"
<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:1.25;">
<b>Dear Supplier,</b><br>
This is a request to add in / pull in material. Please kindly check the details and provide your feedback with the following<br>
<b>information:</b>
<div style="margin:12px 0;padding:10px 12px;background-color:#fff2cc;border-left:5px solid #f2b632;color:#7f6000;">
<b>Action required:</b> Please download the attached MR file, complete the requested information, and reply to this same email thread with the completed file attached.
</div>
$emailTable
<ol style="margin-top:8px;margin-bottom:8px;">
  <li>Fill in column <b>$etaInstructionColumn</b> with the ETA in <b>DD-MMM-YYYY</b> format (for example, 29-Jul-2026).</li>
  <li>If your provided ETA cannot meet our requested date, please fill in column <b>$cannotInstructionColumn</b> with the detailed reason, such as:
    <ul>
      <li>Capacity constraints</li>
      <li>Material constraints (please specify the material)</li>
      <li>Manufacturing constraints</li>
      <li><b>Available options (spot buy, truck, air shipment, etc.) and the related cost</b></li>
      <li>Any actions or options you can take to meet our requested ETA</li>
    </ul>
  </li>
</ol>
Please provide all details in the designated column.<br>
Your feedback is required within <b>24 hours</b>. We appreciate your prompt cooperation and support.<br><br>
<b>Dear MCs,</b>$mcTag<br>
Please kindly follow up on this request with your suppliers and ensure their feedback is received within the required timeline.<br>
Thank you.
</div>
"@
                $mailAttachment=$mail.Attachments.Add($xlsx)
                $resolvedTo=[string]$mail.To;$resolvedCc=[string]$mail.CC
                if(!$Display){
                    $mail.Send()
                    $record=[pscustomobject]@{Vendor=$vendorName;To=$resolvedTo;CC=$resolvedCc;Subject=$subject;Sent=(Get-Date -Format s);Attachment=$xlsx;Replied="";Reminded=""}
                    appendSentRecord $record $sentFile
                    $sentCount++
                }else{
                    $mail.Display();$displayedCount++
                }
                log "$(if($Display){'Displayed'}else{'Sent'}): $vendorName" "SUCCESS"
            }catch{
                $failedCount++
                Write-Warning "Failed supplier $vendorName`: $($_.Exception.Message)"
            }finally{
                if($wb){try{$wb.Close($false)|Out-Null}catch{}}
                foreach($mcRecipient in @($mcRecipients)){releaseCom $mcRecipient};releaseCom $mailAttachment;releaseCom $recipients;releaseCom $mail
                releaseCom $materialNeedRange;releaseCom $header;releaseCom $range;releaseCom $ws;releaseCom $wb
            }
        }
    }finally{
        if($xl){try{$xl.Quit()|Out-Null}catch{}}
        releaseCom $xl;releaseCom $ol
    }
    terminalProgressDone "Send MR"
    terminalSection "Result"
    log ("Duration: {0}" -f (terminalDuration))
    $sendSummary="Send complete: $sentCount sent, $displayedCount displayed, $failedCount failed"
    if($failedCount -gt 0){log $sendSummary "WARNING"}else{log $sendSummary "SUCCESS"}
}

if($Mode -eq "scan"){
    $track=@();if(Test-Path $sentFile){$track=Import-Csv $sentFile}
    $pending=$track|?{!$_.Replied -and (vendorSelected $_.Vendor)}
    log "Pending requests: $(@($pending).Count)"
    $weekStart=(Get-Date).Date.AddDays(-[int](Get-Date).DayOfWeek);$weekEnd=$weekStart.AddDays(7);$weekFolder="{0:dd.MM}-{1:dd.MM.yyyy}" -f $weekStart,$weekStart.AddDays(6)
    $saveDir=Join-Path (Join-Path $OutDir "Replies") $weekFolder;New-Item -ItemType Directory -Force -Path $saveDir|Out-Null
    $sourceIndex=@{};$fallbackIndex=@{}
    foreach($masterRow in @($data)){
        if($sk){
            $sourceKey=([string]$masterRow.$sk).Trim()
            if($sourceKey){if(!$sourceIndex.ContainsKey($sourceKey)){$sourceIndex[$sourceKey]=[Collections.Generic.List[object]]::new()};$sourceIndex[$sourceKey].Add($masterRow)}
        }
        $fallbackKey=replyMatchKey $masterRow.$vc $masterRow.$pc $masterRow.$qc $masterRow.$dc
        if(!$fallbackIndex.ContainsKey($fallbackKey)){$fallbackIndex[$fallbackKey]=[Collections.Generic.List[object]]::new()}
        $fallbackIndex[$fallbackKey].Add($masterRow)
    }

    $ol=$null;$namespace=$null;$inbox=$null;$scanFolder=$null;$items=$null;$xl=$null
    $changed=0;$matched=0;$files=0;$supplyRows=0;$skippedOriginal=0
    try{
        $ol=New-Object -ComObject Outlook.Application
        $namespace=$ol.GetNamespace("MAPI");$inbox=$namespace.GetDefaultFolder(6);$scanFolder=$inbox
        if($ReplyFolder){try{$scanFolder=$inbox.Folders.Item($ReplyFolder)}catch{Write-Warning "Folder Inbox\$ReplyFolder not found, scanning Inbox instead"}}
        log "Scan folder: $($scanFolder.FolderPath)"
        log "Scan week: $weekFolder"
        log "Reply output folder: $saveDir"
        $items=$scanFolder.Items;$items.Sort("[ReceivedTime]",$true)
        $itemTotal=0
        try{$itemTotal=[int]$items.Count}catch{}
        terminalProgress "Scan replies" "Scanning Outlook replies..." 0 $itemTotal
        $itemNo=0
        $xl=New-Object -ComObject Excel.Application
        $xl.DisplayAlerts=$false;$xl.ScreenUpdating=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=3
        try{$xl.Calculation=-4135}catch{}
        foreach($m in $items){
            $itemNo++
            terminalProgress "Scan replies" ("Scanning Outlook item {0}/{1}" -f $itemNo,$itemTotal) $itemNo $itemTotal
            try{
            try{if($m.ReceivedTime -ge $weekEnd){continue};if($m.ReceivedTime -lt $weekStart){break}}catch{continue}
            $subject=[string]$m.Subject
            if($subject -notmatch 'MR_REQUEST\|(.+?)\|\d{8}_\d{6}'){continue}
            $replyVendor=$Matches[1].Trim()
            if($subject -notmatch '^\s*(RE|FW|FWD)\s*:'){$skippedOriginal++;continue}
            if(!(vendorSelected $replyVendor)){continue}
                    $matched++;log "Matched reply: $replyVendor / $($m.ReceivedTime)" "SUCCESS"
            for($attachmentIndex=1;$attachmentIndex -le $m.Attachments.Count;$attachmentIndex++){
                $a=$null;$wb=$null;$ws=$null;$used=$null
                try{
                    $a=$m.Attachments.Item($attachmentIndex)
                    $extension=[IO.Path]::GetExtension([string]$a.FileName)
                    if($extension -notmatch '^(\.xls[xm]?|\.csv)$'){continue}
                    $files++
                    $safeAttachmentName=([IO.Path]::GetFileName([string]$a.FileName) -replace '[\\/:*?"<>|]','_')
                    $savedName="{0}_{1}_{2}" -f (Get-Date -Format "yyyyMMdd_HHmmss_fff"),$attachmentIndex,$safeAttachmentName
                    $path=Join-Path $saveDir $savedName
                    $a.SaveAsFile($path)
                    log "Saved attachment: $path"

                    $xl.AutomationSecurity=3
                    $wb=$xl.Workbooks.Open($path)
                    $ws=$wb.Worksheets.Item(1);$used=$ws.UsedRange
                    $rows=$used.Rows.Count;$cols=$used.Columns.Count;$values=$used.Value2
                    $replyHeaders=@()
                    for($headerIndex=0;$headerIndex -lt $cols;$headerIndex++){$replyHeaders+=([string](cellValue $values 1 ($headerIndex+1))).Trim()}
                    $map=@{};for($headerIndex=0;$headerIndex -lt $replyHeaders.Count;$headerIndex++){$map[$replyHeaders[$headerIndex]]=$headerIndex+1}
                    $traceReplyIndexes=columnIndexes $replyHeaders $traceCols $false
                    $xkey=$traceReplyIndexes["_SourceKey"];$xrq=$map["Reply Qty"];$xpc=$map["Part No."];if(!$xpc){$xpc=$map["Part No"]};$xqc=$map["Request Qty"];if(!$xqc){$xqc=$map["New Request Qty"]};if(!$xqc){$xqc=$map["Shortage Qty"]};$xdc=$map["Material Need By Date"];if(!$xdc){$xdc=$map["Material need by date"]};if(!$xdc){$xdc=$map["Material Ready Date"]};$xsc=$map["ETA Vendor can Supply"];if(!$xsc){$xsc=$map["ETA Vendor can Suppy"]};if(!$xsc){$xsc=$map["SUPPLIER NEED CHECK"]};if(!$xsc){$xsc=$map["Vendor supply"]};$xif=$map["IF CAN NOT"];if(!$xif){$xif=$map["Remark"]}
                    $badEtaRows=@(invalidEtaRows $values $xsc $rows)
                    if($badEtaRows.Count){
                        Write-Warning ("Skipped attachment {0}: invalid ETA Vendor can Supply at row(s) {1}. Expected a valid date, e.g. 29-Jul-2026." -f $a.FileName,($badEtaRows -join ", "))
                        continue
                    }

                    $attachmentChanged=0
                    if($xkey -and $rows -ge 2){
                        for($replyRow=2;$replyRow -le $rows;$replyRow++){
                            $key=([string](cellValue $values $replyRow $xkey)).Trim()
                            if(!$key -or !$sourceIndex.ContainsKey($key)){continue}
                            $replyQty=if($xrq){([string](cellValue $values $replyRow $xrq)).Trim()}else{""}
                            $check=if($xsc){([string](cellValue $values $replyRow $xsc)).Trim()}else{""};if($check){$check=normalizeEtaDate $check}
                            $cannot=if($xif){([string](cellValue $values $replyRow $xif)).Trim()}else{""}
                            if(!$replyQty -and !$check -and !$cannot){continue}
                            $supplyRows++
                            foreach($targetRow in @($sourceIndex[$key])){
                                if((norm $targetRow.$vc) -ne (norm $replyVendor)){continue}
                                $rowChanged=$false
                                if($replyQty -and [string]$targetRow.$rq -ne $replyQty){$targetRow.$rq=$replyQty;$rowChanged=$true}
                                if($check -and [string]$targetRow.$mc -ne $check){$targetRow.$mc=$check;$rowChanged=$true}
                                if($cannot -and [string]$targetRow.$ifCannot -ne $cannot){$targetRow.$ifCannot=$cannot;$rowChanged=$true}
                                if($rowChanged){$changed++;$attachmentChanged++}
                            }
                        }
                    }elseif($xpc -and $xqc -and $xdc -and ($xsc -or $xif) -and $rows -ge 2){
                        for($replyRow=2;$replyRow -le $rows;$replyRow++){
                            $p=([string](cellValue $values $replyRow $xpc)).Trim();$q=([string](cellValue $values $replyRow $xqc)).Trim();$d=([string](cellValue $values $replyRow $xdc)).Trim()
                            $supply=if($xsc){([string](cellValue $values $replyRow $xsc)).Trim()}else{""};if($supply){$supply=normalizeEtaDate $supply}
                            $cannot=if($xif){([string](cellValue $values $replyRow $xif)).Trim()}else{""}
                            if(!$supply -and !$cannot){continue}
                            $supplyRows++
                            $fallbackKey=replyMatchKey $replyVendor $p $q $d
                            if(!$fallbackIndex.ContainsKey($fallbackKey)){continue}
                            foreach($targetRow in @($fallbackIndex[$fallbackKey])){
                                $rowChanged=$false
                                if($supply -and [string]$targetRow.$mc -ne $supply){$targetRow.$mc=$supply;$rowChanged=$true}
                                if($cannot -and [string]$targetRow.$ifCannot -ne $cannot){$targetRow.$ifCannot=$cannot;$rowChanged=$true}
                                if($rowChanged){$changed++;$attachmentChanged++}
                            }
                        }
                    }else{
                        Write-Warning "Skipped attachment missing required headers: $($a.FileName)"
                    }
                    if($attachmentChanged -gt 0){
                        $replyTime=Get-Date -Format s
                        $track|?{$_.Subject -and $subject.IndexOf($_.Subject,[StringComparison]::OrdinalIgnoreCase) -ge 0}|%{$_.Replied=$replyTime}
                    }
                    log "Updated rows from attachment: $attachmentChanged"
                }catch{
                    Write-Warning "Skipped attachment $([string]$a.FileName): $($_.Exception.Message)"
                }finally{
                    if($wb){try{$wb.Close($false)|Out-Null}catch{}}
                    releaseCom $used;releaseCom $ws;releaseCom $wb;releaseCom $a
                }
            }
            }catch{Write-Warning "Skipped Outlook item: $($_.Exception.Message)"}
        }
    }finally{
        if($xl){try{$xl.Quit()|Out-Null}catch{}}
        releaseCom $xl;releaseCom $items
        if($scanFolder -ne $inbox){releaseCom $scanFolder}
        releaseCom $inbox;releaseCom $namespace;releaseCom $ol
    }
    if($changed -gt 0){
        writeMasterOutputs $data $InputFile ([IO.Path]::ChangeExtension($InputFile,".xlsx")) $Template $masterCols $masterDisplayCols $delim
    }
    $track|Export-Csv $sentFile -NoTypeInformation -Encoding UTF8
    terminalProgressDone "Scan replies"
    terminalSection "Result"
    log ("Duration: {0}" -f (terminalDuration))
    log "Matched replies: $matched"
    log "Skipped original MR mails: $skippedOriginal"
    log "Excel attachments: $files"
    log "Rows with supplier response: $supplyRows"
    log "Updated rows: $changed"
    log "Scan complete" "SUCCESS"
}
