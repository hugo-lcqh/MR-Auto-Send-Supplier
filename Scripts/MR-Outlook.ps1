param(
    [ValidateSet("send","scan","remind","prepare")][string]$Mode="send",
    [string]$Vendor,
    [Alias("Input")][string]$InputFile="Input\MR-Outlook\input.txt",
    [string]$InputRoot="Input",
    [string]$Template="Input\Template.xlsx",
    [string]$Supplier="Config\suppliers.csv",
    [string]$OutDir="MR_Out",
    [string]$ReplyFolder="MR_REQUEST",
    [switch]$Display
)

$ErrorActionPreference="Stop"
$Root=if($env:AUTOTOOLS_ROOT){$env:AUTOTOOLS_ROOT}else{Split-Path -Parent $PSScriptRoot}
if(![IO.Path]::IsPathRooted($InputFile)){$InputFile=Join-Path $Root $InputFile}
if(![IO.Path]::IsPathRooted($InputRoot)){$InputRoot=Join-Path $Root $InputRoot}
if(![IO.Path]::IsPathRooted($Template)){$Template=Join-Path $Root $Template}
if(![IO.Path]::IsPathRooted($Supplier)){$Supplier=Join-Path $Root $Supplier}
if(![IO.Path]::IsPathRooted($OutDir)){$OutDir=Join-Path $Root $OutDir}
function col($o,$a,$req=$true){foreach($n in $a){$p=$o.PSObject.Properties[$n];if($p){return $p.Name}};if($req){throw "Missing column: $($a -join ', ')"}}
function val($o,$a){foreach($n in $a){$p=$o.PSObject.Properties[$n];if($p){return [string]$p.Value}};return ""}
function log($m){Write-Host ("[MR] {0}" -f $m)}
function norm($s){(([string]$s) -replace '\s+',' ').Trim().ToUpperInvariant()}
function addSupplierKey($map,$key,$row){$k=norm $key;if($k -and !$map.ContainsKey($k)){$map[$k]=$row}}
function findSupplier($name,$exact,$aliases){
    $k=norm $name
    if($exact.ContainsKey($k)){return $exact[$k]}
    foreach($a in @($aliases.Keys|Sort-Object Length -Descending)){
        if($k -eq $a -or $k.StartsWith("$a ") -or $k.StartsWith("$a(") -or $k.StartsWith("$a-")){return $aliases[$a]}
    }
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
function readTemplateHeaders($path){
    if(!(Test-Path -LiteralPath $path)){return @()}
    $xl=New-Object -ComObject Excel.Application;$xl.DisplayAlerts=$false
    try{
        $wb=$xl.Workbooks.Open($path)
        try{try{$ws=$wb.Worksheets.Item("MR")}catch{$ws=$wb.Worksheets.Item(1)};$used=$ws.UsedRange;$lastCol=$used.Column+$used.Columns.Count-1;$headers=@();for($c=1;$c -le $lastCol;$c++){$h=([string]$ws.Cells.Item(1,$c).Text).Trim();if($h){$headers+=$h}};return $headers}
        finally{$wb.Close($false)|Out-Null}
    }finally{$xl.Quit()|Out-Null}
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
}
function sourceNames($h){
    $d=displayHeader $h
    if($masterAliases[$d]){return $masterAliases[$d]}
    if($masterAliases[$h]){return $masterAliases[$h]}
    return @($d,$h)|Select-Object -Unique
}
function headerInfo($ws,$firstCol,$cols,$lastRow){
    $bestRow=0;$bestScore=0;$bestMap=@{}
    for($r=1;$r -le [Math]::Min(100,$lastRow);$r++){
        $map=@{};for($c=$firstCol;$c -lt $firstCol+$cols;$c++){$h=norm $ws.Cells.Item($r,$c).Text;if($h -and !$map.ContainsKey($h)){$map[$h]=$c}}
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
    if($s -notmatch '^\d{1,2}/\d{1,2}/\d{4}$'){return $null}
    [string[]]$formats=@("d/M/yyyy","dd/M/yyyy","d/MM/yyyy","dd/MM/yyyy")
    $dt=[datetime]::MinValue
    if([datetime]::TryParseExact($s,$formats,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$dt)){return $dt}
    return $null
}
function validEtaDate($value){return $null -ne (etaDateValue $value)}
function normalizeEtaDate($value){
    $dt=etaDateValue $value
    if($null -eq $dt){return ""}
    return $dt.ToString("dd/MM/yyyy")
}
function configureEtaInput($ws,$headers,$lastRow){
    $etaCol=headerIndex $headers @("ETA Vendor can Supply")
    if(!$etaCol -or $lastRow -lt 2){return}
    $etaRange=$ws.Range($ws.Cells.Item(2,$etaCol),$ws.Cells.Item($lastRow,$etaCol))
    $etaRange.NumberFormat="@"
    $etaCell=$ws.Cells.Item(2,$etaCol).Address($false,$false)
    $formula=('=OR({0}="",IFERROR(LET(s,{0},p,FIND("/",s),q,FIND("/",s,p+1),d,--LEFT(s,p-1),m,--MID(s,p+1,q-p-1),y,--RIGHT(s,4),x,DATE(y,m,d),AND(p>=2,p<=3,q-p>=2,q-p<=3,q=LEN(s)-4,DAY(x)=d,MONTH(x)=m,YEAR(x)=y)),FALSE))' -f $etaCell)
    $etaRange.Validation.Delete()
    $etaRange.Validation.Add(7,1,1,$formula)
    $etaRange.Validation.IgnoreBlank=$true
    $etaRange.Validation.ShowInput=$true
    $etaRange.Validation.InputTitle="ETA Vendor can Supply"
    $etaRange.Validation.InputMessage="Enter day/month/year, e.g. 22/2/2026 or 22/02/2026."
    $etaRange.Validation.ShowError=$true
    $etaRange.Validation.ErrorTitle="Invalid ETA"
    $etaRange.Validation.ErrorMessage="Enter a valid day-first date, e.g. 22/2/2026 or 22/02/2026."
}
function invalidEtaRows($ws,$etaCol,$lastRow){
    if(!$etaCol -or $lastRow -lt 2){return}
    $invalid=@()
    for($row=2;$row -le $lastRow;$row++){
        $eta=([string]$ws.Cells.Item($row,$etaCol).Text).Trim()
        if($eta -and !(validEtaDate $eta)){$invalid+=$row}
    }
    return $invalid
}
function cellMasterValue($ws,$values,$row,$absRow,$col,$absCol,$header){
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
function exportMasterXlsx($txtPath,$xlsx,$templatePath,$headers,$displayHeaders){
    $xl=New-Object -ComObject Excel.Application;$xl.DisplayAlerts=$false
    try{
        $fromTemplate=Test-Path -LiteralPath $templatePath
        if($fromTemplate){if(Test-Path -LiteralPath $xlsx){Remove-Item -LiteralPath $xlsx -Force};Copy-Item -LiteralPath $templatePath -Destination $xlsx -Force;$wb=$xl.Workbooks.Open($xlsx)}else{$wb=$xl.Workbooks.Add()}
        try{
            try{$ws=$wb.Worksheets.Item("MR")}catch{$ws=$wb.Worksheets.Item(1);$ws.Name="MR"}
            $used=$ws.UsedRange;$lastUsedRow=$used.Row+$used.Rows.Count-1;$lastUsedCol=[Math]::Max($headers.Count,$used.Column+$used.Columns.Count-1)
            $ws.Range($ws.Cells.Item(1,1),$ws.Cells.Item($lastUsedRow,$lastUsedCol)).ClearContents()|Out-Null
            $list=@(Import-Csv -LiteralPath $txtPath -Delimiter "`t")
            if($list.Count -gt 0){
                $dataHeaders=@($list[0].PSObject.Properties.Name)
                $arr=New-Object 'object[,]' $list.Count,$dataHeaders.Count
                for($r=0;$r -lt $list.Count;$r++){for($c=0;$c -lt $dataHeaders.Count;$c++){$v=[string]$list[$r].($dataHeaders[$c]);$arr[$r,$c]=if(isDateHeader $dataHeaders[$c]){excelDateValue $v}else{$v}}}
                $ws.Range($ws.Cells.Item(2,1),$ws.Cells.Item($list.Count+1,$dataHeaders.Count)).Value2=$arr
            }
            for($c=0;$c -lt $displayHeaders.Count;$c++){$ws.Cells.Item(1,$c+1).Value2=$displayHeaders[$c]}
            for($c=0;$c -lt $displayHeaders.Count;$c++){if(isDateHeader $headers[$c]){$ws.Columns.Item($c+1).NumberFormat=dateNumberFormat $headers[$c]}}
            if($ws.AutoFilterMode){$ws.AutoFilterMode=$false}
            $ws.Range($ws.Cells.Item(1,1),$ws.Cells.Item(1,$headers.Count)).AutoFilter()|Out-Null
            $ws.Range($ws.Cells.Item(1,1),$ws.Cells.Item(1,$headers.Count)).Font.Bold=$true
            $ws.Columns.AutoFit()|Out-Null
            if($fromTemplate){$wb.Save()}else{$wb.SaveAs($xlsx,51)}
            log "Master xlsx created: $xlsx"
        }finally{$wb.Close($false)|Out-Null}
    }finally{$xl.Quit()|Out-Null}
}
function buildMaster($folder,$masterFile){
    $rows=[System.Collections.ArrayList]::new();$xl=New-Object -ComObject Excel.Application;$xl.DisplayAlerts=$false;try{$xl.AutomationSecurity=3}catch{}
    try{
        $files=@(Get-ChildItem -LiteralPath $folder -File|?{$_.Name -notlike '~$*' -and $_.Name -notlike 'MR_Master_Input*' -and $_.Extension -match '^\.xls'})
        log "Input folder: $folder"
        log "Excel files to read: $($files.Count)"
        $fileNo=0
        foreach($file in $files){
            $fileNo++;$before=$rows.Count
            log "Reading $fileNo/$($files.Count): $($file.Name)"
            $wb=$null
            log "Opening workbook: $($file.Name)"
            try{$wb=$xl.Workbooks.Open($file.FullName)}catch{Write-Warning "Cannot open input workbook: $($file.Name) - $($_.Exception.Message)";continue}
            log "Opened workbook: $($file.Name)"
            try{
                try{$ws=$wb.Worksheets.Item("MR")}catch{Write-Warning "Sheet MR not found: $($file.Name)";continue}
                $used=$ws.UsedRange;$firstRow=$used.Row;$firstCol=$used.Column;$cols=$used.Columns.Count;$lastRow=$firstRow+$used.Rows.Count-1
                log "Sheet MR size: $lastRow rows x $cols columns"
                $hi=headerInfo $ws $firstCol $cols $lastRow;$map=$hi.Map
                if($hi.Score -lt 3){Write-Warning "Cannot detect header row in MR sheet: $($file.Name)";continue}
                log "Header row: $($hi.Row) ($($hi.Score) matched columns)"
                $values=$used.Value2
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
                    if((($r-$hi.Row) % 500) -eq 0){log "Reading rows: $($r-$hi.Row)/$($lastDataRow-$hi.Row) from $($file.Name)"}
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
            }finally{if($wb){$wb.Close($false)|Out-Null}}
            log "Rows added from $($file.Name): $($rows.Count-$before)"
        }
    }finally{$xl.Quit()|Out-Null}
    if($rows.Count -eq 0){throw "No rows found in sheet MR under $folder"}
    exportTsv $rows $masterFile
    log "Master input created: $masterFile ($($rows.Count) rows)"
    exportMasterXlsx $masterFile ([IO.Path]::ChangeExtension($masterFile,".xlsx")) $Template $masterCols $masterDisplayCols
}

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$sentFile=Join-Path $OutDir "mr_sent.csv"
log "Mode: $Mode"
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
$qc=col $data[0] @("Shortage Qty","New Request Qty","Request Qty","Qty")
$dc=col $data[0] @("Material Need By Date","Material need by date","Material Ready Date","MR Date","Ready Date")
$rq=col $data[0] @("Reply Qty") $false;if(!$rq){$rq="Reply Qty";$data|%{$_|Add-Member NoteProperty $rq "" -Force}}
$mc=col $data[0] @("ETA Vendor can Supply","ETA Vendor can Suppy","SUPPLIER NEED CHECK","MC double check","MC Double Check") $false
if(!$mc){$mc="SUPPLIER NEED CHECK";$data|%{$_|Add-Member NoteProperty $mc "" -Force}}
$ifCannot=col $data[0] @("IF CAN NOT","Remark") $false
if(!$ifCannot){$ifCannot="Remark";$data|%{$_|Add-Member NoteProperty $ifCannot "" -Force}}
$sk=col $data[0] @("_SourceKey") $false

if($Mode -eq "prepare"){
    log "Prepare complete. Master files are ready under: $latestFolder"
    return
}

if($Mode -eq "send"){
    $sup=Import-Csv -LiteralPath $Supplier
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
    $ol=New-Object -ComObject Outlook.Application
    $xl=New-Object -ComObject Excel.Application;$xl.DisplayAlerts=$false
    $new=@()
    $outCols=@($data[0].PSObject.Properties.Name)
    $outHeaders=@($outCols|%{supplierHeader $_})
    $hiddenSupplierCols=@($traceCols+@(
        "Item Seq","Follow By Mps","Change Type","Adjust date","Batch No","Phase Code",
        "Item No","Master pack","Item Type","Description","Customer Item","End Buyer",
        "Customer Name","Region","Customer PO Number","Order Number",
        "Customer Due Date","Original Qty","Original Promise Date","Original Schedule Ship Date",
        "Request Qty","Request Date","Order Priority","New Request Qty","New Request Date",
        "Reply Qty","Reply Promise Date","Reply Schedule Date","New Priority","Material Ready Date",
        "Part Remarks","MPS Remark","MC Remarks for MR","MC Remark","Planner Code","Buyer","Part Desc","Vendor Name","End Supplier Name","Lead Time",
        "Outstanding","Materials Gap","OTD Gap","Creation Date","Release Date","Ecc Schedule Ship Date","Source Line Id","supply Source","Bu Category","System update on","ASCP supply",
        "ASCP supply Alert","Next Shortage Date","Highlight next shortage (if within 1M)"
    ))
    $hiddenSupplierIndexes=columnIndexes $outHeaders $hiddenSupplierCols $false
    $aliases=@{
        "Request Qty"=@("Request Qty","New Request Qty","Qty")
        "Part No."=@("Part No.","Part No","PartNo")
        "Material Need By Date"=@("Material Need By Date","Material need by date","Material Ready Date","MR Date","Ready Date")
        "MC double check"=@("ETA Vendor can Supply","ETA Vendor can Suppy","SUPPLIER NEED CHECK","MC double check","MC Double Check")
        "SUPPLIER NEED CHECK"=@("ETA Vendor can Supply","ETA Vendor can Suppy","SUPPLIER NEED CHECK","MC double check","MC Double Check")
        "Remark"=@("IF CAN NOT","Remark")
    }
    try{
        $groups=@($data|?{!$Vendor -or $_.$vc -like "*$Vendor*"}|Group-Object $vc)
        log "Suppliers to process: $($groups.Count)"
        $sendNo=0
        $groups|%{
            $sendNo++
            $vendorName=$_.Name.Trim();if(!$vendorName){return}
            log "Sending $sendNo/$($groups.Count): $vendorName ($(@($_.Group).Count) rows)"
            $s=findSupplier $vendorName $byVendor $supplierAliases;if(!$s -or [string]::IsNullOrWhiteSpace([string]$s.$to)){Write-Warning "No email: $vendorName";return}
            $stamp=Get-Date -Format "yyyyMMdd_HHmmss"
            $safe=$vendorName -replace '[\\/:*?"<>|]','_'
            $xlsx=Join-Path (Resolve-Path $OutDir) "MR_${safe}_$stamp.xlsx"
            $wb=$xl.Workbooks.Add();$ws=$wb.Worksheets.Item(1);$ws.Name="MR";$ws.Columns.NumberFormat="@"
            $c=1;$outHeaders|%{$ws.Cells.Item(1,$c++).Value2=$_}
            $r=2;foreach($x in $_.Group){$c=1;foreach($h in $outCols){$names=if($aliases[$h]){$aliases[$h]}else{@($h)};$ws.Cells.Item($r,$c++).Value2=val $x $names};$r++}
            $range=$ws.Range($ws.Cells.Item(1,1),$ws.Cells.Item([Math]::Max(1,$r-1),$outCols.Count))
            $range.Borders.LineStyle=1;$range.Borders.Color=14277081
            $header=$ws.Range($ws.Cells.Item(1,1),$ws.Cells.Item(1,$outCols.Count));$header.Font.Bold=$true;$header.Font.Color=16777215;$header.Interior.Color=6299648;$header.AutoFilter()|Out-Null
            $importantCols=@("Part No.","Shortage Qty","Material need by date","ETA Vendor can Supply","IF CAN NOT")
            $importantCols|%{$ic=headerIndex $outHeaders @($_);if($ic -gt 0){$ws.Cells.Item(1,$ic).Interior.Color=255;$ws.Cells.Item(1,$ic).Font.Color=16777215;if($r -gt 2){$ws.Range($ws.Cells.Item(2,$ic),$ws.Cells.Item($r-1,$ic)).Interior.Color=13434879}}}
            addGuidelineSheet $wb $ws
            $ws.Activate()|Out-Null
            for($i=1;$i -le $outCols.Count;$i++){$ws.Columns.Item($i).ColumnWidth=14}
            $importantCols|%{$ic=headerIndex $outHeaders @($_);if($ic -gt 0){$ws.Columns.Item($ic).ColumnWidth=24}}
            configureEtaInput $ws $outHeaders ($r-1)
            $hiddenSupplierCols|%{$ic=$hiddenSupplierIndexes[$_];if($ic -gt 0){$ws.Columns.Item($ic).Hidden=$true}}
            $ws.Rows.Item(1).RowHeight=24;$ws.Application.ActiveWindow.SplitRow=1;$ws.Application.ActiveWindow.FreezePanes=$true;$wb.SaveAs($xlsx,51);$wb.Close($false)
            log "Attachment created: $xlsx"
            $subject="MR_REQUEST|$vendorName|$stamp"
            $mcInCharge=if($mcContact){([string]$s.$mcContact).Trim()}else{([string](@($s.PSObject.Properties)[5].Value)).Trim()}
            if($mcInCharge -in @("TRUE","FALSE")){$mcInCharge=""}
            $ccList=@();if($cc){$ccList+=[string]$s.$cc};if($mcInCharge){$ccList+=$mcInCharge}
            $mail=$ol.CreateItem(0);$mail.To=[string]$s.$to;$mail.CC=($ccList|?{!([string]::IsNullOrWhiteSpace($_))}|Select-Object -Unique) -join ";"
            $mail.Recipients.ResolveAll()|Out-Null
            $mcTag="";if($mcInCharge){$mcName=[string]$mail.Recipients.Item($mail.Recipients.Count).Name;$mcTag=" <span style='color:#9db9ff;background:#263f6b;border-radius:3px;padding:1px 3px;'>$([Net.WebUtility]::HtmlEncode("@$mcName"))</span>"}
            $mail.Subject=$subject;$mail.BodyFormat=2;$mail.HTMLBody=@"
<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:1.25;">
<b>Dear Supplier,</b><br>
This is a request to add in / pull in material. Please kindly check the details and provide your feedback with the following<br>
<b>information:</b>
<ol style="margin-top:8px;margin-bottom:8px;">
  <li>Fill in column <b>AL</b> with the ETA in <b>DD/MM/YYYY</b> format.</li>
  <li>If your provided ETA cannot meet our requested date, please fill in column <b>AM</b> with the detailed reason, such as:
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
Your feedback is required within <b>24 hours</b>. We appreciate your prompt cooperation and support.<br>
<b>Dear MCs,</b>$mcTag<br>
Please kindly follow up on this request with your suppliers and ensure their feedback is received within the required timeline.<br>
Thank you.
</div>
"@
            $mail.Attachments.Add($xlsx)|Out-Null;if($Display){$mail.Display()}else{$mail.Send()}
            $new+=[pscustomobject]@{Vendor=$vendorName;To=$mail.To;CC=$mail.CC;Subject=$subject;Sent=(Get-Date -Format s);Attachment=$xlsx;Replied="";Reminded=""}
            log "$(if($Display){'Displayed'}else{'Sent'}): $vendorName"
        }
    }finally{$xl.Quit()|Out-Null}
    $old=@();if(Test-Path $sentFile){$old=@(Import-Csv $sentFile)}
    @($old+$new)|Export-Csv $sentFile -NoTypeInformation -Encoding UTF8
    log "Send complete: $(@($new).Count) supplier mails"
}

if($Mode -eq "scan"){
    $track=@();if(Test-Path $sentFile){$track=Import-Csv $sentFile}
    $pending=$track|?{!$_.Replied -and (!$Vendor -or $_.Vendor -like "*$Vendor*")}
    log "Pending requests: $(@($pending).Count)"
    $weekStart=(Get-Date).Date.AddDays(-[int](Get-Date).DayOfWeek);$weekEnd=$weekStart.AddDays(7);$weekFolder="{0:dd.MM}-{1:dd.MM.yyyy}" -f $weekStart,$weekStart.AddDays(6)
    $saveDir=Join-Path (Join-Path $OutDir "Replies") $weekFolder;New-Item -ItemType Directory -Force -Path $saveDir|Out-Null
    $ol=New-Object -ComObject Outlook.Application;$inbox=$ol.GetNamespace("MAPI").GetDefaultFolder(6);$scanFolder=$inbox
    if($ReplyFolder){try{$scanFolder=$inbox.Folders.Item($ReplyFolder)}catch{Write-Warning "Folder Inbox\$ReplyFolder not found, scanning Inbox instead"}}
    log "Scan folder: $($scanFolder.FolderPath)"
    log "Scan week: $weekFolder"
    log "Reply output folder: $saveDir"
    $items=$scanFolder.Items;$items.Sort("[ReceivedTime]",$true)
    $xl=New-Object -ComObject Excel.Application;$xl.DisplayAlerts=$false;$changed=0;$matched=0;$files=0;$supplyRows=0;$skippedOriginal=0
    try{
        foreach($m in $items){
            try{if($m.ReceivedTime -ge $weekEnd){continue};if($m.ReceivedTime -lt $weekStart){break}}catch{continue}
            $subject=[string]$m.Subject
            if($subject -notmatch 'MR_REQUEST\|(.+?)\|\d{8}_\d{6}'){continue}
            $replyVendor=$Matches[1].Trim()
            if($subject -notmatch '^\s*(RE|FW|FWD)\s*:'){$skippedOriginal++;continue}
            if($Vendor -and $replyVendor -notlike "*$Vendor*"){continue}
            $matched++;log "Matched reply: $replyVendor / $($m.ReceivedTime)"
            for($i=1;$i -le $m.Attachments.Count;$i++){
                $a=$m.Attachments.Item($i);if($a.FileName -notmatch '\.xls[xm]?$'){continue}
                $files++
                $path=Join-Path $saveDir ("{0}_{1}" -f (Get-Date -Format "yyyyMMdd_HHmmss"),$a.FileName);$a.SaveAsFile($path)
                log "Saved attachment: $path"
                $beforeChanged=$changed
                $wb=$xl.Workbooks.Open($path);$ws=$wb.Worksheets.Item(1);$rows=$ws.UsedRange.Rows.Count;$cols=$ws.UsedRange.Columns.Count
                $replyHeaders=@(1..$cols|%{([string]$ws.Cells.Item(1,$_).Text).Trim()})
                $map=@{};for($i=0;$i -lt $replyHeaders.Count;$i++){$map[$replyHeaders[$i]]=$i+1}
                $traceReplyIndexes=columnIndexes $replyHeaders $traceCols $false
                $xkey=$traceReplyIndexes["_SourceKey"];$xrq=$map["Reply Qty"];$xpc=$map["Part No."];if(!$xpc){$xpc=$map["Part No"]};$xqc=$map["Shortage Qty"];if(!$xqc){$xqc=$map["New Request Qty"]};$xdc=$map["Material Need By Date"];if(!$xdc){$xdc=$map["Material need by date"]};if(!$xdc){$xdc=$map["Material Ready Date"]};$xsc=$map["ETA Vendor can Supply"];if(!$xsc){$xsc=$map["ETA Vendor can Suppy"]};if(!$xsc){$xsc=$map["SUPPLIER NEED CHECK"]};if(!$xsc){$xsc=$map["Vendor supply"]};$xif=$map["IF CAN NOT"];if(!$xif){$xif=$map["Remark"]}
                $badEtaRows=@(invalidEtaRows $ws $xsc $rows)
                if($badEtaRows.Count){
                    Write-Warning ("Skipped attachment {0}: invalid ETA Vendor can Supply at row(s) {1}. Expected day/month/year, e.g. 22/2/2026 or 22/02/2026." -f $a.FileName,($badEtaRows -join ", "))
                    $wb.Close($false)
                    continue
                }
                if($xkey -and $rows -ge 2){
                    2..$rows|%{
                        $key=([string]$ws.Cells.Item($_,$xkey).Text).Trim()
                        if($key){$replyQty=if($xrq){([string]$ws.Cells.Item($_,$xrq).Text).Trim()}else{""};$check=if($xsc){([string]$ws.Cells.Item($_,$xsc).Text).Trim()}else{""};if($check){$check=normalizeEtaDate $check};$cannot=if($xif){([string]$ws.Cells.Item($_,$xif).Text).Trim()}else{""};if($replyQty -or $check -or $cannot){$supplyRows++;$data|?{$_._SourceKey -eq $key}|%{if($replyQty){$_.$rq=$replyQty};if($check){$_.$mc=$check};if($cannot){$_.$ifCannot=$cannot};$changed++}}}
                    }
                    $track|?{$_.Vendor -eq $replyVendor -or $subject.IndexOf($_.Subject,[StringComparison]::OrdinalIgnoreCase) -ge 0}|%{$_.Replied=Get-Date -Format s}
                }elseif($xpc -and $xqc -and $xdc -and ($xsc -or $xif) -and $rows -ge 2){
                    2..$rows|%{
                        $p=([string]$ws.Cells.Item($_,$xpc).Text).Trim();$q=([string]$ws.Cells.Item($_,$xqc).Text).Trim();$d=([string]$ws.Cells.Item($_,$xdc).Text).Trim();$supply=if($xsc){([string]$ws.Cells.Item($_,$xsc).Text).Trim()}else{""};if($supply){$supply=normalizeEtaDate $supply};$cannot=if($xif){([string]$ws.Cells.Item($_,$xif).Text).Trim()}else{""}
                        if($supply -or $cannot){$supplyRows++;$data|?{$_.$vc -eq $replyVendor -and ([string]$_.$pc).Trim() -eq $p -and ([string]$_.$qc).Trim() -eq $q -and ([string]$_.$dc).Trim() -eq $d}|%{if($supply){$_.$mc=$supply};if($cannot){$_.$ifCannot=$cannot};$changed++}}
                    }
                    $track|?{$_.Vendor -eq $replyVendor -or $subject.IndexOf($_.Subject,[StringComparison]::OrdinalIgnoreCase) -ge 0}|%{$_.Replied=Get-Date -Format s}
                }else{
                    Write-Warning "Skipped attachment missing required headers: $($a.FileName)"
                }
                $wb.Close($false)
                log "Updated rows from attachment: $($changed-$beforeChanged)"
            }
        }
    }finally{$xl.Quit()|Out-Null}
    if($delim -eq "`t"){
        exportTsv $data $InputFile
        exportMasterXlsx $InputFile ([IO.Path]::ChangeExtension($InputFile,".xlsx")) $Template $masterCols $masterDisplayCols
    }else{
        $data|Export-Csv $InputFile -NoTypeInformation -Encoding UTF8
    }
    $track|Export-Csv $sentFile -NoTypeInformation -Encoding UTF8
    log "Matched replies: $matched"
    log "Skipped original MR mails: $skippedOriginal"
    log "Excel attachments: $files"
    log "Rows with supplier response: $supplyRows"
    log "Updated cells: $changed"
}

if($Mode -eq "remind"){
    $track=@();if(Test-Path $sentFile){$track=Import-Csv $sentFile}
    log "Tracking rows: $(@($track).Count)"
    $ol=New-Object -ComObject Outlook.Application
    $sent=$ol.GetNamespace("MAPI").GetDefaultFolder(5);$items=$sent.Items;$items.Sort("[SentOn]",$true);$known=@{};$track|%{$known[$_.Subject]=$true};$cut=(Get-Date).AddDays(-30)
    foreach($m in $items){try{if($m.SentOn -lt $cut){break};if(([string]$m.Subject) -notlike "MR_REQUEST|*" -or $known[$m.Subject]){continue};$vn=([string]$m.Subject) -replace '^MR_REQUEST\|(.+)\|\d{8}_\d{6}.*$','$1';$track+=[pscustomobject]@{Vendor=$vn;To=$m.To;CC=$m.CC;Subject=$m.Subject;Sent=([datetime]$m.SentOn).ToString("s");Attachment="";Replied="";Reminded=""}}catch{}}
    $due=$track|?{!$_.Replied -and !$_.Reminded -and (!$Vendor -or $_.Vendor -like "*$Vendor*") -and ([datetime]$_.Sent).AddHours(24) -lt (Get-Date)}
    log "Reminders due: $(@($due).Count)"
    foreach($request in $due){
        $mail=newReminderReply $items $request.Subject $cut
        if(!$mail){Write-Warning "Skipped reminder: original Sent Items mail not found for Subject: $($request.Subject)";continue}
        $mail.BodyFormat=2
        $mail.HTMLBody='<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;"><b>Dear Supplier,</b><br><br>Please help reply the MR file today.<br><br>Thanks.</div><br>'+$mail.HTMLBody
        if(Test-Path $request.Attachment){$mail.Attachments.Add($request.Attachment)|Out-Null}
        if($Display){$mail.Display()}else{$mail.Send()}
        $request.Reminded=Get-Date -Format s;log "$(if($Display){'Displayed reminder'}else{'Sent reminder'}): $($request.Vendor)"
    }
    $track|Export-Csv $sentFile -NoTypeInformation -Encoding UTF8
    log "Remind complete"
}
