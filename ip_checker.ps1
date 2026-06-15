param([string]$InputFile, [string]$OutputFile)

$BatchSize = 100
$DelayMs   = 1500
$MaxRetry  = 3

# Doc file IP
try   { $raw = Get-Content $InputFile -Encoding UTF8 }
catch { $raw = Get-Content $InputFile -Encoding Default }

$ipList = $raw | Where-Object { $_ -match '\S' -and $_ -notmatch '^\s*#' } |
          ForEach-Object { $_.Trim() } | Select-Object -Unique
$total = $ipList.Count
Write-Host "  => Tim thay $total IP hop le"
Write-Host ""

$batches = @()
for ($i = 0; $i -lt $total; $i += $BatchSize) {
    $end = [Math]::Min($i + $BatchSize - 1, $total - 1)
    $batches += , @($ipList[$i..$end])
}
$totalBatches = $batches.Count
Write-Host "  Chia thanh $totalBatches batch (moi batch $BatchSize IP)"
Write-Host "  API: ip-api.com (mien phi, khong can key)"
Write-Host ""

$results = [System.Collections.Generic.List[hashtable]]::new()
$sw = [System.Diagnostics.Stopwatch]::StartNew()

for ($b = 0; $b -lt $totalBatches; $b++) {
    $batch   = $batches[$b]
    $payload = $batch | ForEach-Object {
        @{ query = $_; fields = 'status,message,country,countryCode,regionName,city,isp,org,as,query' }
    }
    $json = $payload | ConvertTo-Json -Compress
    $data = $null
    for ($r = 0; $r -lt $MaxRetry; $r++) {
        try {
            $data = Invoke-RestMethod -Uri 'http://ip-api.com/batch' -Method Post `
                    -Body $json -ContentType 'application/json' -TimeoutSec 15 -ErrorAction Stop
            break
        } catch {
            Start-Sleep -Milliseconds (1000 * [Math]::Pow(2, $r))
        }
    }
    for ($i = 0; $i -lt $batch.Count; $i++) {
        $ip = $batch[$i]
        $d  = if ($data) { $data[$i] } else { $null }
        if ($d -and $d.status -eq 'success') {
            $results.Add(@{
                stt     = ($results.Count + 1)
                ip      = $d.query
                country = $d.country
                code    = $d.countryCode
                region  = $d.regionName
                city    = $d.city
                isp     = $d.isp
                org     = $d.org
                asn     = $d.as
                status  = 'Success'
            })
        } else {
            $msg = if ($d) { $d.message } else { 'Request failed' }
            $results.Add(@{
                stt     = ($results.Count + 1)
                ip      = $ip
                country = 'N/A'; code = 'N/A'; region = 'N/A'; city = 'N/A'
                isp     = 'N/A'; org  = 'N/A'; asn    = 'N/A'
                status  = "Error: $msg"
            })
        }
    }
    $processed = ($b + 1) * $BatchSize
    if ($processed -gt $total) { $processed = $total }
    $elapsed = $sw.Elapsed.TotalSeconds
    $speed   = if ($elapsed -gt 0) { $processed / $elapsed } else { 1 }
    $eta     = [int](($total - $processed) / $speed)
    $pct     = [int]($processed * 100 / $total)
    $filled  = [int]($pct * 40 / 100)
    $bar     = ('#' * $filled) + ('-' * (40 - $filled))
    $etaStr  = '{0:D2}:{1:D2}' -f [int]($eta / 60), ($eta % 60)
    Write-Host "`r  [$bar] $pct%  Batch $($b+1)/$totalBatches  |  $processed/$total IP  |  ETA: $etaStr" -NoNewline
    if ($b -lt $totalBatches - 1) { Start-Sleep -Milliseconds $DelayMs }
}

$elapsed2 = $sw.Elapsed
$elStr    = '{0:D2}:{1:D2} phut' -f [int]$elapsed2.TotalMinutes, $elapsed2.Seconds
Write-Host ""
Write-Host ""
Write-Host "  => Hoan thanh $total IP trong $elStr"
Write-Host ""
Write-Host "  Dang tao file Excel..."

$succCnt = ($results | Where-Object { $_.status -eq 'Success' }).Count
$errCnt  = $total - $succCnt
$runDate = Get-Date -Format 'dd/MM/yyyy HH:mm:ss'
$cGroup  = $results | Where-Object { $_.status -eq 'Success' } |
           Group-Object country | Sort-Object Count -Descending

# ===== TAO EXCEL (Open XML / ZIP) =====
Add-Type -AssemblyName System.IO.Compression.FileSystem
$tmp = Join-Path $env:TEMP ('ipchk_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
New-Item -ItemType Directory (Join-Path $tmp 'xl')            | Out-Null
New-Item -ItemType Directory (Join-Path $tmp 'xl\_rels')      | Out-Null
New-Item -ItemType Directory (Join-Path $tmp 'xl\worksheets') | Out-Null
New-Item -ItemType Directory (Join-Path $tmp '_rels')         | Out-Null

# --- Shared Strings ---
$allStr = [System.Collections.Generic.List[string]]::new()
$strIdx = @{}
function GetStrIdx([string]$val) {
    if (-not $strIdx.ContainsKey($val)) {
        $strIdx[$val] = $allStr.Count
        $allStr.Add($val)
    }
    return $strIdx[$val]
}

# --- Column letter ---
function ColLetter([int]$col) {
    $result = ''
    $n = $col
    while ($n -gt 0) {
        $result = [char](65 + ($n - 1) % 26) + $result
        $n = [int](($n - 1) / 26)
    }
    return $result
}

# --- Cell builders (unique names to avoid alias conflicts) ---
function xmlSC([int]$row, [int]$col, [string]$val, [int]$style) {
    $colLetter = ColLetter $col
    $idx = GetStrIdx $val
    return "<c r=`"$colLetter$row`" t=`"s`" s=`"$style`"><v>$idx</v></c>"
}
function xmlNC([int]$row, [int]$col, $num, [int]$style) {
    $colLetter = ColLetter $col
    return "<c r=`"$colLetter$row`" s=`"$style`"><v>$num</v></c>"
}
function xmlFC([int]$row, [int]$col, [string]$formula, [int]$style) {
    $colLetter = ColLetter $col
    return "<c r=`"$colLetter$row`" s=`"$style`"><f>$formula</f></c>"
}

# Pre-index strings
$titleDetail  = "IP LOCATION REPORT  -  $runDate  -  Tong so: $total IP"
$titleCountry = "THONG KE THEO QUOC GIA  -  $($cGroup.Count) quoc gia  -  $succCnt IP thanh cong"

@('STT','IP Address','Country','Code','Region','City','ISP','Organization','AS Number','Status',
  'IP Count','% Total','BAO CAO TONG QUAN - IP LOCATION',
  'Ngay gio chay','File IP dau vao','Tong so IP (tu file)',
  'IP thanh cong','IP loi / khong xac dinh','So quoc gia phat hien',
  'Thoi gian xu ly','API su dung','ip-api.com (batch, mien phi)',
  'Success','N/A', $titleDetail, $titleCountry, $runDate, $InputFile, $elStr) |
    ForEach-Object { GetStrIdx $_ | Out-Null }

$results | ForEach-Object {
    GetStrIdx $_.ip      | Out-Null
    GetStrIdx $_.country | Out-Null
    GetStrIdx $_.code    | Out-Null
    GetStrIdx $_.region  | Out-Null
    GetStrIdx $_.city    | Out-Null
    GetStrIdx $_.isp     | Out-Null
    GetStrIdx $_.org     | Out-Null
    GetStrIdx $_.asn     | Out-Null
    GetStrIdx $_.status  | Out-Null
}
$cGroup | ForEach-Object { GetStrIdx $_.Name | Out-Null }

# --- Write XML files ---
function WriteFile([string]$path, [string]$content) {
    [System.IO.File]::WriteAllText($path, $content, [System.Text.Encoding]::UTF8)
}

WriteFile (Join-Path $tmp '[Content_Types].xml') @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
'@

WriteFile (Join-Path $tmp '_rels\.rels') @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@

WriteFile (Join-Path $tmp 'xl\_rels\workbook.xml.rels') @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"     Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"     Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"     Target="worksheets/sheet3.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"        Target="styles.xml"/>
</Relationships>
'@

WriteFile (Join-Path $tmp 'xl\workbook.xml') @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
  <sheet name="Summary"            sheetId="1" r:id="rId1"/>
  <sheet name="IP Location Detail" sheetId="2" r:id="rId2"/>
  <sheet name="Country Summary"    sheetId="3" r:id="rId3"/>
</sheets>
</workbook>
'@

WriteFile (Join-Path $tmp 'xl\styles.xml') @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="5">
  <font><sz val="10"/><name val="Arial"/></font>
  <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
  <font><b/><sz val="13"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
  <font><sz val="9"/><name val="Arial"/></font>
  <font><b/><sz val="11"/><name val="Arial"/></font>
</fonts>
<fills count="9">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFDEEAF1"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFBDD7EE"/></patternFill></fill>
</fills>
<borders count="2">
  <border><left/><right/><top/><bottom/><diagonal/></border>
  <border>
    <left style="thin"><color rgb="FFB8CCE4"/></left>
    <right style="thin"><color rgb="FFB8CCE4"/></right>
    <top style="thin"><color rgb="FFB8CCE4"/></top>
    <bottom style="thin"><color rgb="FFB8CCE4"/></bottom>
    <diagonal/>
  </border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="11">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
  <xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
  <xf numFmtId="0" fontId="3" fillId="6" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
  <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="9" fontId="3" fillId="4" borderId="1" xfId="0" applyAlignment="1" applyNumberFormat="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="3" fillId="7" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
  <xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
  <xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
</cellXfs>
</styleSheet>
'@

# ===== SHEET 2: IP Location Detail =====
Write-Host "    [Sheet 2/3] IP Location Detail..."
$sb = [System.Text.StringBuilder]::new(10 * 1024 * 1024)
[void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
[void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
[void]$sb.Append('<sheetViews><sheetView workbookViewId="0">')
[void]$sb.Append('<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/>')
[void]$sb.Append('</sheetView></sheetViews>')
[void]$sb.Append('<sheetFormatPr defaultRowHeight="18"/>')
[void]$sb.Append('<cols>')
[void]$sb.Append('<col min="1"  max="1"  width="6"  customWidth="1"/>')
[void]$sb.Append('<col min="2"  max="2"  width="18" customWidth="1"/>')
[void]$sb.Append('<col min="3"  max="3"  width="22" customWidth="1"/>')
[void]$sb.Append('<col min="4"  max="4"  width="7"  customWidth="1"/>')
[void]$sb.Append('<col min="5"  max="5"  width="22" customWidth="1"/>')
[void]$sb.Append('<col min="6"  max="6"  width="18" customWidth="1"/>')
[void]$sb.Append('<col min="7"  max="7"  width="32" customWidth="1"/>')
[void]$sb.Append('<col min="8"  max="8"  width="32" customWidth="1"/>')
[void]$sb.Append('<col min="9"  max="9"  width="28" customWidth="1"/>')
[void]$sb.Append('<col min="10" max="10" width="18" customWidth="1"/>')
[void]$sb.Append('</cols><sheetData>')

# Title row
[void]$sb.Append("<row r=`"1`" ht=`"28`">")
[void]$sb.Append((xmlSC 1 1 $titleDetail 2))
[void]$sb.Append("</row>")

# Header row
$hdrs = @('STT','IP Address','Country','Code','Region','City','ISP','Organization','AS Number','Status')
[void]$sb.Append('<row r="2" ht="22">')
for ($c = 1; $c -le $hdrs.Count; $c++) {
    [void]$sb.Append((xmlSC 2 $c $hdrs[$c - 1] 1))
}
[void]$sb.Append('</row>')

# Data rows
$rn = 3
foreach ($row in $results) {
    $isErr = ($row.status -ne 'Success')
    $isOdd = (($rn - 3) % 2 -eq 0)
    $sx    = if ($isErr) { 5 } elseif ($isOdd) { 3 } else { 4 }
    $sxC   = if ($isErr) { 5 } else { 6 }
    [void]$sb.Append("<row r=`"$rn`">")
    [void]$sb.Append((xmlNC $rn 1  $row.stt     $sxC))
    [void]$sb.Append((xmlSC $rn 2  $row.ip      $sx))
    [void]$sb.Append((xmlSC $rn 3  $row.country $sx))
    [void]$sb.Append((xmlSC $rn 4  $row.code    $sxC))
    [void]$sb.Append((xmlSC $rn 5  $row.region  $sx))
    [void]$sb.Append((xmlSC $rn 6  $row.city    $sx))
    [void]$sb.Append((xmlSC $rn 7  $row.isp     $sx))
    [void]$sb.Append((xmlSC $rn 8  $row.org     $sx))
    [void]$sb.Append((xmlSC $rn 9  $row.asn     $sx))
    [void]$sb.Append((xmlSC $rn 10 $row.status  $sx))
    [void]$sb.Append('</row>')
    $rn++
}
$lastRow = $rn - 1
[void]$sb.Append("</sheetData>")
[void]$sb.Append("<autoFilter ref=`"A2:J$lastRow`"/>")
[void]$sb.Append("<mergeCells count=`"1`"><mergeCell ref=`"A1:J1`"/></mergeCells>")
[void]$sb.Append("</worksheet>")
WriteFile (Join-Path $tmp 'xl\worksheets\sheet2.xml') $sb.ToString()

# ===== SHEET 3: Country Summary =====
Write-Host "    [Sheet 3/3] Country Summary..."
$sb3 = [System.Text.StringBuilder]::new()
[void]$sb3.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
[void]$sb3.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
[void]$sb3.Append('<sheetViews><sheetView workbookViewId="0">')
[void]$sb3.Append('<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/>')
[void]$sb3.Append('</sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/>')
[void]$sb3.Append('<cols>')
[void]$sb3.Append('<col min="1" max="1" width="8"  customWidth="1"/>')
[void]$sb3.Append('<col min="2" max="2" width="30" customWidth="1"/>')
[void]$sb3.Append('<col min="3" max="3" width="14" customWidth="1"/>')
[void]$sb3.Append('<col min="4" max="4" width="12" customWidth="1"/>')
[void]$sb3.Append('</cols><sheetData>')
[void]$sb3.Append("<row r=`"1`" ht=`"28`">")
[void]$sb3.Append((xmlSC 1 1 $titleCountry 2))
[void]$sb3.Append("</row>")
[void]$sb3.Append('<row r="2" ht="22">')
[void]$sb3.Append((xmlSC 2 1 'STT'      1))
[void]$sb3.Append((xmlSC 2 2 'Country'  1))
[void]$sb3.Append((xmlSC 2 3 'IP Count' 1))
[void]$sb3.Append((xmlSC 2 4 '% Total'  1))
[void]$sb3.Append('</row>')
$fills3 = @(4, 3, 8, 6)
for ($i = 0; $i -lt $cGroup.Count; $i++) {
    $r3 = $i + 3
    $fx = $fills3[$i % $fills3.Count]
    [void]$sb3.Append("<row r=`"$r3`">")
    [void]$sb3.Append((xmlNC $r3 1 ($i + 1)           6))
    [void]$sb3.Append((xmlSC $r3 2 $cGroup[$i].Name   $fx))
    [void]$sb3.Append((xmlNC $r3 3 $cGroup[$i].Count   6))
    [void]$sb3.Append((xmlFC $r3 4 "C$r3/$succCnt"     7))
    [void]$sb3.Append('</row>')
}
$lastC3 = $cGroup.Count + 2
[void]$sb3.Append("</sheetData>")
[void]$sb3.Append("<autoFilter ref=`"A2:D$lastC3`"/>")
[void]$sb3.Append("<mergeCells count=`"1`"><mergeCell ref=`"A1:D1`"/></mergeCells>")
[void]$sb3.Append("</worksheet>")
WriteFile (Join-Path $tmp 'xl\worksheets\sheet3.xml') $sb3.ToString()

# ===== SHEET 1: Summary =====
Write-Host "    [Sheet 1/3] Summary..."
$sumRows = @(
    @{ l = 'Ngay gio chay';           v = $runDate },
    @{ l = 'File IP dau vao';         v = $InputFile },
    @{ l = 'Tong so IP (tu file)';    v = "$total" },
    @{ l = 'IP thanh cong';           v = "$succCnt" },
    @{ l = 'IP loi / khong xac dinh'; v = "$errCnt" },
    @{ l = 'So quoc gia phat hien';   v = "$($cGroup.Count)" },
    @{ l = 'Thoi gian xu ly';         v = $elStr },
    @{ l = 'API su dung';             v = 'ip-api.com (batch, mien phi)' }
)
$sumRows | ForEach-Object { GetStrIdx $_.l | Out-Null; GetStrIdx $_.v | Out-Null }

$sb1 = [System.Text.StringBuilder]::new()
[void]$sb1.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
[void]$sb1.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
[void]$sb1.Append('<sheetFormatPr defaultRowHeight="24"/>')
[void]$sb1.Append('<cols>')
[void]$sb1.Append('<col min="1" max="1" width="34" customWidth="1"/>')
[void]$sb1.Append('<col min="2" max="2" width="32" customWidth="1"/>')
[void]$sb1.Append('</cols><sheetData>')
[void]$sb1.Append("<row r=`"1`" ht=`"32`">")
[void]$sb1.Append((xmlSC 1 1 'BAO CAO TONG QUAN - IP LOCATION' 2))
[void]$sb1.Append("</row>")
for ($i = 0; $i -lt $sumRows.Count; $i++) {
    $ri = $i + 2
    $sr = $sumRows[$i]
    [void]$sb1.Append("<row r=`"$ri`">")
    [void]$sb1.Append((xmlSC $ri 1 $sr.l 9))
    [void]$sb1.Append((xmlSC $ri 2 $sr.v 10))
    [void]$sb1.Append("</row>")
}
[void]$sb1.Append('</sheetData>')
[void]$sb1.Append('<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>')
[void]$sb1.Append('</worksheet>')
WriteFile (Join-Path $tmp 'xl\worksheets\sheet1.xml') $sb1.ToString()

# ===== Shared Strings =====
$sbSS = [System.Text.StringBuilder]::new()
[void]$sbSS.Append("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>")
[void]$sbSS.Append("<sst xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" count=`"$($allStr.Count)`" uniqueCount=`"$($allStr.Count)`">")
foreach ($s in $allStr) {
    $esc = $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
    [void]$sbSS.Append("<si><t xml:space=`"preserve`">$esc</t></si>")
}
[void]$sbSS.Append('</sst>')
WriteFile (Join-Path $tmp 'xl\sharedStrings.xml') $sbSS.ToString()

# ===== Zip -> .xlsx =====
if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }
[System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $OutputFile)
Remove-Item $tmp -Recurse -Force

Write-Host ""
Write-Host "  => Da luu: $OutputFile"
Write-Host "  Tong ket: $total IP  |  $succCnt thanh cong  |  $errCnt loi  |  $($cGroup.Count) quoc gia"
