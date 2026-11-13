$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$outFile = Join-Path $PSScriptRoot 'VentaSoft_ventas.xlsx'
if (Test-Path -LiteralPath $outFile) { Remove-Item -LiteralPath $outFile }
$vendors = @(
    @{ id='V001'; name='Ana Torres' },
    @{ id='V002'; name='Luis Mendoza' },
    @{ id='V003'; name='Carla Rojas' },
    @{ id='V004'; name='Diego Salazar' },
    @{ id='V005'; name='María Paredes' }
)
$products = @(
    @{ id='POS-01'; name='VentaSoft Punto de Venta'; price=249 },
    @{ id='INV-02'; name='Control de Inventario'; price=189 },
    @{ id='PAN-03'; name='Gestión para Panaderías'; price=279 },
    @{ id='HW-04'; name='Kit de Punto de Venta'; price=890 },
    @{ id='HW-05'; name='Lector de Código de Barras'; price=165 },
    @{ id='SRV-06'; name='Instalación y Capacitación'; price=120 }
)
$sales = @()
$start = [datetime]'2026-01-02'
for ($i=1; $i -le 100; $i++) {
    $seller = $vendors[($i * 3 + 1) % $vendors.Count]
    $product = $products[($i * 7 + 2) % $products.Count]
    $qty = ($i % 4) + 1
    $day = $start.AddDays(($i * 2) % 300)
    $sales += ,@("VTA-{0:D4}" -f $i, $day.ToString('yyyy-MM-dd'), $seller.id, $seller.name, $product.id, $product.name, [string]$qty, [string]$product.price, [string]($qty*$product.price), 'Sí')
}
function Escape-Xml([string]$s) { return [System.Security.SecurityElement]::Escape($s) }
function Cell-Xml([string]$value, [int]$row, [int]$col, [bool]$header=$false) {
    $letters=''; $n=$col
    while ($n -gt 0) { $r=($n-1)%26; $letters=[char](65+$r)+$letters; $n=[math]::Floor(($n-1)/26) }
    $ref="$letters$row"; $style=if($header){' s="1"'}else{''}
    if ($value -match '^\d+(\.\d+)?$') { return "<c r=`"$ref`"$style><v>$value</v></c>" }
    return "<c r=`"$ref`" t=`"inlineStr`"$style><is><t>$(Escape-Xml $value)</t></is></c>"
}
function Sheet-Xml($headers, $rows, $lastCol, $lastRow) {
    $body=New-Object System.Text.StringBuilder
    for($r=1;$r -le $lastRow;$r++){
        [void]$body.Append("<row r=`"$r`">")
        $vals=if($r -eq 1){$headers}else{$rows[$r-2]}
        for($c=1;$c -le $headers.Count;$c++){
            $isHeader=($r -eq 1); $v=[string]$vals[$c-1]
            [void]$body.Append((Cell-Xml $v $r $c $isHeader))
        }
        [void]$body.Append('</row>')
    }
    return "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><worksheet xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><dimension ref=`"A1:$lastCol$lastRow`"/><sheetViews><sheetView workbookViewId=`"0`"><pane ySplit=`"1`" topLeftCell=`"A2`" activePane=`"bottomLeft`" state=`"frozen`"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight=`"18`"/><sheetData>$body</sheetData><autoFilter ref=`"A1:$lastCol$lastRow`"/><tableParts count=`"1`"><tablePart r:id=`"rId1`"/></tableParts></worksheet>"
}
function Table-Xml($name, $headers, $lastCol, $lastRow, $id) {
    $cols=New-Object System.Text.StringBuilder
    for($i=0;$i -lt $headers.Count;$i++){[void]$cols.Append("<tableColumn id=`"$($i+1)`" name=`"$(Escape-Xml $headers[$i])`"/>")}
    return "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><table xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" id=`"$id`" name=`"$name`" displayName=`"$name`" ref=`"A1:$lastCol$lastRow`" totalsRowShown=`"0`"><autoFilter ref=`"A1:$lastCol$lastRow`"/><tableColumns count=`"$($headers.Count)`">$cols</tableColumns><tableStyleInfo name=`"TableStyleMedium4`" showFirstColumn=`"0`" showLastColumn=`"0`" showRowStripes=`"1`" showColumnStripes=`"0`"/></table>"
}
$headersSales=@('venta_id','fecha','vendedor_id','vendedor','producto_id','producto','cantidad','precio_unitario','total','es_demo')
$sellerRows=@($vendors | ForEach-Object { ,@($_.id,$_.name) })
$saleSheet=Sheet-Xml $headersSales $sales 'J' 101
$sellerSheet=Sheet-Xml @('vendedor_id','nombre') $sellerRows 'B' 6
$temp=Join-Path $env:TEMP ('ventasoft-xlsx-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null
foreach($dir in @('_rels','docProps','xl','xl/_rels','xl/worksheets','xl/worksheets/_rels','xl/tables')){New-Item -ItemType Directory -Path (Join-Path $temp $dir) -Force | Out-Null}
$files=@{
 '[Content_Types].xml'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
 '_rels/.rels'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
 'xl/workbook.xml'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="ventas" sheetId="1" r:id="rId1"/><sheet name="vendedor" sheetId="2" r:id="rId2"/></sheets><calcPr calcId="191029"/></workbook>'
 'xl/_rels/workbook.xml.rels'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
 'xl/styles.xml'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Aptos"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF10231C"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
 'xl/worksheets/sheet1.xml'=$saleSheet
 'xl/worksheets/sheet2.xml'=$sellerSheet
 'xl/worksheets/_rels/sheet1.xml.rels'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>'
 'xl/worksheets/_rels/sheet2.xml.rels'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/></Relationships>'
 'xl/tables/table1.xml'=(Table-Xml 'ventas' $headersSales 'J' 101 1)
 'xl/tables/table2.xml'=(Table-Xml 'vendedor' @('vendedor_id','nombre') 'B' 6 2)
}
$utf8=New-Object System.Text.UTF8Encoding($false)
foreach($path in $files.Keys){$dest=Join-Path $temp $path;[System.IO.File]::WriteAllText($dest,[string]$files[$path],$utf8)}
[System.IO.Compression.ZipFile]::CreateFromDirectory($temp,$outFile)
Remove-Item -LiteralPath $temp -Recurse -Force
Write-Output "Workbook created: $outFile | ventas: $($sales.Count) | vendedor: $($vendors.Count)"
