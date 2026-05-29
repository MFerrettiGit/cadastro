<#
build_cadastro.ps1 — Gera produtos.js para o site Cadastro de Produtos (M. Ferretti).

Modo HIBRIDO:
  - Campos do ERP vem do export da query Protheus (query_cadastro_produtos.sql).
  - Campos que so existem na ficha (.xls) sao complementados casando pelo CODIGO.

Uso:
  powershell -ExecutionPolicy Bypass -File build_cadastro.ps1 `
     -QueryExport "C:\...\export_query.xlsx" `
     -Ficha "C:\...\FICHA CADASTRO M FERRETTI - 2026.05.26.xls" `
     -Output "C:\Users\COMPRASD\cadastro-site\produtos.js"

  -QueryExport pode ser omitido: nesse caso o produtos.js e montado SO da ficha.
#>
param(
  [string]$QueryExport,
  [Parameter(Mandatory=$true)][string]$Ficha,
  [string]$Output = "C:\Users\COMPRASD\cadastro-site\produtos.js"
)
$ErrorActionPreference = "Stop"

# ---- NPOI (reaproveita libs da skill atualiza-cadastro) ----
$lib = "C:\Users\COMPRASD\.claude\skills\atualiza-cadastro\scripts\lib"
Get-ChildItem $lib -Filter *.dll | ForEach-Object { Add-Type -Path $_.FullName }

# Colunas da ficha (0-based) — ver interpretation.md
$F = [ordered]@{
  FORNECEDOR=0; MARCA=1; REFERENCIA=2; CODIGO=3; DESCRICAO=4;
  CAIXA=5; LOTE_MINIMO=6; NCM=7; CEST=8; CFOP=9; CST=10;
  IVA_ENTRADA=11; IVA_SAIDA=12; ICMS=13; PIS_COFINS=14; VALIDADE=15;
  EAN=16; DUN=17;
  ITEM_PESO_BRUTO=18; ITEM_PESO_LIQ=19; ITEM_UN_PESO=20;
  ITEM_A=21; ITEM_C=22; ITEM_L=23; ITEM_UN_DIM=24;
  CAIXA_PESO_BRUTO=25; CAIXA_PESO_LIQ=26; CAIXA_UN_PESO=27;
  CAIXA_A=28; CAIXA_C=29; CAIXA_L=30; CAIXA_UN_DIM=31;
  PALETE_LASTRO=32; PALETE_ALT=33
}
$HEADER_ROW = 4   # indice 0-based da linha de cabecalho real
$DATA_START = 5

function CellStr($cell){
  if($null -eq $cell){ return "" }
  switch($cell.CellType){
    ([NPOI.SS.UserModel.CellType]::String)  { return $cell.StringCellValue.Trim() }
    ([NPOI.SS.UserModel.CellType]::Numeric) {
      $d=$cell.NumericCellValue
      if([math]::Floor($d) -eq $d){ return ([long]$d).ToString() }
      return $d.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }
    ([NPOI.SS.UserModel.CellType]::Boolean) { return $cell.BooleanCellValue.ToString() }
    ([NPOI.SS.UserModel.CellType]::Formula) {
      try { return $cell.StringCellValue.Trim() } catch { try { return $cell.NumericCellValue.ToString() } catch { return "" } }
    }
    default { return "" }
  }
}

# Formata valor tributario como percentual.
# Regra (interpretation.md): valor <= 1 e fracao (0.373 -> 37,3%); > 1 ja e percentual (18 -> 18%).
function FmtPct($s){
  if([string]::IsNullOrWhiteSpace("$s")){ return '' }
  $n=0.0
  if([double]::TryParse((("$s" -replace ',','.')),[System.Globalization.NumberStyles]::Any,[System.Globalization.CultureInfo]::InvariantCulture,[ref]$n)){
    if($n -eq 0){ return '' }
    if($n -le 1){ $n = $n*100 }
    return (($n.ToString('0.##',[System.Globalization.CultureInfo]::InvariantCulture)) -replace '\.',',') + '%'
  }
  return "$s"
}

# ---- 1) Le a FICHA (.xls HSSF) -> hash por CODIGO com os 34 campos ----
Write-Host "Lendo ficha: $Ficha"
$fs = [System.IO.File]::OpenRead($Ficha)
$wb = New-Object NPOI.HSSF.UserModel.HSSFWorkbook($fs)
$sheet = $wb.GetSheetAt(0); $fs.Close()

$fichaByCod = @{}
for($r=$DATA_START; $r -le $sheet.LastRowNum; $r++){
  $row = $sheet.GetRow($r); if($null -eq $row){ continue }
  $cod = (CellStr $row.GetCell($F['CODIGO']))
  if([string]::IsNullOrWhiteSpace($cod)){ continue }
  $obj=@{}
  foreach($k in $F.Keys){ $obj[$k] = (CellStr $row.GetCell($F[$k])) }
  $fichaByCod[$cod] = $obj
}
Write-Host ("  -> {0} produtos na ficha" -f $fichaByCod.Count)

# ---- 2) Le o EXPORT DA QUERY (.xlsx) -> lista por CODIGO ----
function Read-Xlsx($path){
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip=[System.IO.Compression.ZipFile]::OpenRead($path)
  function GetE($z,$n){ $e=$z.Entries|?{$_.FullName -eq $n}; if(-not $e){return $null}; $s=New-Object System.IO.StreamReader($e.Open());$t=$s.ReadToEnd();$s.Close();return $t }
  $shared=GetE $zip "xl/sharedStrings.xml"
  $sheetName = ($zip.Entries|?{$_.FullName -match '^xl/worksheets/sheet1\.xml$'}|Select -First 1).FullName
  $sheetXml=GetE $zip $sheetName
  $zip.Dispose()
  $ss=@()
  if($shared){ foreach($m in [regex]::Matches($shared,'<si>(.*?)</si>')){ $tx=[regex]::Matches($m.Groups[1].Value,'<t[^>]*>(.*?)</t>')|%{$_.Groups[1].Value}; $ss += (($tx -join '') | %{ $_ -replace '&amp;','&' -replace '&lt;','<' -replace '&gt;','>' -replace '&quot;','"' -replace '&apos;',"'" }) } }
  function Decode($v){ $v -replace '&amp;','&' -replace '&lt;','<' -replace '&gt;','>' -replace '&quot;','"' -replace '&apos;',"'" }
  $rows=@()
  foreach($rm in [regex]::Matches($sheetXml,'<row[^>]*>(.*?)</row>')){
    $cells=@{}
    foreach($cm in [regex]::Matches($rm.Groups[1].Value,'<c r="([A-Z]+)\d+"(?:[^>]*?\st="([^"]+)")?[^>]*>(?:<v>(.*?)</v>|<is><t[^>]*>(.*?)</t></is>)?')){
      $col=$cm.Groups[1].Value; $t=$cm.Groups[2].Value; $v=$cm.Groups[3].Value; $iv=$cm.Groups[4].Value
      if($t -eq 's' -and $v -ne ''){ $val=$ss[[int]$v] }
      elseif($t -eq 'inlineStr'){ $val=Decode $iv }
      else { $val=Decode $v }
      $cells[$col]=$val
    }
    $rows += ,$cells
  }
  return $rows
}
function ColLetter($i){ $s=''; $i++; while($i -gt 0){ $m=($i-1)%26; $s=[char](65+$m)+$s; $i=[int][math]::Floor(($i-1)/26) }; return $s }

$erpByCod=@{}
if($QueryExport){
  Write-Host "Lendo export da query: $QueryExport"
  $rows = Read-Xlsx $QueryExport
  if($rows.Count -lt 2){ throw "Export vazio ou sem dados." }
  # cabecalho
  $hdr=@{}; $h=$rows[0]
  for($c=0;$c -lt 60;$c++){ $L=ColLetter $c; if($h.ContainsKey($L) -and $h[$L]){ $hdr[$h[$L].Trim().ToUpper()]=$L } }
  for($i=1;$i -lt $rows.Count;$i++){
    $row=$rows[$i]
    function GV($name){ if($hdr.ContainsKey($name)){ $L=$hdr[$name]; if($row.ContainsKey($L)){ return $row[$L] } }; return "" }
    $cod=(GV 'CODIGO')
    if([string]::IsNullOrWhiteSpace($cod)){ continue }
    $erpByCod[$cod]=@{
      CODIGO=$cod; DESCRICAO=(GV 'DESCRICAO'); MARCA=(GV 'MARCA'); FORNECEDOR=(GV 'FORNECEDOR');
      REFERENCIA=(GV 'REFERENCIA'); EAN=(GV 'EAN'); GTIN_NFE=(GV 'GTIN_NFE'); CAIXA=(GV 'CAIXA');
      LOTE_MINIMO=(GV 'LOTE_MINIMO'); UNIDADE=(GV 'UNIDADE'); NCM=(GV 'NCM'); CEST=(GV 'CEST');
      ORIGEM=(GV 'ORIGEM'); SITUACAO_TRIB=(GV 'SITUACAO_TRIB'); ICMS=(GV 'ICMS'); IPI=(GV 'IPI');
      PIS=(GV 'PIS'); COFINS=(GV 'COFINS'); VALIDADE_DIAS=(GV 'VALIDADE_DIAS');
      ITEM_PESO_BRUTO=(GV 'ITEM_PESO_BRUTO'); ITEM_PESO_LIQ=(GV 'ITEM_PESO_LIQ')
    }
  }
  Write-Host ("  -> {0} produtos no export do ERP" -f $erpByCod.Count)
}

# ---- 3) Merge -> objetos finais ----
# Chave-mestra: se houver export, usa os codigos do ERP; senao usa a ficha.
$codigos = if($erpByCod.Count){ $erpByCod.Keys } else { $fichaByCod.Keys }
$semFicha=@(); $list=@()
foreach($cod in $codigos){
  $e = $erpByCod[$cod]; $f = $fichaByCod[$cod]
  if($null -eq $f){ $semFicha += $cod }
  function pick($erpVal,$fichaKey){
    if($erpVal -and "$erpVal".Trim() -ne ''){ return "$erpVal".Trim() }
    if($f -and $f[$fichaKey]){ return $f[$fichaKey] }
    return ""
  }
  $o=[ordered]@{
    CODIGO     = $cod
    DESCRICAO  = (pick $e.DESCRICAO 'DESCRICAO')
    MARCA      = (pick $e.MARCA 'MARCA')
    FORNECEDOR = (pick $e.FORNECEDOR 'FORNECEDOR')
    REFERENCIA = (pick $e.REFERENCIA 'REFERENCIA')
    EAN        = (pick $e.EAN 'EAN')
    DUN        = if($f){ $f['DUN'] } else { '' }            # so na ficha
    CAIXA      = (pick $e.CAIXA 'CAIXA')
    LOTE_MINIMO= (pick $e.LOTE_MINIMO 'LOTE_MINIMO')
    NCM        = (pick $e.NCM 'NCM')
    CEST       = (pick $e.CEST 'CEST')
    CFOP       = if($f){ $f['CFOP'] } else { '' }           # so na ficha
    CST        = if($f -and $f['CST']){ $f['CST'] } elseif($e){ $e.SITUACAO_TRIB } else { '' }
    ICMS       = (FmtPct (pick $e.ICMS 'ICMS'))
    IPI        = if($e){ (FmtPct $e.IPI) } else { '' }
    ORIGEM     = if($e){ $e.ORIGEM } else { '' }
    PIS_COFINS = if($e -and ($e.PIS -or $e.COFINS)){ ("PIS "+(FmtPct $e.PIS)+" / COFINS "+(FmtPct $e.COFINS)) } elseif($f){ (FmtPct $f['PIS_COFINS']) } else { '' }
    IVA_ENTRADA= if($f){ (FmtPct $f['IVA_ENTRADA']) } else { '' }    # so na ficha
    IVA_SAIDA  = if($f){ (FmtPct $f['IVA_SAIDA']) } else { '' }      # so na ficha
    VALIDADE   = if($f -and $f['VALIDADE']){ $f['VALIDADE'] } elseif($e -and $e.VALIDADE_DIAS){ $e.VALIDADE_DIAS+" dias" } else { '' }
    ITEM_PESO_BRUTO=(pick $e.ITEM_PESO_BRUTO 'ITEM_PESO_BRUTO')
    ITEM_PESO_LIQ  =(pick $e.ITEM_PESO_LIQ 'ITEM_PESO_LIQ')
    ITEM_UN_PESO   = if($f -and $f['ITEM_UN_PESO']){ $f['ITEM_UN_PESO'] } else { 'kg' }
    ITEM_A = if($f){ $f['ITEM_A'] } else { '' }
    ITEM_C = if($f){ $f['ITEM_C'] } else { '' }
    ITEM_L = if($f){ $f['ITEM_L'] } else { '' }
    ITEM_UN_DIM = if($f -and $f['ITEM_UN_DIM']){ $f['ITEM_UN_DIM'] } else { 'cm' }
    CAIXA_PESO_BRUTO = if($f){ $f['CAIXA_PESO_BRUTO'] } else { '' }
    CAIXA_PESO_LIQ   = if($f){ $f['CAIXA_PESO_LIQ'] } else { '' }
    CAIXA_UN_PESO    = if($f -and $f['CAIXA_UN_PESO']){ $f['CAIXA_UN_PESO'] } else { 'kg' }
    CAIXA_A = if($f){ $f['CAIXA_A'] } else { '' }
    CAIXA_C = if($f){ $f['CAIXA_C'] } else { '' }
    CAIXA_L = if($f){ $f['CAIXA_L'] } else { '' }
    CAIXA_UN_DIM = if($f -and $f['CAIXA_UN_DIM']){ $f['CAIXA_UN_DIM'] } else { 'cm' }
    PALETE_LASTRO = if($f){ $f['PALETE_LASTRO'] } else { '' }
    PALETE_ALT    = if($f){ $f['PALETE_ALT'] } else { '' }
  }
  $list += ,$o
}

# ordena por MARCA, DESCRICAO
$list = $list | Sort-Object @{e={$_.MARCA}}, @{e={$_.DESCRICAO}}

$json = $list | ConvertTo-Json -Depth 4
$js = "// produtos.js - gerado por build_cadastro.ps1 em $(Get-Date -Format 'yyyy-MM-dd HH:mm')`r`n" +
      "// Fonte: query Protheus (ERP) + ficha .xls (campos complementares).`r`n" +
      "window.PRODUTOS = $json;`r`n"
$js | Out-File $Output -Encoding utf8

Write-Host ""
Write-Host ("OK -> {0} ({1} produtos)" -f $Output, $list.Count)
if($semFicha.Count){
  Write-Host ""
  Write-Host ("ATENCAO: {0} produto(s) do ERP SEM correspondencia na ficha (campos da ficha ficarao em branco):" -f $semFicha.Count) -ForegroundColor Yellow
  ($semFicha | Select-Object -First 30) | ForEach-Object { Write-Host "  - $_" }
  if($semFicha.Count -gt 30){ Write-Host ("  ... e mais {0}" -f ($semFicha.Count-30)) }
}
