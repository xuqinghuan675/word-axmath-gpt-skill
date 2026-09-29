param(
  [Parameter(Mandatory=$true)][string]$InputDocx,
  [Parameter(Mandatory=$true)][string]$OutputDocx,
  [Parameter(Mandatory=$true)][string]$MapPath,
  [string]$TemplatePath='',
  [string]$ReportPath='',
  [switch]$OverwriteOutput
)
$ErrorActionPreference='Stop'
$PrimeContractScript=Join-Path $PSScriptRoot 'axmath_prime_contract.ps1'
if(-not(Test-Path -LiteralPath $PrimeContractScript)){throw 'axmath_prime_contract.ps1 missing'}
. $PrimeContractScript

function Get-SharedSha256 {
  param([Parameter(Mandatory=$true)][string]$Path)
  $stream=$null;$sha=$null
  try{
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    $sha=[Security.Cryptography.SHA256]::Create()
    return ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-','').ToLowerInvariant()
  }finally{if($sha){$sha.Dispose()};if($stream){$stream.Dispose()}}
}
function Resolve-AxMathTemplate {
  param([string]$Requested)
  if($Requested){
    if(Test-Path -LiteralPath $Requested){return (Resolve-Path -LiteralPath $Requested).Path}
    throw "AxMath.dotm not found: $Requested"
  }
  $pf86=[Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
  $pf=[Environment]::GetEnvironmentVariable('ProgramFiles')
  foreach($base in @($pf86,$pf)){
    if($base){
      $c=Join-Path $base 'AxMath\MSOffice\AxMath.dotm'
      if(Test-Path -LiteralPath $c){return $c}
    }
  }
  throw 'AxMath.dotm not found.'
}
function Get-AxMathShapes {
  param($RangeOrDoc)
  $arr=@();$col=$RangeOrDoc.InlineShapes
  for($i=1;$i -le [int]$col.Count;$i++){
    $s=$col.Item($i)
    try{if([string]$s.OLEFormat.ProgID -eq 'Equation.AxMath'){$arr += $s}}catch{}
  }
  return $arr
}

$InputFull=[IO.Path]::GetFullPath($InputDocx)
$OutputFull=[IO.Path]::GetFullPath($OutputDocx)
$MapFull=[IO.Path]::GetFullPath($MapPath)
if([string]::Equals($InputFull,$OutputFull,[StringComparison]::OrdinalIgnoreCase)){throw 'Output must differ from input.'}
if(-not(Test-Path -LiteralPath $InputFull)){throw "Input missing: $InputFull"}
if(-not(Test-Path -LiteralPath $MapFull)){throw "Map missing: $MapFull"}
if((Test-Path -LiteralPath $OutputFull) -and -not $OverwriteOutput){throw "Output exists: $OutputFull"}
if(-not $ReportPath){$ReportPath=$OutputFull+'.visual-line-split.json'}
$ReportFull=[IO.Path]::GetFullPath($ReportPath)

$map=Get-Content -LiteralPath $MapFull -Raw -Encoding UTF8|ConvertFrom-Json
if([string]$map.repair_class -ne 'M2_SOURCE_VISUAL_LINE_SPLIT'){throw 'Map is not M2_SOURCE_VISUAL_LINE_SPLIT.'}
$rows=@($map.rows)
if($rows.Count -eq 0){throw 'Split map is empty.'}
$inputSha=Get-SharedSha256 $InputFull
if(([string]$map.working_sha256).ToLowerInvariant() -ne $inputSha){throw 'Stale split map: working SHA mismatch.'}
$expectedIncrease=[int]$map.expected_axmath_increase
$expectedFinal=[int]$map.expected_final_axmath_count
$firstRepair=[int]$map.first_repair_paragraph
$TemplatePath=Resolve-AxMathTemplate $TemplatePath

$before=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$res=[ordered]@{
 schema='axmath-m2-visual-line-split/v1';success=$false;input=$InputFull;input_sha256=$inputSha;
 output=$OutputFull;map=$MapFull;first_repair_paragraph=$firstRepair;prime_contract='axmath_builtin_prime_v2';rows=@();failed=@();preexisting_word_pids=@($before)
}
$word=$null;$doc=$null;$tmp=$null;$wordPid=$null;$owned=$false
try{
  if(Test-Path -LiteralPath $OutputFull){Remove-Item -LiteralPath $OutputFull -Force}
  Copy-Item -LiteralPath $InputFull -Destination $OutputFull
  $word=New-Object -ComObject Word.Application
  Start-Sleep -Milliseconds 300
  $new=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Where-Object{$before -notcontains $_.Id}|Select-Object -ExpandProperty Id)
  if($new.Count -ne 1){
    $word=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
    throw "Could not prove distinct task-owned Word PID; created=$($new -join ',')"
  }
  $wordPid=[int]$new[0];$owned=$true;$res.word_pid=$wordPid;$res.word_pid_owned=$true
  $word.Visible=$false;$word.DisplayAlerts=0
  try{$word.ScreenUpdating=$false}catch{}
  $found=$false
  foreach($a in $word.AddIns){if($a.Name -eq 'AxMath.dotm'){$found=$true;if(-not $a.Installed){$a.Installed=$true}}}
  if(-not $found){$word.AddIns.Add($TemplatePath,$true)|Out-Null}

  $doc=$word.Documents.Open($OutputFull,$false,$false,$false)
  $res.paragraphs_before=[int]$doc.Paragraphs.Count
  $res.omath_before=[int]$doc.OMaths.Count
  $res.inline_shapes_before=[int]$doc.InlineShapes.Count
  if($res.omath_before -ne 0){throw "Split repair requires post-conversion input with OfficeMath=0; found $($res.omath_before)"}
  if([int]$map.working_paragraph_count -ne $res.paragraphs_before){throw 'Paragraph-count contract mismatch before repair.'}

  foreach($row in @($rows|Sort-Object {[int]$_.working_paragraph_index} -Descending)){
    $pi=[int]$row.working_paragraph_index
    $rec=[ordered]@{paragraph=$pi;source_paragraph=[int]$row.source_paragraph_index;source_ordinal=[int]$row.source_ordinal;success=$false}
    try{
      if($pi -lt $firstRepair){throw "Target paragraph $pi crosses frozen prefix boundary $firstRepair"}
      $pr=$doc.Paragraphs.Item($pi)
      $existing=@(Get-AxMathShapes $pr.Range)
      if($existing.Count -ne [int]$row.expected_current_axmath){throw "paragraph $pi AxMath=$($existing.Count), expected $($row.expected_current_axmath)"}
      if($existing.Count -ne 1){throw "paragraph $pi must contain exactly one pre-repair AxMath"}
      $start=[int]$existing[0].Range.Start
      $existing[0].Range.Delete()
      $pos=$start;$lineNo=0;$sizes=@()
      foreach($line in @($row.lines)){
        $lineNo++
        $tex=[string]$line.tex
        if(-not($tex.StartsWith('$') -and $tex.EndsWith('$'))){throw "paragraph $pi line $lineNo invalid TeX delimiters"}
        Assert-CanonicalPrimeTeX -Tex $tex -Label ("paragraph "+$pi+" line "+$lineNo)
        $tmp=$word.Documents.Add()
        $tmp.Content.Text=$tex
        $tmp.Content.Select()
        $cb=$null;$word.Run('AMSTeX2AM',([ref]$cb))
        $donors=@(Get-AxMathShapes $tmp.Content)
        if($donors.Count -ne 1){throw "paragraph $pi line $lineNo donor AxMath=$($donors.Count)"}
        $shape=$donors[0]
        $w=[double]$shape.Width;$h=[double]$shape.Height
        if($w -le 4 -or $h -le 8 -or $w -gt 2000 -or $h -gt 1000){throw "paragraph $pi line $lineNo implausible donor geometry $w x $h"}
        if($lineNo -gt 1){$doc.Range($pos,$pos).InsertAfter([string][char]11);$pos++}
        $dest=$doc.Range($pos,$pos);$dest.FormattedText=$shape.Range.FormattedText;$pos=[int]$dest.End
        $primeCheck=Verify-AxMathPrimeDonor -Word $word -Donor $shape -ApprovedTex $tex -Label ("paragraph "+$pi+" line "+$lineNo)
        $sizes += [pscustomobject]@{
          line=$lineNo;width=$w;height=$h;provenance=[string]$line.provenance
          prime_orders=@($primeCheck.prime_orders)
          prime_roundtrip_verified=[bool]$primeCheck.verified
          prime_roundtrip_tex=$primeCheck.roundtrip_tex
        }
        $tmp.Close($false);$tmp=$null
      }
      $pr=$doc.Paragraphs.Item($pi);$after=@(Get-AxMathShapes $pr.Range)
      $txt=[string]$pr.Range.Text;$soft=0;foreach($ch in $txt.ToCharArray()){if([int][char]$ch -eq 11){$soft++}}
      if($after.Count -ne [int]$row.expected_after_axmath){throw "paragraph $pi post AxMath=$($after.Count), expected $($row.expected_after_axmath)"}
      if($soft -lt [int]$row.expected_soft_breaks_added){throw "paragraph $pi soft breaks=$soft, expected at least $($row.expected_soft_breaks_added)"}
      if([int]$pr.Range.OMaths.Count -ne 0){throw "paragraph $pi unexpectedly contains OfficeMath"}
      $rec.axmath_after=$after.Count;$rec.soft_breaks=$soft;$rec.sizes=$sizes;$rec.success=$true
      $res.rows += [pscustomobject]$rec
      $doc.Save()
    }catch{
      if($tmp){try{$tmp.Close($false)}catch{};$tmp=$null}
      $rec.error=$_.Exception.Message;$res.failed += [pscustomobject]$rec
      throw
    }
  }

  $res.paragraphs_after=[int]$doc.Paragraphs.Count
  $res.omath_after=[int]$doc.OMaths.Count
  $res.inline_shapes_after=[int]$doc.InlineShapes.Count
  $res.expected_axmath_increase=$expectedIncrease
  $res.expected_final_axmath=$expectedFinal
  $doc.Save()
  $res.success=(
    $res.failed.Count -eq 0 -and
    $res.rows.Count -eq $rows.Count -and
    $res.paragraphs_after -eq $res.paragraphs_before -and
    $res.omath_after -eq 0 -and
    $res.inline_shapes_after -eq ($res.inline_shapes_before+$expectedIncrease)
  )
}catch{$res.error=$_.Exception.Message;$res.hresult=$_.Exception.HResult}
finally{
  if($tmp){try{$tmp.Close($false)}catch{}}
  if($doc){try{$doc.Close($false)}catch{}}
  if($word -and $owned){try{$word.Quit()}catch{}}
  $tmp=$null;$doc=$null;$word=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
  if($wordPid -and $owned){
    $deadline=(Get-Date).AddSeconds(6)
    while((Get-Date)-lt $deadline -and (Get-Process -Id $wordPid -ErrorAction SilentlyContinue)){Start-Sleep -Milliseconds 200}
    if(Get-Process -Id $wordPid -ErrorAction SilentlyContinue){Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue;$res.forced_owned_word_cleanup=$true}
  }
  if(Test-Path -LiteralPath $OutputFull){$res.output_sha256=Get-SharedSha256 $OutputFull}
  $res.finished_at=(Get-Date).ToString('o')
  $res|ConvertTo-Json -Depth 10|Set-Content -LiteralPath $ReportFull -Encoding UTF8
  $res|ConvertTo-Json -Depth 6
}
if(-not $res.success){exit 1}
