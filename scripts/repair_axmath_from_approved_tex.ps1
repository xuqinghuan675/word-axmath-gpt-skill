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
    $stream=[IO.File]::Open(
      $Path,
      [IO.FileMode]::Open,
      [IO.FileAccess]::Read,
      ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
    )
    $sha=[Security.Cryptography.SHA256]::Create()
    return ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-','').ToLowerInvariant()
  }finally{
    if($sha){$sha.Dispose()}
    if($stream){$stream.Dispose()}
  }
}


function Resolve-AxMathTemplate {
  param([string]$Requested)
  if(-not [string]::IsNullOrWhiteSpace($Requested)){
    if(Test-Path -LiteralPath $Requested){return (Resolve-Path -LiteralPath $Requested).Path}
    throw "AxMath.dotm not found: $Requested"
  }
  $pf86=[Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
  $pf=[Environment]::GetEnvironmentVariable('ProgramFiles')
  $candidates=@()
  if($pf86){$candidates += (Join-Path $pf86 'AxMath\MSOffice\AxMath.dotm')}
  if($pf){$candidates += (Join-Path $pf 'AxMath\MSOffice\AxMath.dotm')}
  foreach($candidate in $candidates){
    if(Test-Path -LiteralPath $candidate){return $candidate}
  }
  throw 'AxMath.dotm not found. Install/activate AxMath or pass -TemplatePath explicitly.'
}

function Get-AxMathShapes($d){
  $arr=@()
  $collection=$d.InlineShapes
  for($i=1;$i -le [int]$collection.Count;$i++){
    $s=$collection.Item($i)
    try{if([string]$s.OLEFormat.ProgID -eq 'Equation.AxMath'){$arr += $s}}catch{}
  }
  return $arr
}

function Get-MapTargetAxMath {
  param($Doc,$Row,$FallbackAx)
  $hasParagraph=($Row.PSObject.Properties.Name -contains 'working_paragraph_index')
  $hasLocal=($Row.PSObject.Properties.Name -contains 'working_paragraph_axmath_index')
  if($hasParagraph -and $hasLocal){
    $pi=[int]$Row.working_paragraph_index
    $li=[int]$Row.working_paragraph_axmath_index
    if($pi -lt 1 -or $pi -gt [int]$Doc.Paragraphs.Count){
      throw "working paragraph index out of range: $pi"
    }
    $paragraph=$Doc.Paragraphs.Item($pi)
    $locals=@(Get-AxMathShapes $paragraph.Range)
    if($li -lt 1 -or $li -gt $locals.Count){
      throw "paragraph $pi AxMath local index $li out of range 1..$($locals.Count)"
    }
    return $locals[$li-1]
  }
  $ord=[int]$Row.ordinal
  if($null -eq $FallbackAx){
    throw "map row $ord has no static paragraph locator and no fallback AxMath index"
  }
  if($ord -lt 1 -or $ord -gt $FallbackAx.Count){
    throw "ordinal $ord out of range 1..$($FallbackAx.Count)"
  }
  return $FallbackAx[$ord-1]
}

$InputFull=[IO.Path]::GetFullPath($InputDocx)
$OutputFull=[IO.Path]::GetFullPath($OutputDocx)
$MapFull=[IO.Path]::GetFullPath($MapPath)
if([string]::Equals($InputFull,$OutputFull,[StringComparison]::OrdinalIgnoreCase)){throw 'Refusing to overwrite the input DOCX.'}
if(-not(Test-Path -LiteralPath $InputFull)){throw "Input DOCX not found: $InputFull"}
if(-not(Test-Path -LiteralPath $MapFull)){throw "Approved TeX map not found: $MapFull"}
if((Test-Path -LiteralPath $OutputFull) -and -not $OverwriteOutput){throw "Output already exists: $OutputFull. Use -OverwriteOutput only for an intentional intermediate replacement."}
if([string]::IsNullOrWhiteSpace($ReportPath)){$ReportPath=$OutputFull+'.source-tex-repair.json'}
$ReportFull=[IO.Path]::GetFullPath($ReportPath)
if([string]::Equals($ReportFull,$InputFull,[StringComparison]::OrdinalIgnoreCase) -or [string]::Equals($ReportFull,$OutputFull,[StringComparison]::OrdinalIgnoreCase) -or [string]::Equals($ReportFull,$MapFull,[StringComparison]::OrdinalIgnoreCase)){throw 'ReportPath must not point to the input DOCX, output DOCX, or approved TeX map.'}

$mapData=Get-Content -Raw -Encoding UTF8 -LiteralPath $MapFull|ConvertFrom-Json
$mapMeta=$null
if($mapData.PSObject.Properties.Name -contains 'rows'){
  $mapMeta=$mapData
  $rows=@($mapData.rows)
}else{
  $rows=@($mapData)
}
if($rows.Count -eq 0){throw 'Approved TeX map is empty.'}
if($rows.Count -gt 48){throw "Approved source-semantic repair is intentionally bounded. Refusing $($rows.Count) rows; split reviewed repairs into batches of at most 48."}
$ordinals=@($rows|ForEach-Object{[int]$_.ordinal})
if(@($ordinals|Sort-Object -Unique).Count -ne $rows.Count){throw 'Approved TeX map contains duplicate ordinals.'}
foreach($row in $rows){
  $tex=[string]$row.tex
  if([string]::IsNullOrWhiteSpace($tex)){throw "Approved TeX is empty for ordinal $($row.ordinal)."}
  if(-not($tex.StartsWith('$') -and $tex.EndsWith('$')) -or $tex.StartsWith('$$') -or $tex.EndsWith('$$')){throw "Approved TeX must be wrapped in exactly one leading and trailing dollar delimiter for ordinal $($row.ordinal)."}
  Assert-CanonicalPrimeTeX -Tex $tex -Label ("ordinal "+[string]$row.ordinal)
  if([string]$row.repair_class -eq 'M2_SOURCE_VISUAL_WRAP_LOSS'){
    if([int]$row.expected_visual_lines -lt 2){throw "M2 row $($row.ordinal) must prove at least two source visual lines."}
    if($tex -notmatch '\\begin\{aligned\}'){throw "M2 row $($row.ordinal) must use one aligned multi-line donor, not multiple AxMath objects."}
  }
}

$inputSha=(Get-SharedSha256 $InputFull)
if($mapMeta -and $mapMeta.working_sha256){
  if(([string]$mapMeta.working_sha256).ToLowerInvariant() -ne $inputSha){
    throw 'Stale approved map: working_sha256 does not match the current input DOCX.'
  }
}

$TemplatePath=Resolve-AxMathTemplate $TemplatePath
$beforeWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$result=[ordered]@{
  schema='axmath-approved-tex-repair/v2'
  input=$InputFull
  input_sha256=$inputSha
  output=$OutputFull
  approved_tex_map=$MapFull
  repair_class=if($mapMeta){[string]$mapMeta.repair_class}else{'LEGACY_APPROVED_TEX'}
  source_sha256=if($mapMeta){[string]$mapMeta.source_sha256}else{$null}
  prime_contract='axmath_builtin_prime_v2'
  target_count=$rows.Count
  preexisting_word_pids=@($beforeWord)
  repairs=@()
  failed=@()
  success=$false
}
$word=$null;$doc=$null;$wordPid=$null;$owned=$false

try{
  if(Test-Path -LiteralPath $OutputFull){Remove-Item -LiteralPath $OutputFull -Force}
  Copy-Item -LiteralPath $InputFull -Destination $OutputFull
  $outputItem=Get-Item -LiteralPath $OutputFull
  $result.output_readonly_cleared=[bool]$outputItem.IsReadOnly
  if($outputItem.IsReadOnly){$outputItem.IsReadOnly=$false}

  $word=New-Object -ComObject Word.Application
  Start-Sleep -Milliseconds 300
  $newWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Where-Object{$beforeWord -notcontains $_.Id}|Select-Object -ExpandProperty Id)
  if($newWord.Count -eq 1){
    $wordPid=[int]$newWord[0]
    $owned=$true
    $result.word_pid=$wordPid
    $result.word_pid_owned=$true
  }
  if(-not $owned){throw 'Could not establish a distinct task-owned Word process; pre-existing user Word processes were left untouched.'}

  $word.Visible=$false
  $word.DisplayAlerts=0
  try{$word.ScreenUpdating=$false}catch{}
  try{$word.Options.SaveNormalPrompt=$false}catch{}

  $found=$false
  foreach($a in $word.AddIns){
    if($a.Name -eq 'AxMath.dotm'){
      $found=$true
      if(-not $a.Installed){$a.Installed=$true}
    }
  }
  if(-not $found){$word.AddIns.Add($TemplatePath,$true)|Out-Null}

  $doc=$word.Documents.Open($OutputFull,$false,$false,$false)
  $result.inline_shapes_before=[int]$doc.InlineShapes.Count
  $result.omath_before=[int]$doc.OMaths.Count
  $result.paragraphs_before=[int]$doc.Paragraphs.Count
  if($result.omath_before -ne 0){throw "Approved TeX repair requires a post-conversion candidate with OfficeMath=0; found $($result.omath_before)."}

  $allRowsHaveStaticLocator=$true
  foreach($mapRow in $rows){
    if(
      -not($mapRow.PSObject.Properties.Name -contains 'working_paragraph_index') -or
      -not($mapRow.PSObject.Properties.Name -contains 'working_paragraph_axmath_index')
    ){$allRowsHaveStaticLocator=$false;break}
  }

  $fallbackAx=$null
  if($mapMeta -and $mapMeta.working_axmath_count){
    $result.axmath_before=[int]$mapMeta.working_axmath_count
  }elseif($allRowsHaveStaticLocator){
    throw 'Located approved map is missing working_axmath_count.'
  }else{
    # Legacy maps pay one O(N) COM scan. Production maps should carry static
    # paragraph/local locators so no whole-document COM scan is needed.
    $fallbackAx=@(Get-AxMathShapes $doc)
    $result.axmath_before=$fallbackAx.Count
    $result.legacy_global_axmath_scan=$true
  }

  foreach($row in @($rows|Sort-Object{[int]$_.ordinal} -Descending)){
    $ord=[int]$row.ordinal
    $rec=[ordered]@{
      ordinal=$ord
      repair_class=[string]$row.repair_class
      success=$false
      tex=[string]$row.tex
      expected_visual_lines=$row.expected_visual_lines
    }
    $tmp=$null
    try{
      $target=Get-MapTargetAxMath $doc $row $fallbackAx
      $oldStart=[int]$target.Range.Start
      $oldEnd=[int]$target.Range.End
      $rec.before=[ordered]@{width=[double]$target.Width;height=[double]$target.Height;start=$oldStart;end=$oldEnd}

      $tmp=$word.Documents.Add()
      $tmp.Content.Text=[string]$row.tex
      $tmp.Content.Select()
      $cb=$null
      $word.Run('AMSTeX2AM',([ref]$cb))
      $donorAx=@(Get-AxMathShapes $tmp)
      if($donorAx.Count -ne 1){
        throw "approved TeX donor count=$($donorAx.Count); every M2/Class-E row must remain exactly one AxMath object"
      }
      $donor=$donorAx[0]
      $dw=[double]$donor.Width;$dh=[double]$donor.Height
      if($dw -le 4 -or $dh -le 8 -or $dw -gt 2000 -or $dh -gt 1000){
        throw "approved donor geometry is implausible/corrupt: $dw x $dh"
      }
      $rec.donor=[ordered]@{width=$dw;height=$dh}
      $primeCheck=Verify-AxMathPrimeDonor -Word $word -Donor $donor -ApprovedTex ([string]$row.tex) -Label ("ordinal "+$ord)
      $rec.prime_orders=@($primeCheck.prime_orders)
      $rec.prime_roundtrip_verified=[bool]$primeCheck.verified
      $rec.prime_roundtrip_tex=$primeCheck.roundtrip_tex

      $donor.Range.Copy()
      $doc.Range($oldStart,$oldEnd).Delete()
      $doc.Range($oldStart,$oldStart).Paste()
      $tmp.Close($false);$tmp=$null

      if(
        ($row.PSObject.Properties.Name -contains 'working_paragraph_index') -and
        ($row.PSObject.Properties.Name -contains 'working_paragraph_axmath_index')
      ){
        $replacement=Get-MapTargetAxMath $doc $row $null
      }else{
        # Legacy maps do not carry a static paragraph/local locator. They already
        # paid one O(N) COM scan before mutation; after replacement, do one fresh
        # scan and recover the same ordinal deterministically. Production maps
        # avoid this path by carrying working_paragraph_index/local_index.
        $afterLegacy=@(Get-AxMathShapes $doc)
        if($afterLegacy.Count -ne $result.axmath_before){
          throw "legacy replacement changed AxMath count: $($afterLegacy.Count) != $($result.axmath_before)"
        }
        $replacement=$afterLegacy[$ord-1]
      }
      try{$prog=[string]$replacement.OLEFormat.ProgID}catch{$prog=''}
      if($prog -ne 'Equation.AxMath'){throw "replacement ordinal $ord is not Equation.AxMath"}
      $rec.after=[ordered]@{width=[double]$replacement.Width;height=[double]$replacement.Height;start=[int]$replacement.Range.Start;end=[int]$replacement.Range.End}
      $rec.success=$true
      $result.repairs += [pscustomobject]$rec
      $doc.Save()
    }catch{
      if($tmp -ne $null){try{$tmp.Close($false)}catch{};$tmp=$null}
      $rec.error=$_.Exception.Message
      $result.failed += [pscustomobject]$rec
      throw
    }
  }

  $result.inline_shapes_after=[int]$doc.InlineShapes.Count
  $result.axmath_after=$result.axmath_before
  $result.omath_after=[int]$doc.OMaths.Count
  $result.paragraphs_after=[int]$doc.Paragraphs.Count
  $doc.Save()
  $result.success=(
    $result.failed.Count -eq 0 -and
    $result.repairs.Count -eq $rows.Count -and
    $result.inline_shapes_after -eq $result.inline_shapes_before -and
    $result.omath_after -eq 0 -and
    $result.paragraphs_after -eq $result.paragraphs_before
  )
  $result.final_static_audit_required=$true
}catch{
  $result.error=$_.Exception.Message
  $result.hresult=$_.Exception.HResult
}finally{
  if($doc -ne $null){try{$doc.Close($false)}catch{}}
  if($word -ne $null -and $owned){try{$word.Quit()}catch{$result.word_quit_error=$_.Exception.Message}}
  $doc=$null;$word=$null;$fallbackAx=$null;$afterLegacy=$null;$target=$null;$donorAx=$null;$donor=$null;$replacement=$null;$primeCheck=$null
  [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
  if($wordPid -and $owned){
    $deadline=(Get-Date).AddSeconds(6)
    while((Get-Date)-lt $deadline -and (Get-Process -Id $wordPid -ErrorAction SilentlyContinue)){Start-Sleep -Milliseconds 200}
    if(Get-Process -Id $wordPid -ErrorAction SilentlyContinue){
      Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue
      $result.forced_owned_word_cleanup=$true
    }
  }
  if(Test-Path -LiteralPath $OutputFull){$result.output_sha256=(Get-SharedSha256 $OutputFull)}
  $result.finished_at=(Get-Date).ToString('o')
  $result|ConvertTo-Json -Depth 10|Set-Content -LiteralPath $ReportFull -Encoding UTF8
  $result|ConvertTo-Json -Depth 7
}
if(-not $result.success){exit 1}
