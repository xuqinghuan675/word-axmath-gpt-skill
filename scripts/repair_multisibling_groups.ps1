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
  foreach($candidate in $candidates){if(Test-Path -LiteralPath $candidate){return $candidate}}
  throw 'AxMath.dotm not found.'
}

function Get-AxMathShapes {
  param($RangeOrDoc)
  $arr=@()
  $collection=$RangeOrDoc.InlineShapes
  for($i=1;$i -le [int]$collection.Count;$i++){
    $s=$collection.Item($i)
    try{if([string]$s.OLEFormat.ProgID -eq 'Equation.AxMath'){$arr += $s}}catch{}
  }
  return $arr
}

$InputFull=[IO.Path]::GetFullPath($InputDocx)
$OutputFull=[IO.Path]::GetFullPath($OutputDocx)
$MapFull=[IO.Path]::GetFullPath($MapPath)
if([string]::Equals($InputFull,$OutputFull,[StringComparison]::OrdinalIgnoreCase)){throw 'Refusing to overwrite the input DOCX.'}
if(-not(Test-Path -LiteralPath $InputFull)){throw "Input DOCX not found: $InputFull"}
if(-not(Test-Path -LiteralPath $MapFull)){throw "Repair map not found: $MapFull"}
if((Test-Path -LiteralPath $OutputFull) -and -not $OverwriteOutput){throw "Output already exists: $OutputFull"}
if([string]::IsNullOrWhiteSpace($ReportPath)){$ReportPath=$OutputFull+'.multisibling-repair.json'}
$ReportFull=[IO.Path]::GetFullPath($ReportPath)

$map=Get-Content -LiteralPath $MapFull -Raw -Encoding UTF8|ConvertFrom-Json
if([string]$map.repair_class -ne 'M1_MULTISIBLING_OMATHPARA'){throw 'Map is not an M1 multisibling repair map.'}
$rows=@($map.rows)
if($rows.Count -eq 0){throw 'M1 repair map has no rows.'}
$inputSha=(Get-SharedSha256 $InputFull)
if($map.working_sha256 -and ([string]$map.working_sha256).ToLowerInvariant() -ne $inputSha){
  throw "Stale M1 map: working SHA does not match input."
}
$expectedFinal=[int]$map.expected_final_axmath_count
$expectedIncrease=0
foreach($row in $rows){
  $lineCount=@($row.lines).Count
  if($lineCount -lt 2){throw "M1 row paragraph $($row.paragraph_index) has fewer than two source lines."}
  $expectedIncrease += ($lineCount-1)
}

$TemplatePath=Resolve-AxMathTemplate $TemplatePath
$beforeWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$result=[ordered]@{
  schema='axmath-m1-multisibling-repair/v1'
  input=$InputFull
  input_sha256=$inputSha
  output=$OutputFull
  map=$MapFull
  target_group_count=$rows.Count
  expected_increase=$expectedIncrease
  expected_final_axmath=$expectedFinal
  prime_contract='axmath_builtin_prime_v2'
  preexisting_word_pids=@($beforeWord)
  repairs=@()
  failed=@()
  success=$false
}
$word=$null;$doc=$null;$tmp=$null;$wordPid=$null;$owned=$false

try{
  if(Test-Path -LiteralPath $OutputFull){Remove-Item -LiteralPath $OutputFull -Force}
  Copy-Item -LiteralPath $InputFull -Destination $OutputFull
  $outItem=Get-Item -LiteralPath $OutputFull
  if($outItem.IsReadOnly){$outItem.IsReadOnly=$false}

  $word=New-Object -ComObject Word.Application
  Start-Sleep -Milliseconds 300
  $new=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Where-Object{$beforeWord -notcontains $_.Id}|Select-Object -ExpandProperty Id)
  if($new.Count -eq 1){$wordPid=[int]$new[0];$owned=$true;$result.word_pid=$wordPid;$result.word_pid_owned=$true}
  if(-not $owned){throw 'Could not establish a distinct task-owned Word process.'}
  $word.Visible=$false
  $word.DisplayAlerts=0
  try{$word.ScreenUpdating=$false}catch{}
  try{$word.Options.SaveNormalPrompt=$false}catch{}

  $found=$false
  foreach($a in $word.AddIns){if($a.Name -eq 'AxMath.dotm'){$found=$true;if(-not $a.Installed){$a.Installed=$true}}}
  if(-not $found){$word.AddIns.Add($TemplatePath,$true)|Out-Null}

  $doc=$word.Documents.Open($OutputFull,$false,$false,$false)
  $result.inline_shapes_before=[int]$doc.InlineShapes.Count
  if($map.working_axmath_count){
    $result.axmath_before=[int]$map.working_axmath_count
  }else{
    # Backward-compatible maps are SHA-bound; derive the only count compatible
    # with this reviewed partial M1 batch instead of scanning every OLE via COM.
    $result.axmath_before=$expectedFinal-$expectedIncrease
    $result.axmath_before_derived_from_map=$true
  }
  $result.omath_before=[int]$doc.OMaths.Count
  $result.paragraphs_before=[int]$doc.Paragraphs.Count
  if($result.omath_before -ne 0){throw "M1 input still contains OfficeMath: $($result.omath_before)"}
  if(($result.axmath_before+$expectedIncrease) -ne $expectedFinal){
    throw "M1 count contract mismatch: before=$($result.axmath_before), increase=$expectedIncrease, expected_final=$expectedFinal"
  }

  foreach($row in @($rows|Sort-Object {[int]$_.paragraph_index} -Descending)){
    $pi=[int]$row.paragraph_index
    $rec=[ordered]@{paragraph_index=$pi;success=$false;source_ordinals=@($row.source_ordinals)}
    try{
      $paragraph=$doc.Paragraphs.Item($pi)
      $existing=@(Get-AxMathShapes $paragraph.Range)
      if($existing.Count -ne [int]$row.expected_current_axmath){
        throw "paragraph $pi current AxMath count=$($existing.Count), expected=$($row.expected_current_axmath)"
      }
      if($existing.Count -ne 1){throw "M1 paragraph $pi is not in the expected collapsed-one-object state."}

      $start=[int]$existing[0].Range.Start
      $existing[0].Range.Delete()
      $pos=$start
      $rec.lines=@()
      $lineNo=0
      foreach($line in @($row.lines)){
        $lineNo++
        $tex=[string]$line.tex
        if([string]::IsNullOrWhiteSpace($tex) -or -not($tex.StartsWith('$') -and $tex.EndsWith('$'))){
          throw "paragraph $pi line $lineNo has invalid approved TeX delimiters"
        }
        Assert-CanonicalPrimeTeX -Tex $tex -Label ("paragraph "+$pi+" line "+$lineNo)
        $tmp=$word.Documents.Add()
        $tmp.Content.Text=$tex
        $tmp.Content.Select()
        $cb=$null
        $word.Run('AMSTeX2AM',([ref]$cb))
        $donors=@(Get-AxMathShapes $tmp)
        if($donors.Count -ne 1){throw "paragraph $pi line $lineNo donor AxMath count=$($donors.Count)"}
        $donor=$donors[0]
        $dw=[double]$donor.Width;$dh=[double]$donor.Height
        if($dw -le 4 -or $dh -le 8 -or $dw -gt 2000 -or $dh -gt 1000){
          throw "paragraph $pi line $lineNo donor geometry is corrupt: $dw x $dh"
        }
        if($lineNo -gt 1){
          $doc.Range($pos,$pos).InsertAfter([string][char]11)
          $pos++
        }
        $dest=$doc.Range($pos,$pos)
        $dest.FormattedText=$donor.Range.FormattedText
        $pos=[int]$dest.End
        $primeCheck=Verify-AxMathPrimeDonor -Word $word -Donor $donor -ApprovedTex $tex -Label ("paragraph "+$pi+" line "+$lineNo)
        $rec.lines += [pscustomobject]@{
          source_ordinal=[int]$line.source_ordinal
          width=$dw
          height=$dh
          provenance=[string]$line.provenance
          prime_orders=@($primeCheck.prime_orders)
          prime_roundtrip_verified=[bool]$primeCheck.verified
          prime_roundtrip_tex=$primeCheck.roundtrip_tex
        }
        $tmp.Close($false);$tmp=$null
      }

      $paragraph=$doc.Paragraphs.Item($pi)
      $after=@(Get-AxMathShapes $paragraph.Range)
      if($after.Count -ne [int]$row.expected_after_axmath){
        throw "paragraph $pi post AxMath count=$($after.Count), expected=$($row.expected_after_axmath)"
      }
      if([int]$paragraph.Range.OMaths.Count -ne 0){throw "paragraph $pi unexpectedly contains OfficeMath after M1 repair"}
      $rec.axmath_after=$after.Count
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
  $result.axmath_after=$result.axmath_before+$expectedIncrease
  $result.omath_after=[int]$doc.OMaths.Count
  $result.paragraphs_after=[int]$doc.Paragraphs.Count
  $doc.Save()
  $result.success=(
    $result.failed.Count -eq 0 -and
    $result.repairs.Count -eq $rows.Count -and
    $result.axmath_after -eq $expectedFinal -and
    $result.inline_shapes_after -eq ($result.inline_shapes_before+$expectedIncrease) -and
    $result.omath_after -eq 0 -and
    $result.paragraphs_after -eq $result.paragraphs_before
  )
  $result.final_static_audit_required=$true
}catch{
  $result.error=$_.Exception.Message
  $result.hresult=$_.Exception.HResult
}finally{
  if($tmp -ne $null){try{$tmp.Close($false)}catch{}}
  if($doc -ne $null){try{$doc.Close($false)}catch{}}
  if($word -ne $null -and $owned){try{$word.Quit()}catch{$result.word_quit_error=$_.Exception.Message}}
  $tmp=$null;$doc=$null;$word=$null;$after=$null;$existing=$null;$paragraph=$null;$donor=$null;$dest=$null
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
  $result|ConvertTo-Json -Depth 6
}
if(-not $result.success){exit 1}
