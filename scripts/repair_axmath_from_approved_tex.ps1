param(
  [Parameter(Mandatory=$true)][string]$InputDocx,
  [Parameter(Mandatory=$true)][string]$OutputDocx,
  [Parameter(Mandatory=$true)][string]$MapPath,
  [string]$TemplatePath='',
  [string]$ReportPath='',
  [switch]$OverwriteOutput
)
$ErrorActionPreference='Stop'

function Resolve-AxMathTemplate {
  param([string]$Requested)
  if(-not [string]::IsNullOrWhiteSpace($Requested)){
    if(Test-Path -LiteralPath $Requested){ return (Resolve-Path -LiteralPath $Requested).Path }
    throw "AxMath.dotm not found: $Requested"
  }
  $candidates=@()
  if(${env:ProgramFiles(x86)}){$candidates += (Join-Path ${env:ProgramFiles(x86)} 'AxMath\MSOffice\AxMath.dotm')}
  if($env:ProgramFiles){$candidates += (Join-Path $env:ProgramFiles 'AxMath\MSOffice\AxMath.dotm')}
  foreach($candidate in $candidates){
    if(Test-Path -LiteralPath $candidate){ return $candidate }
  }
  throw 'AxMath.dotm not found. Install/activate AxMath or pass -TemplatePath explicitly.'
}

function Get-AxMathShapes($d){
  $arr=@()
  foreach($s in $d.InlineShapes){
    try{if([string]$s.OLEFormat.ProgID -eq 'Equation.AxMath'){$arr += $s}}catch{}
  }
  return ,@($arr)
}

$InputFull=[IO.Path]::GetFullPath($InputDocx)
$OutputFull=[IO.Path]::GetFullPath($OutputDocx)
$MapFull=[IO.Path]::GetFullPath($MapPath)
if([string]::Equals($InputFull,$OutputFull,[StringComparison]::OrdinalIgnoreCase)){throw 'Refusing to overwrite the input DOCX.'}
if(-not (Test-Path -LiteralPath $InputFull)){throw "Input DOCX not found: $InputFull"}
if(-not (Test-Path -LiteralPath $MapFull)){throw "Approved TeX map not found: $MapFull"}
if((Test-Path -LiteralPath $OutputFull) -and -not $OverwriteOutput){throw "Output already exists: $OutputFull. Use -OverwriteOutput only for an intentional intermediate replacement."}
if([string]::IsNullOrWhiteSpace($ReportPath)){$ReportPath=$OutputDocx+'.source-tex-repair.json'}
$ReportFull=[IO.Path]::GetFullPath($ReportPath)
if([string]::Equals($ReportFull,$InputFull,[StringComparison]::OrdinalIgnoreCase) -or [string]::Equals($ReportFull,$OutputFull,[StringComparison]::OrdinalIgnoreCase) -or [string]::Equals($ReportFull,$MapFull,[StringComparison]::OrdinalIgnoreCase)){throw 'ReportPath must not point to the input DOCX, output DOCX, or approved TeX map.'}

$rows=@(Get-Content -Raw -Encoding UTF8 -LiteralPath $MapFull | ConvertFrom-Json)
if($rows.Count -eq 0){throw 'Approved TeX map is empty.'}
if($rows.Count -gt 48){throw "Approved source-semantic repair is intentionally bounded. Refusing $($rows.Count) rows; split reviewed repairs into batches of at most 48."}
$ordinals=@($rows | ForEach-Object {[int]$_.ordinal})
if(@($ordinals|Sort-Object -Unique).Count -ne $rows.Count){throw 'Approved TeX map contains duplicate ordinals.'}
foreach($row in $rows){
  $tex=[string]$row.tex
  if([string]::IsNullOrWhiteSpace($tex)){throw "Approved TeX is empty for ordinal $($row.ordinal)."}
  if(-not ($tex.StartsWith('$') -and $tex.EndsWith('$')) -or $tex.StartsWith('$$') -or $tex.EndsWith('$$')){throw "Approved TeX must be wrapped in exactly one leading and trailing dollar delimiter for ordinal $($row.ordinal)."}
}

$TemplatePath=Resolve-AxMathTemplate $TemplatePath
$beforeWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$result=[ordered]@{
  input=$InputFull
  output=$OutputFull
  approved_tex_map=$MapFull
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
  $newWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object{$beforeWord -notcontains $_.Id} | Select-Object -ExpandProperty Id)
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
  $initial=Get-AxMathShapes $doc
  $result.axmath_before=$initial.Count
  $result.omath_before=[int]$doc.OMaths.Count
  $result.paragraphs_before=[int]$doc.Paragraphs.Count

  foreach($row in @($rows | Sort-Object {[int]$_.ordinal} -Descending)){
    $ord=[int]$row.ordinal
    $rec=[ordered]@{ordinal=$ord;success=$false;tex=[string]$row.tex}
    $tmp=$null
    try{
      $ax=Get-AxMathShapes $doc
      if($ax.Count -ne $result.axmath_before){throw "AxMath count drifted before ordinal ${ord}: $($ax.Count) != $($result.axmath_before)"}
      if($ord -lt 1 -or $ord -gt $ax.Count){throw "ordinal $ord out of range 1..$($ax.Count)"}
      $target=$ax[$ord-1]
      $oldStart=[int]$target.Range.Start
      $oldEnd=[int]$target.Range.End
      $rec.before=[ordered]@{width=[double]$target.Width;height=[double]$target.Height;start=$oldStart;end=$oldEnd}

      $tmp=$word.Documents.Add()
      $tmp.Content.Text=[string]$row.tex
      $tmp.Content.Select()
      $cb=$null
      $word.Run('AMSTeX2AM',([ref]$cb))
      $donorAx=Get-AxMathShapes $tmp
      if($donorAx.Count -ne 1){throw "approved TeX donor count=$($donorAx.Count)"}
      $donor=$donorAx[0]
      $dw=[double]$donor.Width;$dh=[double]$donor.Height
      if($dw -le 4.0 -or $dw -gt 560.0 -or $dh -le 8.0 -or $dh -gt 160.0){
        throw "approved donor geometry rejected: $dw x $dh"
      }
      $rec.donor=[ordered]@{width=$dw;height=$dh}

      $donor.Range.Copy()
      $doc.Range($oldStart,$oldEnd).Delete()
      $doc.Range($oldStart,$oldStart).Paste()
      $tmp.Close($false);$tmp=$null

      $after=Get-AxMathShapes $doc
      if($after.Count -ne $result.axmath_before){throw "AxMath count changed after ordinal ${ord}: $($after.Count)"}
      $replacement=$after[$ord-1]
      try{$prog=[string]$replacement.OLEFormat.ProgID}catch{$prog=''}
      if($prog -ne 'Equation.AxMath'){throw "replacement ordinal $ord is not Equation.AxMath"}
      $rec.after=[ordered]@{width=[double]$replacement.Width;height=[double]$replacement.Height;start=[int]$replacement.Range.Start;end=[int]$replacement.Range.End}
      $rec.success=$true
      $result.repairs += [pscustomobject]$rec

      # Persist every reviewed semantic replacement as its own recovery boundary.
      $doc.Save()
    }catch{
      if($tmp -ne $null){try{$tmp.Close($false)}catch{};$tmp=$null}
      $rec.error=$_.Exception.Message
      $result.failed += [pscustomobject]$rec
      throw
    }
  }

  $final=Get-AxMathShapes $doc
  $result.axmath_after=$final.Count
  $result.omath_after=[int]$doc.OMaths.Count
  $result.paragraphs_after=[int]$doc.Paragraphs.Count
  $doc.Save()
  $result.success=(
    $result.failed.Count -eq 0 -and
    $result.repairs.Count -eq $rows.Count -and
    $result.axmath_after -eq $result.axmath_before -and
    $result.omath_after -eq $result.omath_before -and
    $result.paragraphs_after -eq $result.paragraphs_before
  )
}catch{
  $result.error=$_.Exception.Message
  $result.hresult=$_.Exception.HResult
}finally{
  if($doc -ne $null){try{$doc.Close($false)}catch{}}
  if($word -ne $null -and $owned){
    try{$word.Quit()}catch{$result.word_quit_error=$_.Exception.Message}
  } elseif($word -ne $null) {
    $result.word_quit_skipped_unowned=$true
  }
  $doc=$null;$word=$null
  [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
  if($wordPid -and $owned){
    Start-Sleep -Milliseconds 500
    if(Get-Process -Id $wordPid -ErrorAction SilentlyContinue){
      Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue
      $result.forced_owned_word_cleanup=$true
    }
  }
  if(Test-Path -LiteralPath $OutputFull){$result.output_sha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $OutputFull).Hash.ToLowerInvariant()}
  $result|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $ReportFull -Encoding UTF8
  $result|ConvertTo-Json -Depth 8
}
if(-not $result.success){exit 1}
