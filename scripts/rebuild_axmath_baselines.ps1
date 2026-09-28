param(
  [Parameter(Mandatory=$true)][string]$InputDocx,
  [Parameter(Mandatory=$true)][string]$OutputDocx,
  [Parameter(Mandatory=$true)][string]$Ordinals,
  [string]$TemplatePath='',
  [string]$ReportPath='',
  [switch]$OverwriteOutput
)
$ErrorActionPreference='Stop'
$InputFull=[IO.Path]::GetFullPath($InputDocx)
$OutputFull=[IO.Path]::GetFullPath($OutputDocx)
if([string]::Equals($InputFull,$OutputFull,[StringComparison]::OrdinalIgnoreCase)){throw 'Refusing to overwrite the input DOCX.'}
if(-not (Test-Path -LiteralPath $InputFull)){throw "Input DOCX not found: $InputFull"}
if((Test-Path -LiteralPath $OutputFull) -and -not $OverwriteOutput){throw "Output already exists: $OutputFull. Use -OverwriteOutput only for an intentional intermediate replacement."}
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
$TemplatePath=Resolve-AxMathTemplate $TemplatePath
if([string]::IsNullOrWhiteSpace($ReportPath)){$ReportPath=$OutputDocx+'.rebuild.json'}
$ReportFull=[IO.Path]::GetFullPath($ReportPath)
if([string]::Equals($ReportFull,$InputFull,[StringComparison]::OrdinalIgnoreCase) -or [string]::Equals($ReportFull,$OutputFull,[StringComparison]::OrdinalIgnoreCase)){throw 'ReportPath must not point to the input or output DOCX.'}
$targets=@($Ordinals -split ',' | ForEach-Object { if($_.Trim()){ [int]$_.Trim() } } | Sort-Object -Unique -Descending)
if($targets.Count -eq 0){throw 'No ordinals supplied.'}
if($targets.Count -gt 3){
  throw "ConvertAMERebuild is probe-only in production. Refusing $($targets.Count) ordinals; use at most 3 explicit formulas and re-audit before any further repair."
}

$result=[ordered]@{
  input=$InputDocx
  output=$OutputDocx
  targets=$targets
  macro='AxMath_Proj.AMCCAMEqn2TeX.ConvertAMERebuild'
  probe_only=$true
  repaired=@()
  failed=@()
  success=$false
}
$word=$null;$doc=$null;$wordPid=$null;$owned=$false
$beforeWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$result.preexisting_word_pids=@($beforeWord)

try{
  if(Test-Path -LiteralPath $OutputDocx){Remove-Item -LiteralPath $OutputDocx -Force}
  Copy-Item -LiteralPath $InputDocx -Destination $OutputDocx -Force
  $outputItem=Get-Item -LiteralPath $OutputFull
  $result.output_readonly_cleared=[bool]$outputItem.IsReadOnly
  if($outputItem.IsReadOnly){$outputItem.IsReadOnly=$false}

  $word=New-Object -ComObject Word.Application
  Start-Sleep -Milliseconds 300
  $afterWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object{$beforeWord -notcontains $_.Id} | Select-Object -ExpandProperty Id)
  if($afterWord.Count -eq 1){
    $wordPid=[int]$afterWord[0]
    $owned=$true
    $result.word_pid=$wordPid
    $result.word_pid_owned=$true
  }
  if(-not $owned){
    throw 'Could not establish a distinct task-owned Word process; pre-existing user Word processes were left untouched.'
  }
  $word.Visible=$true
  $word.DisplayAlerts=0

  $found=$false
  foreach($a in $word.AddIns){
    if($a.Name -eq 'AxMath.dotm'){
      $found=$true
      if(-not $a.Installed){$a.Installed=$true}
    }
  }
  if(-not $found){$word.AddIns.Add($TemplatePath,$true)|Out-Null}

  $doc=$word.Documents.Open($OutputDocx)
  $doc.Activate()
  Start-Sleep -Milliseconds 300

  $initial=@()
  foreach($s in $doc.InlineShapes){try{if($s.OLEFormat.ProgID -eq 'Equation.AxMath'){$initial+=$s}}catch{}}
  $result.axmath_before=$initial.Count

  foreach($ord in $targets){
    $rec=[ordered]@{ordinal=$ord;success=$false}
    try{
      $ax=@()
      foreach($s in $doc.InlineShapes){try{if($s.OLEFormat.ProgID -eq 'Equation.AxMath'){$ax+=$s}}catch{}}
      if($ord -lt 1 -or $ord -gt $ax.Count){throw "ordinal $ord out of range 1..$($ax.Count)"}
      $t=$ax[$ord-1]
      $start=[int]$t.Range.Start
      $rec.before=[ordered]@{
        width=[double]$t.Width
        height=[double]$t.Height
        x=[double]$t.Range.Information(5)
        y=[double]$t.Range.Information(6)
      }

      $t.Range.Select()
      Start-Sleep -Milliseconds 100
      $sw=[Diagnostics.Stopwatch]::StartNew()
      $word.Run('AxMath_Proj.AMCCAMEqn2TeX.ConvertAMERebuild')
      $sw.Stop()
      $rec.seconds=$sw.Elapsed.TotalSeconds
      Start-Sleep -Milliseconds 200

      $ax2=@()
      foreach($s in $doc.InlineShapes){try{if($s.OLEFormat.ProgID -eq 'Equation.AxMath'){$ax2+=$s}}catch{}}
      if($ax2.Count -ne $initial.Count){throw "AxMath count changed $($initial.Count) -> $($ax2.Count)"}
      $repl=$null
      foreach($s in $ax2){
        if([math]::Abs([int]$s.Range.Start-$start)-le 2){$repl=$s;break}
      }
      if($null -eq $repl){throw 'rebuilt object not found near original range'}
      $rec.after=[ordered]@{
        width=[double]$repl.Width
        height=[double]$repl.Height
        x=[double]$repl.Range.Information(5)
        y=[double]$repl.Range.Information(6)
      }
      $rec.success=$true
      $result.repaired += [pscustomobject]$rec
      Write-Output ("REBUILT {0}: {1:N1}x{2:N1} -> {3:N1}x{4:N1}" -f $ord,$rec.before.width,$rec.before.height,$rec.after.width,$rec.after.height)
    }catch{
      $rec.error=$_.Exception.Message
      $result.failed += [pscustomobject]$rec
      Write-Output ("FAILED {0}: {1}" -f $ord,$rec.error)
    }
  }

  $doc.Save()
  $final=@()
  foreach($s in $doc.InlineShapes){try{if($s.OLEFormat.ProgID -eq 'Equation.AxMath'){$final+=$s}}catch{}}
  $result.axmath_after=$final.Count
  $result.pages=[int]$doc.ComputeStatistics(2)
  $result.success=($result.failed.Count -eq 0 -and $final.Count -eq $initial.Count)
}catch{
  $result.fatal_error=$_.Exception.Message
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
    $deadline=(Get-Date).AddSeconds(6)
    while((Get-Date)-lt $deadline -and (Get-Process -Id $wordPid -ErrorAction SilentlyContinue)){Start-Sleep -Milliseconds 250}
    if(Get-Process -Id $wordPid -ErrorAction SilentlyContinue){
      Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue
      $result.word_forced_cleanup=$true
    }
  }
  $result.finished_at=(Get-Date).ToString('o')
  $result|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $ReportPath -Encoding UTF8
  $result|ConvertTo-Json -Depth 8
}
if(-not $result.success){exit 1}
