param(
  [Parameter(Mandatory=$true)][string]$InputDocx,
  [Parameter(Mandatory=$true)][string]$OutputDocx,
  [string]$TemplatePath='',
  [string]$ReportPath='',
  [string]$ControlPath='',
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
if([string]::IsNullOrWhiteSpace($ReportPath)){$ReportPath=$OutputDocx+'.conversion.json'}
if([string]::IsNullOrWhiteSpace($ControlPath)){$ControlPath=$OutputDocx+'.control.json'}
$ReportFull=[IO.Path]::GetFullPath($ReportPath)
$ControlFull=[IO.Path]::GetFullPath($ControlPath)
foreach($pair in @(@('ReportPath',$ReportFull),@('ControlPath',$ControlFull))){
  if([string]::Equals($pair[1],$InputFull,[StringComparison]::OrdinalIgnoreCase) -or [string]::Equals($pair[1],$OutputFull,[StringComparison]::OrdinalIgnoreCase)){
    throw "$($pair[0]) must not point to the input or output DOCX."
  }
}
if([string]::Equals($ReportFull,$ControlFull,[StringComparison]::OrdinalIgnoreCase)){throw 'ReportPath and ControlPath must be different files.'}

$key='HKCU:\Software\AxMath\WordCmds'
$result=[ordered]@{
  source=$InputDocx
  output=$OutputDocx
  macro='AMSMML2AM'
  strategy='official-batch-loop-mode0-guarded'
  success=$false
  complete=$false
  batches=@()
  word_forced_cleanup=$false
}
$total=[Diagnostics.Stopwatch]::StartNew()
$word=$null
$doc=$null
$oldNo=$null
$oldWait=$null
$beforeDumps=@()
$knownDumps=@()
$wordPid=$null
$wordPidOwned=$false
$normalWasSaved=$null

function Write-ControlState {
  param(
    [string]$Phase,
    [string]$Token='',
    [int]$Batch=0,
    [object[]]$Preexisting=@(),
    [int]$Before=0,
    [int]$After=0
  )
  $state=[ordered]@{
    phase=$Phase
    token=$Token
    batch=$Batch
    preexisting_dialog_pids=@($Preexisting)
    before_omath=$Before
    after_omath=$After
    updated_at=(Get-Date).ToString('o')
  }
  $json=$state|ConvertTo-Json -Depth 4
  Set-Content -LiteralPath $ControlPath -Value $json -Encoding UTF8
}

try {
  $existingWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
  $result.preexisting_word_pids=@($existingWord)

  if(Test-Path -LiteralPath $OutputDocx){Remove-Item -LiteralPath $OutputDocx -Force}
  Copy-Item -LiteralPath $InputDocx -Destination $OutputDocx
  # A frozen source may intentionally carry the Windows read-only attribute.
  # The working copy must be writable or Word will divert Save() into an
  # interactive Save As dialog after the first AxMath batch.
  $outputItem=Get-Item -LiteralPath $OutputFull
  $result.output_readonly_cleared=[bool]$outputItem.IsReadOnly
  if($outputItem.IsReadOnly){$outputItem.IsReadOnly=$false}

  if(-not(Test-Path $key)){New-Item -Path $key -Force|Out-Null}
  $props=Get-ItemProperty $key
  $oldNo=$props.DoNoWinVerb
  $oldWait=$props.WaitingConvert
  $result.old_DoNoWinVerb=$oldNo
  $result.old_WaitingConvert=$oldWait
  $crashDir=Join-Path $env:LOCALAPPDATA 'CrashDumps'
  $beforeDumps=@(
    Get-ChildItem $crashDir -Filter 'AxMath*.dmp' -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty Name
  )
  $knownDumps=@($beforeDumps)

  Set-ItemProperty $key -Name DoNoWinVerb -Value 0
  Set-ItemProperty $key -Name WaitingConvert -Value 0

  Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class WordPidNative {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@ -ErrorAction SilentlyContinue

  $word=New-Object -ComObject Word.Application

  # Establish ownership before changing application options or opening a
  # document. Pre-existing user Word processes are allowed, but this task must
  # prove that its COM object belongs to a distinct WINWORD process.
  Start-Sleep -Milliseconds 250
  $newWord=@(
    Get-Process WINWORD -ErrorAction SilentlyContinue |
    Where-Object { $existingWord -notcontains $_.Id } |
    Select-Object -ExpandProperty Id
  )
  if($newWord.Count -eq 1){
    $wordPid=[int]$newWord[0]
    $wordPidOwned=$true
    $result.word_pid=$wordPid
    $result.word_pid_owned=$true
  } else {
    try {
      [uint32]$pidValue=0
      [void][WordPidNative]::GetWindowThreadProcessId([IntPtr]$word.Hwnd,[ref]$pidValue)
      if($pidValue -gt 0){
        $wordPid=[int]$pidValue
        $wordPidOwned=($existingWord -notcontains $wordPid)
        $result.word_pid=$wordPid
        $result.word_pid_owned=$wordPidOwned
      }
    } catch {}
  }
  if(-not $wordPidOwned){
    throw "Could not establish an isolated task-owned Word process; pre-existing Word processes were left untouched."
  }

  $word.Visible=$false
  $word.DisplayAlerts=0
  try{$word.ScreenUpdating=$false}catch{}
  try{$word.Options.CheckSpellingAsYouType=$false}catch{}
  try{$word.Options.CheckGrammarAsYouType=$false}catch{}
  try{$word.Options.SaveNormalPrompt=$false}catch{}
  try{$normalWasSaved=[bool]$word.NormalTemplate.Saved}catch{}

  $found=$false
  foreach($addin in $word.AddIns){
    if($addin.Name -eq 'AxMath.dotm'){
      $found=$true
      if(-not $addin.Installed){$addin.Installed=$true}
    }
  }
  $addin=$null
  if(-not $found){$word.AddIns.Add($TemplatePath,$true)|Out-Null}

  $doc=$word.Documents.Open($OutputDocx)
  $doc.Activate()
  $result.before_omath=$doc.OMaths.Count
  $result.before_inline_shapes=$doc.InlineShapes.Count
  $result.before_paragraphs=$doc.Paragraphs.Count

  $batch=0
  while($doc.OMaths.Count -gt 0 -and $batch -lt 64){
    $batch++
    $before=$doc.OMaths.Count
    Set-ItemProperty $key -Name WaitingConvert -Value 0

    $preDialogs=@(Get-Process AxMMDlg -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $token=('batch-{0}-{1}-{2}' -f $batch,$before,[DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds())
    Write-ControlState -Phase 'awaiting_batch_dialog' -Token $token -Batch $batch -Preexisting $preDialogs -Before $before

    $doc.Content.Select()
    $cb=$null
    $sw=[Diagnostics.Stopwatch]::StartNew()
    $word.Run('AMSMML2AM',([ref]$cb))
    $sw.Stop()

    $after=$doc.OMaths.Count
    Write-ControlState -Phase 'idle' -Token $token -Batch $batch -Before $before -After $after

    $converted=$before-$after
    $waitState=(Get-ItemProperty $key).WaitingConvert
    $currentDumps=@(
      Get-ChildItem $crashDir -Filter 'AxMath*.dmp' -ErrorAction SilentlyContinue |
      Select-Object -ExpandProperty Name
    )
    $batchDumps=@($currentDumps | Where-Object {$_ -notin $knownDumps})
    $knownDumps=@($currentDumps)
    $batchRec=[ordered]@{
      batch=$batch
      before_omath=$before
      after_omath=$after
      converted=$converted
      seconds=$sw.Elapsed.TotalSeconds
      waiting_convert=$waitState
      save_seconds=0.0
      crash_dumps_new=@($batchDumps)
    }

    if($converted -le 0){
      $result.batches += [pscustomobject]$batchRec
      throw "AxMath batch $batch made no progress ($before -> $after)."
    }

    # Keep the per-batch save. AxMath 2.7.x has produced crash dumps during
    # successful large-document conversions; persisting each completed batch
    # is a recovery boundary, not disposable overhead.
    $saveSw=[Diagnostics.Stopwatch]::StartNew()
    $doc.Save()
    $saveSw.Stop()
    $batchRec.save_seconds=$saveSw.Elapsed.TotalSeconds
    $result.batches += [pscustomobject]$batchRec
    Write-Output ('BATCH {0}: {1}->{2}, converted={3}, macro={4:N2}s, save={5:N2}s, crashes={6}' -f $batch,$before,$after,$converted,$sw.Elapsed.TotalSeconds,$saveSw.Elapsed.TotalSeconds,$batchDumps.Count)
  }

  $result.after_omath_com=$doc.OMaths.Count
  $result.after_inline_shapes_com=$doc.InlineShapes.Count
  $result.after_paragraphs=$doc.Paragraphs.Count
  $result.complete=($doc.OMaths.Count -eq 0)
  if(-not $result.complete){
    throw "Conversion stopped with $($doc.OMaths.Count) OfficeMath equations remaining."
  }

  try{$word.ScreenUpdating=$true}catch{}
  $result.pages=$doc.ComputeStatistics(2)
  $doc.Save()
  $result.success=$true
}
catch {
  $result.error=$_.Exception.Message
  $result.hresult=$_.Exception.HResult
}
finally {
  try{Write-ControlState -Phase 'finished'}catch{}

  if($doc -ne $null){
    try{$doc.Close($false)}catch{$result.doc_close_error=$_.Exception.Message}
  }

  if($word -ne $null -and $wordPidOwned){
    try{
      if($normalWasSaved -eq $true -and -not [bool]$word.NormalTemplate.Saved){
        $word.NormalTemplate.Saved=$true
      }
    }catch{}
    try{$word.DisplayAlerts=0}catch{}
    try{$word.Quit()}catch{$result.word_quit_error=$_.Exception.Message}
  } elseif($word -ne $null) {
    # Never call Quit() on a COM object whose process ownership could not be
    # proven; it may belong to a user's pre-existing Word session.
    $result.word_quit_skipped_unowned=$true
  }

  $doc=$null
  $word=$null
  $addin=$null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()

  if($null -ne $oldNo){Set-ItemProperty $key -Name DoNoWinVerb -Value $oldNo}
  if($null -ne $oldWait){Set-ItemProperty $key -Name WaitingConvert -Value $oldWait}

  if($wordPid -and $wordPidOwned){
    $deadline=(Get-Date).AddSeconds(6)
    while((Get-Date) -lt $deadline -and (Get-Process -Id $wordPid -ErrorAction SilentlyContinue)){
      Start-Sleep -Milliseconds 250
    }
    if(Get-Process -Id $wordPid -ErrorAction SilentlyContinue){
      try{
        Stop-Process -Id $wordPid -Force -ErrorAction Stop
        $result.word_forced_cleanup=$true
      }catch{
        $result.word_cleanup_error=$_.Exception.Message
      }
    }
  }

  Start-Sleep -Milliseconds 300
  $afterDumps=@(
    Get-ChildItem $crashDir -Filter 'AxMath*.dmp' -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty Name
  )
  $result.new_crash_dumps=@($afterDumps|Where-Object{$_ -notin $beforeDumps})
  $total.Stop()
  $result.total_seconds=$total.Elapsed.TotalSeconds
  $macroMeasure=$result.batches | Measure-Object -Property seconds -Sum
  $saveMeasure=$result.batches | Measure-Object -Property save_seconds -Sum
  $convertedMeasure=$result.batches | Measure-Object -Property converted -Maximum
  $result.macro_seconds_total=[double]$macroMeasure.Sum
  $result.save_seconds_total=[double]$saveMeasure.Sum
  $result.non_macro_seconds=[math]::Max(0.0,$result.total_seconds-$result.macro_seconds_total)
  $result.macro_share_percent=if($result.total_seconds -gt 0){100.0*$result.macro_seconds_total/$result.total_seconds}else{0.0}
  $result.observed_max_batch_converted=[int]$convertedMeasure.Maximum
  $result.batch_limit_owner='AxMath plugin; do not force/bypass WaitingConvert batch semantics'

  $json=$result|ConvertTo-Json -Depth 8
  Set-Content -LiteralPath $ReportPath -Value $json -Encoding UTF8
  Write-Output $json
}
