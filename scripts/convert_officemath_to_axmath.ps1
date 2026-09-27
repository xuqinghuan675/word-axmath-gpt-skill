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
  if($existingWord.Count -gt 0){
    throw "Refusing background AxMath conversion because Word is already running: $($existingWord -join ',')"
  }

  if(Test-Path -LiteralPath $OutputDocx){Remove-Item -LiteralPath $OutputDocx -Force}
  Copy-Item -LiteralPath $InputDocx -Destination $OutputDocx

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
  $word.Visible=$false
  $word.DisplayAlerts=0
  try{$word.ScreenUpdating=$false}catch{}
  try{$word.Options.CheckSpellingAsYouType=$false}catch{}
  try{$word.Options.CheckGrammarAsYouType=$false}catch{}
  try{$word.Options.SaveNormalPrompt=$false}catch{}
  try{$normalWasSaved=[bool]$word.NormalTemplate.Saved}catch{}

  # Resolve the isolated WINWORD PID by process-diff first. Hwnd is a fallback
  # because some Word COM builds expose it inconsistently during startup.
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
    $batchRec=[ordered]@{
      batch=$batch
      before_omath=$before
      after_omath=$after
      converted=$converted
      seconds=$sw.Elapsed.TotalSeconds
      waiting_convert=$waitState
    }
    $result.batches += [pscustomobject]$batchRec

    if($converted -le 0){
      throw "AxMath batch $batch made no progress ($before -> $after)."
    }

    $doc.Save()
    Write-Output ('BATCH {0}: {1}->{2}, converted={3}, seconds={4:N2}' -f $batch,$before,$after,$converted,$sw.Elapsed.TotalSeconds)
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

  if($word -ne $null){
    try{
      if($normalWasSaved -eq $true -and -not [bool]$word.NormalTemplate.Saved){
        $word.NormalTemplate.Saved=$true
      }
    }catch{}
    try{$word.DisplayAlerts=0}catch{}
    try{$word.Quit()}catch{$result.word_quit_error=$_.Exception.Message}
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

  $json=$result|ConvertTo-Json -Depth 8
  Set-Content -LiteralPath $ReportPath -Value $json -Encoding UTF8
  Write-Output $json
}
