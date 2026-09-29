param(
  [Parameter(Mandatory=$true)][string]$SourceDocx,
  [Parameter(Mandatory=$true)][string]$Ordinals,
  [Parameter(Mandatory=$true)][string]$OutputJson
)
$ErrorActionPreference='Stop'

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


$SourceFull=[IO.Path]::GetFullPath($SourceDocx)
$OutputFull=[IO.Path]::GetFullPath($OutputJson)
if(-not (Test-Path -LiteralPath $SourceFull)){throw "Source DOCX not found: $SourceFull"}
if([string]::Equals($SourceFull,$OutputFull,[StringComparison]::OrdinalIgnoreCase)){throw 'OutputJson must not point to the source DOCX.'}
$targets=@($Ordinals -split ',' | ForEach-Object { if($_.Trim()){ [int]$_.Trim() } } | Sort-Object -Unique)
if($targets.Count -eq 0){throw 'No ordinals supplied.'}
if($targets.Count -gt 48){throw "Source LaTeX export is intentionally bounded. Refusing $($targets.Count) ordinals; inspect/export at most 48 reviewed formulas per batch."}

$sourceShaBefore=(Get-SharedSha256 $SourceFull)
$beforeWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$result=[ordered]@{
  source=$SourceFull
  source_sha256_before=$sourceShaBefore
  ordinals=$targets
  preexisting_word_pids=@($beforeWord)
  exports=@()
  failed=@()
  success=$false
}
$word=$null;$src=$null;$wordPid=$null;$owned=$false

try{
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
  try{$word.Options.SaveNormalPrompt=$false}catch{}
  $src=$word.Documents.OpenNoRepairDialog($SourceFull,$false,$true,$false)
  $sourceCount=[int]$src.OMaths.Count
  $result.source_omath_count=$sourceCount

  foreach($ord in $targets){
    $rec=[ordered]@{ordinal=$ord;success=$false}
    $tmp=$null
    try{
      if($ord -lt 1 -or $ord -gt $sourceCount){throw "ordinal $ord out of range 1..$sourceCount"}
      $tmp=$word.Documents.Add()
      $tmp.Range(0,0).FormattedText=$src.OMaths.Item($ord).Range.FormattedText
      if([int]$tmp.OMaths.Count -ne 1){throw "temporary OfficeMath count=$($tmp.OMaths.Count)"}

      # Word dialog 2844 switches the selected OfficeMath representation to
      # LaTeX; Linearize then exposes that representation as plain text. Work
      # only in the temporary document, never in the frozen source.
      $tmp.Content.Select()
      [void]$word.Dialogs.Item(2844).Execute()
      if([int]$tmp.OMaths.Count -ne 1){throw "OfficeMath disappeared during Word LaTeX export"}
      $tmp.OMaths.Item(1).Linearize()
      $latex=[string]$tmp.Content.Text.Trim([char]13,[char]10,[char]32,[char]9)
      if([string]::IsNullOrWhiteSpace($latex)){throw 'Word LaTeX export returned empty text'}
      $rec.word_latex=$latex
      $rec.success=$true
      $result.exports += [pscustomobject]$rec
    }catch{
      $rec.error=$_.Exception.Message
      $result.failed += [pscustomobject]$rec
    }finally{
      if($tmp -ne $null){try{$tmp.Close($false)}catch{}}
    }
  }

  $result.success=($result.failed.Count -eq 0 -and $result.exports.Count -eq $targets.Count)
}catch{
  $result.error=$_.Exception.Message
  $result.hresult=$_.Exception.HResult
}finally{
  if($src -ne $null){try{$src.Close($false)}catch{}}
  if($word -ne $null -and $owned){
    try{$word.Quit()}catch{$result.word_quit_error=$_.Exception.Message}
  } elseif($word -ne $null) {
    $result.word_quit_skipped_unowned=$true
  }
  $src=$null;$word=$null
  [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
  if($wordPid -and $owned){
    Start-Sleep -Milliseconds 500
    if(Get-Process -Id $wordPid -ErrorAction SilentlyContinue){
      Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue
      $result.forced_owned_word_cleanup=$true
    }
  }

  $sourceShaAfter=(Get-SharedSha256 $SourceFull)
  $result.source_sha256_after=$sourceShaAfter
  $result.source_unchanged=($sourceShaAfter -eq $sourceShaBefore)
  if(-not $result.source_unchanged){$result.success=$false;$result.source_changed_error='Frozen source SHA changed during LaTeX export.'}
  $result|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $OutputFull -Encoding UTF8
  $result|ConvertTo-Json -Depth 6
}
if(-not $result.success){exit 1}
