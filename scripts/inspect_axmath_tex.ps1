param(
  [Parameter(Mandatory=$true)][string]$InputDocx,
  [string]$Ordinals='',
  [switch]$AllAxMath,
  [Parameter(Mandatory=$true)][string]$OutputJson,
  [string]$ExpectedInputSha256='',
  [int]$ExpectedAxMathCount=0,
  [string]$TemplatePath=''
)
$ErrorActionPreference='Stop'
$WordPidContractScript=Join-Path $PSScriptRoot 'word_pid_contract.ps1'
if(-not(Test-Path -LiteralPath $WordPidContractScript)){throw 'word_pid_contract.ps1 missing'}
. $WordPidContractScript

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
  foreach($base in @($pf86,$pf)){
    if($base){
      $candidate=Join-Path $base 'AxMath\MSOffice\AxMath.dotm'
      if(Test-Path -LiteralPath $candidate){return $candidate}
    }
  }
  throw 'AxMath.dotm not found.'
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

$InputFull=[IO.Path]::GetFullPath($InputDocx)
$OutputFull=[IO.Path]::GetFullPath($OutputJson)
if(-not(Test-Path -LiteralPath $InputFull)){throw "Input DOCX not found: $InputFull"}
if([string]::Equals($InputFull,$OutputFull,[StringComparison]::OrdinalIgnoreCase)){
  throw 'OutputJson must not overwrite the input DOCX.'
}
New-Item -ItemType Directory -Force -Path ([IO.Path]::GetDirectoryName($OutputFull))|Out-Null
if($AllAxMath -eq (-not [string]::IsNullOrWhiteSpace($Ordinals))){
  throw 'Supply either -AllAxMath or -Ordinals, but not both.'
}
$targets=@()
if(-not $AllAxMath){
  $targets=@($Ordinals -split ','|ForEach-Object{if($_.Trim()){[int]$_.Trim()}}|Sort-Object -Unique)
  if($targets.Count -eq 0){throw 'No ordinals supplied.'}
  if($targets.Count -gt 24){throw 'Targeted internal-AxMath diagnostic is bounded to 24 reviewed ordinals; use -AllAxMath for the final integrity audit.'}
}
$shaBefore=(Get-SharedSha256 $InputFull)
if($ExpectedInputSha256 -and $shaBefore -ne $ExpectedInputSha256.ToLowerInvariant()){
  throw 'Input SHA-256 differs from -ExpectedInputSha256.'
}
$TemplatePath=Resolve-AxMathTemplate $TemplatePath

$beforeWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$startedAt=Get-Date
$result=[ordered]@{
  schema='axmath-internal-tex-diagnostic/v1'
  input=$InputFull
  input_sha256_before=$shaBefore
  target_scope=if($AllAxMath){'all'}else{'selected'}
  ordinals=if($AllAxMath){@()}else{$targets}
  diagnostic_only=$true
  note='AMSAM2TeX is run only on temporary copies of the selected OLE objects; the input DOCX is opened read-only and never saved.'
  exports=@()
  failed=@()
  integrity_failures=@()
  success=$false
  preexisting_word_pids=@($beforeWord)
  started_at=$startedAt.ToString('o')
}
$word=$null;$doc=$null;$tmp=$null;$wordPid=$null;$wordIdentity=$null;$owned=$false

try{
  $launchStartedAt=Get-Date
  $word=New-Object -ComObject Word.Application
  Start-Sleep -Milliseconds 300
  $wordIdentity=Resolve-OwnedWordIdentity -Word $word -BeforePids $beforeWord -LaunchStartedAt $launchStartedAt
  $wordPid=[int]$wordIdentity.pid;$owned=$true
  $result.word_pid=$wordPid;$result.word_pid_owned=$true
  $result.word_pid_created_at=$wordIdentity.started_at
  $result.word_pid_ownership_proof=$wordIdentity.proof
  $word.Visible=$false;$word.DisplayAlerts=0
  try{$word.ScreenUpdating=$false}catch{}

  $found=$false
  foreach($a in $word.AddIns){if($a.Name -eq 'AxMath.dotm'){$found=$true;if(-not $a.Installed){$a.Installed=$true}}}
  if(-not $found){$word.AddIns.Add($TemplatePath,$true)|Out-Null}

  $doc=$word.Documents.OpenNoRepairDialog($InputFull,$false,$true,$false)
  $ax=@(Get-AxMathShapes $doc)
  $result.axmath_count=$ax.Count
  if($ExpectedAxMathCount -gt 0 -and $ax.Count -ne $ExpectedAxMathCount){
    throw "Actual AxMath count $($ax.Count) != -ExpectedAxMathCount $ExpectedAxMathCount."
  }
  if($AllAxMath){
    if($ax.Count -gt 0){$targets=@(1..$ax.Count)}else{$targets=@()}
  }
  foreach($ord in $targets){
    $rec=[ordered]@{ordinal=$ord;success=$false}
    try{
      if($ord -lt 1 -or $ord -gt $ax.Count){throw "ordinal $ord out of range 1..$($ax.Count)"}
      $target=$ax[$ord-1]
      $rec.geometry=[ordered]@{width=[double]$target.Width;height=[double]$target.Height;start=[int]$target.Range.Start;end=[int]$target.Range.End}
      $target.Range.Copy()
      $tmp=$word.Documents.Add()
      $tmp.Range(0,0).Paste()
      $donors=@(Get-AxMathShapes $tmp)
      if($donors.Count -ne 1){throw "temporary copied AxMath count=$($donors.Count)"}
      $tmp.Content.Select()
      $cb=$null
      $word.Run('AMSAM2TeX',([ref]$cb))
      $tex=[string]$tmp.Content.Text
      $rec.axmath_exported_tex=$tex
      $rec.export_empty=[string]::IsNullOrWhiteSpace($tex)
      $rec.export_contains_control=[bool]($tex -match '[\x00-\x08\x0b\x0c\x0e-\x1f]')
      $rec.success=$true
      $result.exports += [pscustomobject]$rec
      if($rec.export_empty -or $rec.export_contains_control){
        $result.integrity_failures += [pscustomobject]@{
          ordinal=$ord
          reason=if($rec.export_empty){'empty_internal_export'}else{'control_character_in_internal_export'}
        }
      }
      $tmp.Close($false);$tmp=$null
    }catch{
      if($tmp -ne $null){try{$tmp.Close($false)}catch{};$tmp=$null}
      $rec.error=$_.Exception.Message
      $result.failed += [pscustomobject]$rec
    }
    if($AllAxMath -and (($result.exports.Count+$result.failed.Count) % 24 -eq 0)){
      Write-Host "AxMath internal audit: $($result.exports.Count+$result.failed.Count)/$($targets.Count) inspected"
    }
  }
  $result.success=($result.failed.Count -eq 0 -and $result.exports.Count -eq $targets.Count)
  $result.integrity_pass=($result.success -and $result.integrity_failures.Count -eq 0)
}catch{
  $result.error=$_.Exception.Message
  $result.hresult=$_.Exception.HResult
}finally{
  if($tmp -ne $null){try{$tmp.Close($false)}catch{}}
  if($doc -ne $null){try{$doc.Close($false)}catch{}}
  if($word -ne $null -and $owned -and (Test-OwnedWordIdentity $wordIdentity)){try{$word.Quit()}catch{}}
  $tmp=$null;$doc=$null;$word=$null;$ax=$null;$target=$null;$donors=$null
  [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
  if($wordPid -and $owned){
    $deadline=(Get-Date).AddSeconds(6)
    while((Get-Date)-lt $deadline -and (Test-OwnedWordIdentity $wordIdentity)){Start-Sleep -Milliseconds 200}
    if(Test-OwnedWordIdentity $wordIdentity){
      Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue
      $result.forced_owned_word_cleanup=$true
    }
  }
  $shaAfter=(Get-SharedSha256 $InputFull)
  $result.input_sha256_after=$shaAfter
  $result.input_unchanged=($shaAfter -eq $shaBefore)
  if(-not $result.input_unchanged){$result.success=$false}
  $result.integrity_pass=($result.success -and $result.integrity_failures.Count -eq 0)
  $result.export_count=$result.exports.Count
  $result.finished_at=(Get-Date).ToString('o')
  $result.elapsed_seconds=[Math]::Round(((Get-Date)-$startedAt).TotalSeconds,3)
  $result|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $OutputFull -Encoding UTF8
  if($AllAxMath){
    [pscustomobject]@{
      report=$OutputFull
      axmath_count=$result.axmath_count
      export_count=$result.export_count
      failed_count=$result.failed.Count
      integrity_failures=@($result.integrity_failures)
      success=$result.success
      integrity_pass=$result.integrity_pass
      input_unchanged=$result.input_unchanged
    }|ConvertTo-Json -Depth 5
  }else{$result|ConvertTo-Json -Depth 6}
}
if(-not $result.success){exit 1}
