param(
  [Parameter(Mandatory=$true)][string]$InputDocx,
  [Parameter(Mandatory=$true)][string]$Ordinals,
  [Parameter(Mandatory=$true)][string]$OutputJson,
  [string]$TemplatePath=''
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
$targets=@($Ordinals -split ','|ForEach-Object{if($_.Trim()){[int]$_.Trim()}}|Sort-Object -Unique)
if($targets.Count -eq 0){throw 'No ordinals supplied.'}
if($targets.Count -gt 24){throw 'Internal-AxMath diagnostic is intentionally bounded to 24 reviewed ordinals.'}
$shaBefore=(Get-SharedSha256 $InputFull)
$TemplatePath=Resolve-AxMathTemplate $TemplatePath

$beforeWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$result=[ordered]@{
  schema='axmath-internal-tex-diagnostic/v1'
  input=$InputFull
  input_sha256_before=$shaBefore
  ordinals=$targets
  diagnostic_only=$true
  note='AMSAM2TeX is run only on temporary copies of the selected OLE objects; the input DOCX is opened read-only and never saved.'
  exports=@()
  failed=@()
  success=$false
  preexisting_word_pids=@($beforeWord)
}
$word=$null;$doc=$null;$tmp=$null;$wordPid=$null;$owned=$false

try{
  $word=New-Object -ComObject Word.Application
  Start-Sleep -Milliseconds 300
  $new=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Where-Object{$beforeWord -notcontains $_.Id}|Select-Object -ExpandProperty Id)
  if($new.Count -eq 1){$wordPid=[int]$new[0];$owned=$true;$result.word_pid=$wordPid;$result.word_pid_owned=$true}
  if(-not $owned){throw 'Could not establish a distinct task-owned Word process.'}
  $word.Visible=$false;$word.DisplayAlerts=0
  try{$word.ScreenUpdating=$false}catch{}

  $found=$false
  foreach($a in $word.AddIns){if($a.Name -eq 'AxMath.dotm'){$found=$true;if(-not $a.Installed){$a.Installed=$true}}}
  if(-not $found){$word.AddIns.Add($TemplatePath,$true)|Out-Null}

  $doc=$word.Documents.OpenNoRepairDialog($InputFull,$false,$true,$false)
  $ax=@(Get-AxMathShapes $doc)
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
      $tmp.Close($false);$tmp=$null
    }catch{
      if($tmp -ne $null){try{$tmp.Close($false)}catch{};$tmp=$null}
      $rec.error=$_.Exception.Message
      $result.failed += [pscustomobject]$rec
    }
  }
  $result.success=($result.failed.Count -eq 0 -and $result.exports.Count -eq $targets.Count)
}catch{
  $result.error=$_.Exception.Message
  $result.hresult=$_.Exception.HResult
}finally{
  if($tmp -ne $null){try{$tmp.Close($false)}catch{}}
  if($doc -ne $null){try{$doc.Close($false)}catch{}}
  if($word -ne $null -and $owned){try{$word.Quit()}catch{}}
  $tmp=$null;$doc=$null;$word=$null;$ax=$null;$target=$null;$donors=$null
  [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
  if($wordPid -and $owned){
    $deadline=(Get-Date).AddSeconds(6)
    while((Get-Date)-lt $deadline -and (Get-Process -Id $wordPid -ErrorAction SilentlyContinue)){Start-Sleep -Milliseconds 200}
    if(Get-Process -Id $wordPid -ErrorAction SilentlyContinue){
      Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue
      $result.forced_owned_word_cleanup=$true
    }
  }
  $shaAfter=(Get-SharedSha256 $InputFull)
  $result.input_sha256_after=$shaAfter
  $result.input_unchanged=($shaAfter -eq $shaBefore)
  if(-not $result.input_unchanged){$result.success=$false}
  $result.finished_at=(Get-Date).ToString('o')
  $result|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $OutputFull -Encoding UTF8
  $result|ConvertTo-Json -Depth 6
}
if(-not $result.success){exit 1}
