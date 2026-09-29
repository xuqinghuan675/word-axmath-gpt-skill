param(
  [Parameter(Mandatory=$true)][string]$SourceDocx,
  [Parameter(Mandatory=$true)][string]$Ordinals,
  [Parameter(Mandatory=$true)][string]$OutputJson,
  [string]$FragmentsDir='',
  [double]$YTolerancePt=1.5,
  [int]$MaxRetries=2
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


function Get-WordPids {
  return @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
}

function Get-LayoutPoint {
  param($Doc,[int]$Position)
  $r=$Doc.Range($Position,$Position)
  $page=[int]$r.Information(3)
  $x=[double]$r.Information(5)
  $y=[double]$r.Information(6)
  if($page -le 0 -or $x -lt 0 -or $y -lt 0){
    throw "Word layout position unavailable at document position $Position"
  }
  return [pscustomobject]@{page=$page;x=$x;y=$y}
}

function Same-VisualLine {
  param($A,$B,[double]$Tolerance)
  return (
    $A.page -eq $B.page -and
    [Math]::Abs([double]$A.y-[double]$B.y) -le $Tolerance
  )
}

function Get-VisualLineRanges {
  param($Doc,$MathRange,[double]$Tolerance)
  $start=[int]$MathRange.Start
  $end=[int]$MathRange.End
  if($end -le $start){throw 'OfficeMath range is empty.'}
  $ranges=@()
  $cursor=$start
  while($cursor -lt $end){
    if($ranges.Count -ge 32){throw 'Refusing more than 32 visual lines for one formula.'}
    $p0=Get-LayoutPoint $Doc $cursor
    $pend=Get-LayoutPoint $Doc $end
    if(Same-VisualLine $p0 $pend $Tolerance){
      $ranges += [pscustomobject]@{
        start=$cursor;end=$end;page=$p0.page;x=$p0.x;y=$p0.y
      }
      break
    }

    # Word's collapsed-position Y is piecewise constant across a visual line.
    # Find the first position whose page/Y differs from the current line.
    $lo=$cursor+1
    $hi=$end
    while($lo -lt $hi){
      $mid=[int][Math]::Floor(($lo+$hi)/2.0)
      $pm=Get-LayoutPoint $Doc $mid
      if(Same-VisualLine $p0 $pm $Tolerance){$lo=$mid+1}else{$hi=$mid}
    }
    $boundary=$lo
    if($boundary -le $cursor -or $boundary -gt $end){
      throw "Could not resolve a monotonic visual-line boundary at $cursor"
    }

    # Do not trust the binary-search assumption blindly. Sample the resolved
    # interval; if Word exposes non-monotonic internal layout coordinates,
    # stop instead of mistaking a fraction/superscript for a line break.
    foreach($fraction in @(0.25,0.5,0.75)){
      $probe=$cursor+[int][Math]::Floor(($boundary-$cursor)*$fraction)
      if($probe -gt $cursor -and $probe -lt $boundary){
        $pp=Get-LayoutPoint $Doc $probe
        if(-not (Same-VisualLine $p0 $pp $Tolerance)){
          throw "Non-monotonic OfficeMath layout inside visual line at $probe; manual/source-semantic review required."
        }
      }
    }

    $ranges += [pscustomobject]@{
      start=$cursor;end=$boundary;page=$p0.page;x=$p0.x;y=$p0.y
    }
    $cursor=$boundary
  }
  return @($ranges)
}

function Invoke-ExportOne {
  param(
    [string]$SourceFull,
    [int]$Ordinal,
    [string]$FragmentsFull,
    [double]$Tolerance
  )
  $before=@(Get-WordPids)
  $word=$null;$src=$null;$tmp=$null;$wordPid=$null;$owned=$false
  $rec=[ordered]@{ordinal=$Ordinal;success=$false;preexisting_word_pids=@($before)}
  try{
    $word=New-Object -ComObject Word.Application
    Start-Sleep -Milliseconds 300
    $new=@(Get-WordPids | Where-Object{$before -notcontains $_})
    if($new.Count -eq 1){
      $wordPid=[int]$new[0]
      $owned=$true
      $rec.word_pid=$wordPid
      $rec.word_pid_owned=$true
    }
    if(-not $owned){
      throw 'Could not establish a distinct task-owned Word process.'
    }
    $word.Visible=$false
    $word.DisplayAlerts=0
    try{$word.ScreenUpdating=$false}catch{}
    try{$word.Options.SaveNormalPrompt=$false}catch{}

    $src=$word.Documents.OpenNoRepairDialog($SourceFull,$false,$true,$false)
    if($Ordinal -lt 1 -or $Ordinal -gt [int]$src.OMaths.Count){
      throw "ordinal $Ordinal out of range 1..$($src.OMaths.Count)"
    }
    $math=$src.OMaths.Item($Ordinal)
    $ranges=@(Get-VisualLineRanges $src $math.Range $Tolerance)
    if($ranges.Count -lt 1){throw 'No visual line ranges were resolved.'}
    $rec.visual_line_count=$ranges.Count
    $rec.lines=@()
    $lineNo=0
    foreach($range in $ranges){
      $lineNo++
      $tmp=$word.Documents.Add()
      $tmp.Range(0,0).FormattedText=$src.Range([int]$range.start,[int]$range.end).FormattedText
      if([int]$tmp.OMaths.Count -ne 1){
        throw "visual line $lineNo did not preserve exactly one OfficeMath object"
      }

      $fragment=Join-Path $FragmentsFull ("omath-{0:D4}-line-{1:D2}.docx" -f $Ordinal,$lineNo)
      if(Test-Path -LiteralPath $fragment){Remove-Item -LiteralPath $fragment -Force}
      $tmp.SaveAs2($fragment,16)

      $wordLatex=''
      $wordLatexError=$null
      try{
        $tmp.Content.Select()
        [void]$word.Dialogs.Item(2844).Execute()
        if([int]$tmp.OMaths.Count -ne 1){throw 'OfficeMath disappeared during Word LaTeX export'}
        $tmp.OMaths.Item(1).Linearize()
        $wordLatex=[string]$tmp.Content.Text.Trim([char]13,[char]10,[char]32,[char]9)
      }catch{
        $wordLatexError=$_.Exception.Message
      }

      $rec.lines += [pscustomobject]@{
        line=$lineNo
        start=[int]$range.start
        end=[int]$range.end
        page=[int]$range.page
        x_pt=[double]$range.x
        y_pt=[double]$range.y
        fragment_docx=$fragment
        word_latex=$wordLatex
        word_latex_error=$wordLatexError
      }
      $tmp.Close($false)
      $tmp=$null
    }
    $rec.success=$true
  }catch{
    $rec.error=$_.Exception.Message
    $rec.hresult=$_.Exception.HResult
  }finally{
    if($tmp -ne $null){try{$tmp.Close($false)}catch{}}
    if($src -ne $null){try{$src.Close($false)}catch{}}
    if($word -ne $null -and $owned){try{$word.Quit()}catch{$rec.word_quit_error=$_.Exception.Message}}
    $tmp=$null;$src=$null;$word=$null
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
    if($wordPid -and $owned){
      $deadline=(Get-Date).AddSeconds(6)
      while((Get-Date)-lt $deadline -and (Get-Process -Id $wordPid -ErrorAction SilentlyContinue)){
        Start-Sleep -Milliseconds 200
      }
      if(Get-Process -Id $wordPid -ErrorAction SilentlyContinue){
        Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue
        $rec.forced_owned_word_cleanup=$true
      }
    }
  }
  return [pscustomobject]$rec
}

$SourceFull=[IO.Path]::GetFullPath($SourceDocx)
$OutputFull=[IO.Path]::GetFullPath($OutputJson)
if(-not(Test-Path -LiteralPath $SourceFull)){throw "Source DOCX not found: $SourceFull"}
if([string]::IsNullOrWhiteSpace($FragmentsDir)){
  $FragmentsDir=[IO.Path]::Combine(
    [IO.Path]::GetDirectoryName($OutputFull),
    [IO.Path]::GetFileNameWithoutExtension($OutputFull)+'-fragments'
  )
}
$FragmentsFull=[IO.Path]::GetFullPath($FragmentsDir)
[IO.Directory]::CreateDirectory($FragmentsFull)|Out-Null
$targets=@($Ordinals -split ',' | ForEach-Object{if($_.Trim()){[int]$_.Trim()}} | Sort-Object -Unique)
if($targets.Count -eq 0){throw 'No ordinals supplied.'}
if($targets.Count -gt 48){throw "Visual-line export is intentionally bounded to 48 reviewed ordinals."}
if($YTolerancePt -le 0 -or $YTolerancePt -gt 4){throw 'YTolerancePt must be >0 and <=4.'}
if($MaxRetries -lt 1 -or $MaxRetries -gt 3){throw 'MaxRetries must be 1..3.'}

try{
  $sourceShaBefore=(Get-SharedSha256 $SourceFull)
}catch{
  throw "Cannot hash source. Use the frozen source copy from the AxMath workspace, not a live/exclusively locked original. $($_.Exception.Message)"
}

$result=[ordered]@{
  schema='axmath-source-visual-lines/v1'
  source=$SourceFull
  source_sha256_before=$sourceShaBefore
  ordinals=$targets
  y_tolerance_pt=$YTolerancePt
  isolation='one-task-owned-word-process-per-ordinal'
  exports=@()
  failed=@()
  success=$false
}

foreach($ord in $targets){
  $last=$null
  for($attempt=1;$attempt -le $MaxRetries;$attempt++){
    $last=Invoke-ExportOne -SourceFull $SourceFull -Ordinal $ord -FragmentsFull $FragmentsFull -Tolerance $YTolerancePt
    $last | Add-Member -NotePropertyName attempt -NotePropertyValue $attempt -Force
    if($last.success){break}
    Start-Sleep -Milliseconds 350
  }
  if($last.success){$result.exports += $last}else{$result.failed += $last}
}

$sourceShaAfter=(Get-SharedSha256 $SourceFull)
$result.source_sha256_after=$sourceShaAfter
$result.source_unchanged=($sourceShaAfter -eq $sourceShaBefore)
$result.success=(
  $result.failed.Count -eq 0 -and
  $result.exports.Count -eq $targets.Count -and
  $result.source_unchanged
)
$result.finished_at=(Get-Date).ToString('o')
$result|ConvertTo-Json -Depth 9|Set-Content -LiteralPath $OutputFull -Encoding UTF8
$result|ConvertTo-Json -Depth 6
if(-not $result.success){exit 1}
