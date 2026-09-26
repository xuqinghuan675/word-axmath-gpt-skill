param(
  [Parameter(Mandatory=$true)][string]$InputDocx,
  [Parameter(Mandatory=$true)][string]$OutputDocx,
  [Parameter(Mandatory=$true)][string]$Ordinals,
  [string]$TemplatePath='',
  [string]$ReportPath=''
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
$TemplatePath=Resolve-AxMathTemplate $TemplatePath
if([string]::IsNullOrWhiteSpace($ReportPath)){$ReportPath=$OutputDocx+'.inline-roundtrip.json'}
$targets=@($Ordinals -split ',' | ForEach-Object{[int]$_.Trim()} | Sort-Object -Descending -Unique)
$res=[ordered]@{input=$InputDocx;output=$OutputDocx;targets=$targets;macro_out='AMSAM2TeX';macro_in='AMSTeX2AM';repairs=@();failed=@();success=$false}
$word=$null;$doc=$null;$wordPid=$null
$before=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
if($before.Count -gt 0){throw "Refusing inline roundtrip because Word is already running: $($before -join ',')"}
if(Test-Path -LiteralPath $OutputDocx){Remove-Item -LiteralPath $OutputDocx -Force}
Copy-Item -LiteralPath $InputDocx -Destination $OutputDocx -Force

function Get-AxMathShapes($d){
  $arr=@()
  foreach($s in $d.InlineShapes){
    try{if($s.OLEFormat.ProgID -eq 'Equation.AxMath'){$arr += $s}}catch{}
  }
  return ,@($arr)
}

try{
  $word=New-Object -ComObject Word.Application
  $word.Visible=$false;$word.DisplayAlerts=0
  try{$word.ScreenUpdating=$false}catch{}
  try{$word.Options.SaveNormalPrompt=$false}catch{}
  Start-Sleep -Milliseconds 300
  $new=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Where-Object{$before -notcontains $_.Id}|Select-Object -ExpandProperty Id)
  if($new.Count -eq 1){$wordPid=[int]$new[0]}
  $found=$false
  foreach($a in $word.AddIns){if($a.Name -eq 'AxMath.dotm'){$found=$true;if(-not $a.Installed){$a.Installed=$true}}}
  if(-not $found){$word.AddIns.Add($TemplatePath,$true)|Out-Null}
  $doc=$word.Documents.Open($OutputDocx,$false,$false,$false)
  $initial=Get-AxMathShapes $doc
  $res.axmath_before=$initial.Count
  $res.omath_before=[int]$doc.OMaths.Count
  $res.paragraphs_before=[int]$doc.Paragraphs.Count

  foreach($ord in $targets){
    $rec=[ordered]@{ordinal=$ord;success=$false}
    $tmp=$null
    try{
      $ax=Get-AxMathShapes $doc
      if($ord -lt 1 -or $ord -gt $ax.Count){throw "ordinal $ord out of range 1..$($ax.Count)"}
      $target=$ax[$ord-1]
      $oldStart=[int]$target.Range.Start
      $oldEnd=[int]$target.Range.End
      $rec.before=[ordered]@{width=[double]$target.Width;height=[double]$target.Height;start=$oldStart;end=$oldEnd}

      # Convert only this OLE to LaTeX in an isolated temp document, then force
      # inline delimiter $...$ before converting back to AxMath.
      $tmp=$word.Documents.Add()
      $tmp.Range(0,0).FormattedText=$target.Range.FormattedText
      $tmpAx=Get-AxMathShapes $tmp
      if($tmpAx.Count -ne 1){throw "temporary AxMath count before=$($tmpAx.Count)"}
      $tmp.Activate();$tmp.Content.Select()
      $cb=$null
      $word.Run('AMSAM2TeX',([ref]$cb))
      $tex=$tmp.Content.Text.Trim([char]13,[char]10,[char]32,[char]9)
      if([string]::IsNullOrWhiteSpace($tex)){throw 'AxMath->TeX returned empty text'}
      if($tex.StartsWith('$$') -and $tex.EndsWith('$$') -and $tex.Length -ge 4){
        $tex='$'+$tex.Substring(2,$tex.Length-4)+'$'
      } elseif(-not ($tex.StartsWith('$') -and $tex.EndsWith('$'))){
        $tex='$'+$tex.Trim('$')+'$'
      }
      # AxMath may serialize Word NBSP as \mathrm{<NBSP>}; on round-trip that
      # can render as a visible '?' glyph. Relation operators already provide
      # their own math spacing, so remove only this serialization artifact.
      $nbsp=[char]0x00A0
      $tex=$tex.Replace(('\mathrm{' + $nbsp + '}'),'')
      $rec.tex_inline=$tex

      $tmp.Content.Text=$tex
      $tmp.Content.Select()
      $cb2=$null
      $word.Run('AMSTeX2AM',([ref]$cb2))
      $newAx=Get-AxMathShapes $tmp
      if($newAx.Count -ne 1){throw "TeX->AxMath donor count=$($newAx.Count)"}
      $donor=$newAx[0]
      $rec.donor=[ordered]@{width=[double]$donor.Width;height=[double]$donor.Height}

      # Install the official-macro-generated inline OLE + its own preview.
      $donor.Range.Copy()
      $targetRange=$doc.Range($oldStart,$oldEnd)
      $targetRange.Delete()
      $doc.Range($oldStart,$oldStart).Paste()
      $tmp.Close($false);$tmp=$null

      $ax2=Get-AxMathShapes $doc
      if($ax2.Count -ne $res.axmath_before){throw "AxMath count changed to $($ax2.Count)"}
      $replacement=$null
      foreach($s in $ax2){if([Math]::Abs([int]$s.Range.Start-$oldStart) -le 2){$replacement=$s;break}}
      if($null -eq $replacement){throw "replacement not found near $oldStart"}
      $rec.after=[ordered]@{width=[double]$replacement.Width;height=[double]$replacement.Height;start=[int]$replacement.Range.Start;end=[int]$replacement.Range.End}
      $rec.success=$true
      $res.repairs += [pscustomobject]$rec
    }catch{
      $rec.error=$_.Exception.Message
      $res.failed += [pscustomobject]$rec
      if($tmp -ne $null){try{$tmp.Close($false)}catch{};$tmp=$null}
    }
  }

  $doc.Save()
  $final=Get-AxMathShapes $doc
  $res.axmath_after=$final.Count
  $res.omath_after=[int]$doc.OMaths.Count
  $res.paragraphs_after=[int]$doc.Paragraphs.Count
  $res.success=($res.failed.Count -eq 0 -and $res.axmath_after -eq $res.axmath_before -and $res.omath_after -eq $res.omath_before -and $res.paragraphs_after -eq $res.paragraphs_before)
}catch{
  $res.error=$_.Exception.Message;$res.hresult=$_.Exception.HResult
}finally{
  if($doc -ne $null){try{$doc.Close($false)}catch{}}
  if($word -ne $null){try{$word.Quit()}catch{}}
  $doc=$null;$word=$null
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
  if($wordPid){Start-Sleep -Milliseconds 500;if(Get-Process -Id $wordPid -ErrorAction SilentlyContinue){Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue}}
  $res|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $ReportPath -Encoding UTF8
  $res|ConvertTo-Json -Depth 8
}
if(-not $res.success){exit 1}
