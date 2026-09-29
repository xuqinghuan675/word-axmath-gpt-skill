param(
  [string]$TemplatePath='',
  [string]$ReportPath=''
)
$ErrorActionPreference='Stop'

function Resolve-AxMathTemplate {
  param([string]$Requested)
  if(-not [string]::IsNullOrWhiteSpace($Requested)){
    if(Test-Path -LiteralPath $Requested){return (Resolve-Path -LiteralPath $Requested).Path}
    throw "AxMath.dotm not found: $Requested"
  }
  $candidates=@()
  if(${env:ProgramFiles(x86)}){$candidates += (Join-Path ${env:ProgramFiles(x86)} 'AxMath\MSOffice\AxMath.dotm')}
  if($env:ProgramFiles){$candidates += (Join-Path $env:ProgramFiles 'AxMath\MSOffice\AxMath.dotm')}
  foreach($candidate in $candidates){if(Test-Path -LiteralPath $candidate){return $candidate}}
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

function Get-PrimeOrders {
  param([string]$Tex)
  $orders=@()
  foreach($m in [regex]::Matches($Tex,"'''|''|\\prime\b")){
    $token=[string]$m.Value
    if($token.StartsWith('\prime')){$orders += 1}
    elseif($token.Length -eq 2){$orders += 2}
    elseif($token.Length -eq 3){$orders += 3}
  }
  return @($orders)
}

$TemplatePath=Resolve-AxMathTemplate $TemplatePath
if([string]::IsNullOrWhiteSpace($ReportPath)){$ReportPath=Join-Path $PWD 'AXMATH_PRIME_CONTRACT_PROBE.json'}
$ReportFull=[IO.Path]::GetFullPath($ReportPath)
$beforeWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$result=[ordered]@{
  schema='axmath-prime-contract-probe/v2'
  canonical_contract=[ordered]@{first='\prime';second="''";third="'''"}
  tests=@()
  raw_unicode_triple_test=$null
  success=$false
}
$word=$null;$wordPid=$null;$owned=$false
try{
  $word=New-Object -ComObject Word.Application
  Start-Sleep -Milliseconds 350
  $newWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object{$beforeWord -notcontains $_.Id} | Select-Object -ExpandProperty Id)
  if($newWord.Count -eq 1){$wordPid=[int]$newWord[0];$owned=$true}else{throw 'Could not establish a distinct task-owned Word process.'}
  $word.Visible=$false;$word.DisplayAlerts=0
  $found=$false
  foreach($a in $word.AddIns){if($a.Name -eq 'AxMath.dotm'){$found=$true;if(-not $a.Installed){$a.Installed=$true}}}
  if(-not $found){$word.AddIns.Add($TemplatePath,$true)|Out-Null}

  $tests=@(
    [pscustomobject]@{name='first_prime';tex='$F\prime(x)$';expected='1'},
    [pscustomobject]@{name='second_prime';tex='$F''''(x)$';expected='2'},
    [pscustomobject]@{name='third_prime';tex='$F''''''(\xi)$';expected='3'},
    [pscustomobject]@{name='prime_then_square';tex='${y\prime}^{2}$';expected='1'}
  )
  foreach($t in $tests){
    $doc=$word.Documents.Add()
    try{
      $doc.Content.Text=$t.tex;$doc.Content.Select();$cb=$null;$word.Run('AMSTeX2AM',([ref]$cb))
      $ax=@(Get-AxMathShapes $doc)
      if($ax.Count -ne 1){throw "$($t.name): expected one Equation.AxMath, got $($ax.Count)"}
      $doc.Content.Select();$cb2=$null;$word.Run('AMSAM2TeX',([ref]$cb2))
      $round=[string]$doc.Content.Text.Trim([char]13,[char]10,[char]32,[char]9)
      $orders=@(Get-PrimeOrders $round)
      $pass=(($orders -join ',') -eq $t.expected -and -not $round.Contains('?'))
      $result.tests += [pscustomobject]@{name=$t.name;input=$t.tex;roundtrip=$round;orders=$orders;pass=$pass}
      if(-not $pass){throw "$($t.name): prime contract mismatch: $round"}
    }finally{try{$doc.Close($false)}catch{}}
  }

  $bad=$word.Documents.Add()
  try{
    $bad.Content.Text='$F‴(\xi)$';$bad.Content.Select();$cb3=$null;$word.Run('AMSTeX2AM',([ref]$cb3))
    $badAx=@(Get-AxMathShapes $bad)
    $badRound=''
    if($badAx.Count -eq 1){$bad.Content.Select();$cb4=$null;$word.Run('AMSAM2TeX',([ref]$cb4));$badRound=[string]$bad.Content.Text.Trim([char]13,[char]10,[char]32,[char]9)}
    $badOrders=@(Get-PrimeOrders $badRound)
    $rawTripleUnsafe=($badAx.Count -ne 1 -or $badRound.Contains('?') -or (($badOrders -join ',') -ne '3'))
    $result.raw_unicode_triple_test=[ordered]@{input='$F‴(\xi)$';axmath_count=$badAx.Count;roundtrip=$badRound;orders=$badOrders;unsafe=$rawTripleUnsafe}
    if(-not $rawTripleUnsafe){throw 'Raw U+2034 unexpectedly matched the verified third-prime contract; re-review the normalizer before changing production rules.'}
  }finally{try{$bad.Close($false)}catch{}}

  $result.success=(@($result.tests|Where-Object{-not $_.pass}).Count -eq 0 -and [bool]$result.raw_unicode_triple_test.unsafe)
}catch{
  $result.error=$_.Exception.Message
  $result.hresult=$_.Exception.HResult
}finally{
  if($word -and $owned){try{$word.Quit()}catch{}}
  $word=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
  if($wordPid -and $owned){Start-Sleep -Milliseconds 400;if(Get-Process -Id $wordPid -ErrorAction SilentlyContinue){Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue}}
  $result|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $ReportFull -Encoding UTF8
  $result|ConvertTo-Json -Depth 8
}
if(-not $result.success){exit 1}
