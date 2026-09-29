$ErrorActionPreference='Stop'

function Assert-CanonicalPrimeTeX {
  param([string]$Tex,[string]$Label='formula')
  $noncanonicalPrimeChars=@(
    [char]0x2032,
    [char]0x2033,
    [char]0x2034,
    [char]0x2057,
    [char]0x02B9,
    [char]0x02BA,
    [char]0x2019,
    [char]0x2018
  )
  foreach($ch in $noncanonicalPrimeChars){
    if($Tex.Contains([string]$ch)){
      throw "noncanonical_prime_literal for $Label. Normalize with scripts\normalize_axmath_tex.py before AMSTeX2AM."
    }
  }
  if($Tex -match "(?<!')'(?!')"){
    throw "noncanonical_single_prime_apostrophe for $Label. AxMath 2.7.0.58 canonical first-prime syntax is \prime."
  }
  if($Tex -match "'{4,}"){
    throw "unverified_prime_order for $Label. Only verified AxMath prime orders 1..3 may be auto-rebuilt."
  }
  if($Tex -match '[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻]'){
    throw "unicode_superscript_literal for $Label. Normalize Unicode superscripts before AMSTeX2AM."
  }
  if($Tex -match "\\prime(?=[A-Za-z0-9])"){
    throw "fused_prime_exponent for $Label. Normalize OMML-style fused forms such as {f}^{\\prime2} to {f\\prime}^{2}."
  }
  if($Tex -match "(?:\\prime\b|'{2,3})\s*\^(?:\{|[A-Za-z0-9])"){
    throw "ungrouped_primed_atom_exponent for $Label. Group the primed atom before applying another exponent."
  }
}

function Get-CanonicalPrimeOrders {
  param([string]$Tex)
  $orders=@()
  foreach($m in [regex]::Matches($Tex,"'''|''|\\prime\b")){
    $token=[string]$m.Value
    if($token.StartsWith('\prime')){$orders += 1}
    elseif($token.Length -eq 2){$orders += 2}
    elseif($token.Length -eq 3){$orders += 3}
  }
  return $orders
}

function Assert-AxMathPrimeRoundTrip {
  param([string]$ApprovedTex,[string]$RoundTripTex,[string]$Label='formula')
  $expected=@(Get-CanonicalPrimeOrders $ApprovedTex)
  if($expected.Count -eq 0){return}
  if($RoundTripTex.Contains('?')){
    throw "axmath_prime_roundtrip_corrupt for $($Label): AxMath serialized '?' instead of approved prime semantics."
  }
  $actual=@(Get-CanonicalPrimeOrders $RoundTripTex)
  if(($expected -join ',') -ne ($actual -join ',')){
    throw "axmath_prime_roundtrip_mismatch for $($Label): expected [$($expected -join ',')], got [$($actual -join ',')]."
  }
}

function Verify-AxMathPrimeDonor {
  param(
    [Parameter(Mandatory=$true)]$Word,
    [Parameter(Mandatory=$true)]$Donor,
    [Parameter(Mandatory=$true)][string]$ApprovedTex,
    [string]$Label='formula'
  )
  Assert-CanonicalPrimeTeX -Tex $ApprovedTex -Label $Label
  $expected=@(Get-CanonicalPrimeOrders $ApprovedTex)
  if($expected.Count -eq 0){
    return [pscustomobject]@{verified=$false;prime_orders=@();roundtrip_tex=$null}
  }
  $verifyTmp=$null
  try{
    $verifyTmp=$Word.Documents.Add()
    $Donor.Range.Copy()
    $verifyTmp.Range(0,0).Paste()
    $col=$verifyTmp.InlineShapes
    $ax=@()
    for($i=1;$i -le [int]$col.Count;$i++){
      $s=$col.Item($i)
      try{if([string]$s.OLEFormat.ProgID -eq 'Equation.AxMath'){$ax += $s}}catch{}
    }
    if($ax.Count -ne 1){throw "prime verification copy count=$($ax.Count) for $Label"}
    $verifyTmp.Content.Select()
    $cb=$null
    $Word.Run('AMSAM2TeX',([ref]$cb))
    $roundTrip=[string]$verifyTmp.Content.Text.Trim([char]13,[char]10,[char]32,[char]9)
    Assert-AxMathPrimeRoundTrip -ApprovedTex $ApprovedTex -RoundTripTex $roundTrip -Label $Label
    return [pscustomobject]@{verified=$true;prime_orders=@($expected);roundtrip_tex=$roundTrip}
  }finally{
    if($verifyTmp -ne $null){try{$verifyTmp.Close($false)}catch{}}
  }
}
