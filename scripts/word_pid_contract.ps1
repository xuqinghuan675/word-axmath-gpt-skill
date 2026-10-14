if(-not ('AxMathWordPidNative' -as [type])){
  Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class AxMathWordPidNative {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

function Resolve-OwnedWordIdentity {
  param($Word,[int[]]$BeforePids,[datetime]$LaunchStartedAt)
  [uint32]$pidValue=0
  $proof='com_hwnd'
  try{
    if($null -ne $Word.Hwnd){
      [void][AxMathWordPidNative]::GetWindowThreadProcessId([IntPtr]$Word.Hwnd,[ref]$pidValue)
    }
  }catch{}
  if($pidValue -eq 0){
    # Headless Word can have no HWND. Accept only one newly created WINWORD
    # whose process creation time is inside this COM launch window.
    $new=@(Get-Process WINWORD -ErrorAction SilentlyContinue|
      Where-Object{$BeforePids -notcontains $_.Id})
    if($new.Count -ne 1){throw 'Could not uniquely resolve a newly created Word process.'}
    $pidValue=[uint32]$new[0].Id
    $proof='unique_new_pid_and_creation_time'
  }
  $wordPid=[int]$pidValue
  if($BeforePids -contains $wordPid){throw 'Word COM attached to a pre-existing user process.'}
  $process=Get-Process -Id $wordPid -ErrorAction Stop
  if($process.ProcessName -ne 'WINWORD' -or $process.StartTime -lt $LaunchStartedAt.AddSeconds(-2)){
    throw 'Word COM PID or process creation time does not prove task ownership.'
  }
  return [pscustomobject]@{
    pid=$wordPid
    start_utc_ticks=$process.StartTime.ToUniversalTime().Ticks
    started_at=$process.StartTime.ToString('o')
    proof=$proof
  }
}

function Test-OwnedWordIdentity {
  param($Identity)
  if($null -eq $Identity){return $false}
  $process=Get-Process -Id ([int]$Identity.pid) -ErrorAction SilentlyContinue
  return [bool](
    $process -and $process.ProcessName -eq 'WINWORD' -and
    $process.StartTime.ToUniversalTime().Ticks -eq [long]$Identity.start_utc_ticks
  )
}
