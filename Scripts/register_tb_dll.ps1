enum bitness {
  bit32 = 1
  bit64 = 2
}

function registerDll(){
  param(
    [bitness]$dllBitness
  )
  
  $dllFilename = ""
  $curDir = $PSScriptRoot
  $parentDir = (get-item $curDir).parent.FullName
  $build = "$parentDir\Source\twin_basic\Bulid"
  Set-Location $build

  if ($dllBitness -eq [bitness]::bit32) {
    $dllFilename = "$($build)\fluent_vba_tb_win32.dll"
  } elseif ($dllBitness -eq [bitness]::bit64) {
    $dllFilename = "$($build)\fluent_vba_tb_win64.dll"
  }

  C:\Windows\System32\regsvr32.exe $dllFilename
}

registerDll -dllBitness bit64