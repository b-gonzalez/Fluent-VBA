#NOTE: the program that's running this script must be
#run as administrator e.g. in command prompt,PowerShell, 
#VSCode, etc. You must also select the correct bitness
#for the DLL you want to run. The script defaults to
#64 bit. But if you want to run for 32bit, just comment
#out the 64bit line and uncomment the 32bit line at the
#bottom of the file.

#If run successfully, you will get a msgbox notification
#that dllRegisterServer succeeded. After that, go to tools
#> References in the VBIDE and select the DLL file. After
#that you should be able to use code within the library.

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
  $build = "$parentDir\Source\twin_basic\Build"
  Set-Location $build

  if ($dllBitness -eq [bitness]::bit32) {
    $dllFilename = "$($build)\fluent_vba_tb_win32.dll"
  } elseif ($dllBitness -eq [bitness]::bit64) {
    $dllFilename = "$($build)\fluent_vba_tb_win64.dll"
  }

  C:\Windows\System32\regsvr32.exe $dllFilename
}

# registerDll -dllBitness bit32
registerDll -dllBitness bit64