#NOTE: The program that's running this script MUST be
#run as administrator e.g. in command prompt, PowerShell, 
#VSCode, etc. Otherwise the script will fail.

#You must also select the correct bitness for the DLL you 
#want to run. The script defaults to 64 bit. But if you 
#want to run for 32bit, just comment out the 64bit line 
#and uncomment the 32bit line at the bottom of this file.

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
  
  $curDir = $PSScriptRoot
  $parentDir = (get-item $curDir).parent.FullName
  $build = "$parentDir\Source\twin_basic\Build"
  Set-Location $build

  [string[]]$arr = @()

  if ($dllBitness -band [bitness]::bit32) {
    $arr += "$($build)\fluent_vba_tb_win32.dll"
  } 
  
  if ($dllBitness -band [bitness]::bit64) {
    $arr += "$($build)\fluent_vba_tb_win64.dll"
  }

  $sys32Dir = Join-Path $env:windir "system32\"

  Set-Location $sys32Dir

  foreach($dll in $arr) {
    regsvr32.exe $dllFilename
  }

  Set-Location $curDir
}

# registerDll -dllBitness bit32
registerDll -dllBitness bit64
# registerDll -dllBitness bit32 + bit64