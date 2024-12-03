#NOTE: The program that's running this script MUST be
#run as administrator e.g. in command prompt, PowerShell, 
#VSCode, etc. Otherwise the script will fail.

#You must also select the correct bitness for the DLL you 
#want to run. The script defaults to 64-bit. This is due
#to the function call using a bit64 parameter at the 
#bottom of this file. If you want to run for 32-bit, 
#just comment out the bit64 line and uncomment the bit32 
#line. If you want to run for both 32-bit and 64-bit, then
#comment out the 64-bit line and uncomment the line with 
#bit32 + bit64 

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
    [bitness[]]$bitnessArr
  )

  [int]$dllInt = 0

  foreach($bitness in $bitnessArr) {
    $dllInt += [int]([int][bitness]::$bitness)
  }
  
  $curDir = $PSScriptRoot
  $parentDir = (get-item $curDir).parent.FullName
  $build = "$parentDir\Source\twin_basic\Build"
  Set-Location $build

  [string[]]$arr = @()

  if ($dllInt -band [bitness]::bit32) {
    $arr += "$($build)\fluent_vba_tb_win32.dll"
  } 
  
  if ($dllInt -band [bitness]::bit64) {
    $arr += "$($build)\fluent_vba_tb_win64.dll"
  }

  $sys32Dir = Join-Path $env:windir "system32\"

  Set-Location $sys32Dir

  foreach($dll in $arr) {
    regsvr32.exe $dll
  }

  Set-Location $curDir
}

# registerDll -bitnessArr bit32
registerDll -bitnessArr bit64
# registerDll -bitnessArr bit32, bit64