#NOTE: The program that's running this script MUST be
#run as administrator e.g. in command prompt, PowerShell, 
#VSCode, etc. Otherwise the script will fail.

#This code can be run from either the bin or the scripts
#folder. To run from the scripts folder, you MUST add the
#-runFromScriptsDir switch parameter. You can see an
#example commented out towards the bottom.

#You must also select the correct bitness for the DLL you 
#want to run. The script defaults to 64-bit. This is due
#to the function call using a bit64 parameter at the 
#bottom of this file. If you want to run for 32-bit, 
#just comment out the ([bitness]::bit64) line and uncomment 
#the ([bitness]::bit32) line. If you want to run for both 
#32-bit and 64-bit, then comment out the 64-bit line and 
#uncomment the line with ([bitness]::bit32 + [bitness]::bit64)

#If run successfully, you will get a msgbox notification
#that dllRegisterServer succeeded. After that, go to tools
#> References in the VBIDE and select the DLL file. After
#that you should be able to use code within the library.

[Flags()] enum bitness {
  bit32 = 1
  bit64 = 2
}

function registerDll(){
  param(
    [bitness]$dllBitness,
    [switch]$runFromScriptsDir
  )
  
  $curDir = $PSScriptRoot
  $parentDir = (get-item $curDir).parent.FullName
  
  if ($runFromScriptsDir) {
    $build = "$parentDir\Source\twin_basic\Build"
  } else {
    $build = $curDir
  }

  Set-Location $build

  if ($dllBitness -band [bitness]::bit32) {
    $arr += "$($build)\fluent_vba_tb_win32.dll"
  } 
  
  if ($dllBitness -band [bitness]::bit64) {
    $arr += "$($build)\fluent_vba_tb_win64.dll"
  }

  $sys32Dir = Join-Path $env:windir "system32\"

  Set-Location $sys32Dir

  foreach($dll in $arr) {
    regsvr32.exe $dll
  }

  Set-Location $curDir
}

# registerDll -dllBitness ([bitness]::bit32)
registerDll -dllBitness ([bitness]::bit64)
# registerDll -dllBitness ([bitness]::bit32 + [bitness]::bit64)

# registerDll -dllBitness ([bitness]::bit64) -runFromScriptsDir