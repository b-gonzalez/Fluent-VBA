function get-TagNumbers {
  [string]$latestTag = git describe --abbrev=0 --tags
  return $latestTag
  # [string]$secondLatestTag = git describe --abbrev=0 --tags --exclude="$(git describe --abbrev=0 --tags)"
  # Write-Output "Second latest tag number: $secondLatestTag"
}

function get-binDir(){
  [string]$binPath = (Get-Item -Path ".\bin").FullName
  [System.Object[]]$binDir = Get-ChildItem -Path $binPath
  
  if ($binDir.Count -gt 0) {
    [System.Object[]]$dllFiles = Get-ChildItem -Path $binDir | Where-Object {$_.extension -in @(".dll")}
    [System.Object[]]$psFiles = Get-ChildItem -Path $binDir | Where-Object {$_.extension -in @(".ps1")}
    
    if (!($dllFiles.count -eq 2 -and $psFiles.count -eq 1)){
      throw "File structure incorrect. Delete files in directory and try again"
    }
  } else {
    Copy-Item -Path ".\Source\twin_basic\Build\fluent_vba_tb_win32.dll" -Destination $binPath
    Copy-Item -Path ".\Source\twin_basic\Build\fluent_vba_tb_win64.dll" -Destination $binPath
    Copy-Item -Path ".\Scripts\register_tb_dll.ps1" -Destination $binPath
  }
}

function get-AndPublishPackage {
  param (
    [Parameter(Mandatory = $true)][string]$oldTagNumber,
    [Parameter(Mandatory = $true)][string]$newTagNumber
  )

  $curDir = $PSScriptRoot
  $parentDir = (get-item $curDir).parent.FullName
  $COF = "$parentDir\Scripts\Create Office Files.ps1"
  Import-Module $COF

  try {
    git pull
  
    [string]$curDir = $PSScriptRoot
    [string]$parentDir = (get-item $curDir).parent
    [string]$OldDestination = "$parentDir\Fluent-VBA-${OldTagNumber}.zip"
    [string]$NewDestination = "$parentDir\Fluent-VBA-${NewTagNumber}.zip"
  
    if (Test-Path $OldDestination) {
      Remove-Item $OldDestination
    }
  
    if (Test-Path $NewDestination) {
      Remove-Item $NewDestination
    }
  
    get-officeFiles

    get-binDir
      
    $compress = @{
      Path             = "$parentDir\Distribution", "$parentDir\test_files",  "$parentDir\bin"
      CompressionLevel = "Fastest"
      DestinationPath  = $NewDestination
    }
      
    Compress-Archive @compress
      
    # gh release create "v${newTagNumber}" $NewDestination
  }
  catch {
    Write-Host "An error occurred:"
    Write-Host $_
  }
}

$lastTagNum = get-TagNumbers
Write-Output "Latest tag number: $lastTagNum"
get-AndPublishPackage -oldTagNumber $lastTagNum -newTagNumber "2.5.0"