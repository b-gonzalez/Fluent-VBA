function get-TagNumbers {
  [string]$latestTag = git describe --abbrev=0 --tags
  return $latestTag
  # [string]$secondLatestTag = git describe --abbrev=0 --tags --exclude="$(git describe --abbrev=0 --tags)"
  # Write-Output "Second latest tag number: $secondLatestTag"
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
      
    $compress = @{
      Path             = "$parentDir\Distribution", "$parentDir\test_files"
      CompressionLevel = "Fastest"
      DestinationPath  = $NewDestination
    }
      
    Compress-Archive @compress
      
    gh release create "v${newTagNumber}" $NewDestination
  }
  catch {
    Write-Host "An error occurred:"
    Write-Host $_
  }
}

$lastTagNum = get-TagNumbers
Write-Output "Latest tag number: $lastTagNum"
get-AndPublishPackage -oldTagNumber $lastTagNum -newTagNumber "2.4.1"