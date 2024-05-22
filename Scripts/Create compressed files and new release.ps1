function get-PreviousTagNumber {
  [string]$tag = git describe --tags
  [decimal]$oldTagNumber = ([regex]"[\d|\.]+").match($tag).groups[0].Value
  Write-Output $oldTagNumber
}

function get-AndPublishPackage {
  param (
    [Parameter(Mandatory = $true)][string]$newTagNumber
  )

  try {
    git pull
  
    $curDir = $PSScriptRoot
    $parentDir = (get-item $curDir).parent.FullName
    $COF = "$parentDir\Scripts\Create Office Files.ps1"
    Import-Module $COF
  
    # [string]$tag = git describe --tags
    # [decimal]$oldTagNumber = ([regex]"[\d|\.]+").match($tag).groups[0].Value
    # [decimal]$increm = .01
    # [decimal]$newTagNumber = $oldTagNumber + $increm
    [string]$curDir = $PSScriptRoot
    [string]$parentDir = (get-item $curDir).parent
    [string]$OldDestination = "$parentDir\Fluent-VBA-${OldTagNumber}.zip"
    [string]$NewDestination = "$parentDir\Fluent-VBA-${NewTagNumber}.zip"
  
    if ($newTagNumber - $oldTagNumber -eq $increm) {
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
    else {
      Write-Output "Tag number discrepancy"
    }
  }
  catch {
    Write-Host "An error occurred:"
    Write-Host $_
  }
}

# get-PreviousTagNumbers