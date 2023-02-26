git pull

[string]$tag = git describe --tags
[decimal]$oldTagNumber= ([regex]"[\d|\.]+").match($tag).groups[0].Value
[decimal]$increm = .01
[decimal]$newTagNumber = $oldTagNumber + $increm
[string]$curDir = $PSScriptRoot
[string]$parentDir = (get-item $curDir).parent
[string]$OldDestination = "$parentDir\Fluent-VBA-${OldTagNumber}.zip"
[string]$NewDestination = "$parentDir\Fluent-VBA-${NewTagNumber}.zip"

if ($newTagNumber - $oldTagNumber -eq $increm) {
  if (Test-Path $OldDestination) {
    Remove-Item $OldDestination
  }

  # if (Test-Path $NewDestination) {
  #   Remove-Item $NewDestination
  # }
  
  $compress = @{
    Path = "$parentDir\Source", "$parentDir\Distribution"
    CompressionLevel = "Fastest"
    DestinationPath = $NewDestination
  }
  
  Compress-Archive @compress
  
  gh release create "v${newTagNumber}" $NewDestination
} else {
  Write-Output "Tag number discrepancy"
}