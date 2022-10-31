git pull

[string]$tag = git describe --tags
[string]$oldTagNumber= ([regex]"\d.+\-").match($tag).groups[0].Value
[decimal]$oldTagNumber = $oldTagNumber.Substring(0,$oldTagNumber.Length - 3)
[decimal]$newTagNumber = $oldTagNumber + .01

# Write-Output "tag is ${tag}"
# Write-Output "Old tag number is ${OldTagNumber}"
# Write-Output "Old tag number is ${newTagNumber}"

[string]$curDir = $PSScriptRoot
[string]$parentDir = (get-item $curDir).parent
[string]$OldDestination = "$parentDir\Fluent-VBA-${OldTagNumber}.zip"
[string]$NewDestination = "$parentDir\Fluent-VBA-${NewTagNumber}.zip"

if (Test-Path $OldDestination) {
  Remove-Item $OldDestination
}

$compress = @{
  Path = "$parentDir\Source", "$parentDir\Distribution"
  CompressionLevel = "Fastest"
  DestinationPath = $NewDestination
}

Compress-Archive @compress

gh release create "v${newTagNumber}" $NewDestination