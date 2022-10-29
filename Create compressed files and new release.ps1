git pull

[string]$curDir = $PSScriptRoot
[string]$parentDir = (get-item $curDir).parent
[string]$OldDestination = "$parentDir\Fluent-VBA-${OldTagNumber}.zip"
[string]$NewDestination = "$parentDir\Fluent-VBA-${NewTagNumber}.zip"
[string]$tag = git describe --tags
[decimal]$OldTagNumber = $tag.substring(1,$tag.Length -1)

$newTagNumber = $tagNumber + .01

if (Test-Path $OldDestination) {
  Remove-Item $OldDestination
}

$compress = @{
  Path = "$parentDir\Source", "$parentDir\Distribution"
  CompressionLevel = "Fastest"
  DestinationPath = $NewDestination
}

Compress-Archive @compress

# gh release create "v${newTagNumber}"