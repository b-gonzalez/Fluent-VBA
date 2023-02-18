try {
    $curDir = $PSScriptRoot
    $parentDir = (get-item $curDir).parent.FullName
    $outputPath = "$parentDir\Distribution\Fluent VBA"
    $functions = "$parentDir\Scripts\functions.ps1"
    Import-Module $functions
    
    Set-Location $parentDir
    
    $guidObj = Get-ExcelGuid
    $guidStr = Out-String -NoNewline -InputObject $guidObj
    $excelGuid = $GuidStr.Replace("System.__ComObject","")

    $scriptingGuid = "{420B2830-E718-11CF-893D-00A0C9054228}"
    $macros = Get-ChildItem -Path .\Source -File
    $distFiles = Get-ChildItem -Path .\Distribution -File

    foreach ($file in $distfiles) {
        $file.delete()
    }

    $GUIDs = @()
    $GUIDs += $scriptingGuid
    get-excel -outputPath $outputPath -macros $macros -GUIDs $GUIDs -removePersonalInfo $true
    $GUIDs += $excelGuid
    get-word -outputPath $outputPath -macros $macros -GUIDs $GUIDs -removePersonalInfo $true
    get-powerpoint -outputPath $outputPath -macros $macros -GUIDs $GUIDs -removePersonalInfo $true
    get-access -outputPath $outputPath -macros $macros -GUIDs $GUIDs -removePersonalInfo $true

    foreach ($file in $distfiles) {
        if ($file.FullName -like '*~$uent*') {
            Remove-Item $file.FullName -Force
        }
    }
}

Catch {
    Write-Host "An error occurred:"
    Write-Host $_
}