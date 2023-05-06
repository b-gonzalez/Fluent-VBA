function get-officeFiles {
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
        $srcFiles = Get-ChildItem -Path .\Source -File
        $macros = Get-ChildItem -Path .\Source -File | Where-Object {$_.Name -ne "mTodo.bas"}
        $distFiles = Get-ChildItem -Path .\Distribution -File

        foreach ($file in $srcFiles) {
            if ($file.Extension -eq ".doccls" -or $file.Extension -eq ".doccls" -or $file.Name -like "mTodo.bas") {
                $file.Delete()
            }
        }
    
        foreach ($file in $distfiles) {
            $file.delete()
        }
    
        $GUIDs = @()
        $GUIDs += $scriptingGuid
        $GUIDs += $regexGuid
        get-excel -outputPath $outputPath -macros $macros -GUIDs $GUIDs -removePersonalInfo
        $GUIDs += $excelGuid
        get-word -outputPath $outputPath -macros $macros -GUIDs $GUIDs -removePersonalInfo
        get-powerpoint -outputPath $outputPath -macros $macros -GUIDs $GUIDs -removePersonalInfo
        get-access -outputPath $outputPath -macros $macros -GUIDs $GUIDs -removePersonalInfo
    
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
}

get-officeFiles