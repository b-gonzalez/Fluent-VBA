function get-officeFiles {
    try {
        $curDir = $PSScriptRoot
        $parentDir = (get-item $curDir).parent.FullName
        # $outputPath = "$parentDir\Distribution\Fluent VBA"
        # $outputPath = "$parentDir"
        $functions = "$parentDir\Scripts\functions.ps1"
        Import-Module $functions
        
        Set-Location $parentDir
        
        $guidObj = Get-ExcelGuid
        $guidStr = Out-String -NoNewline -InputObject $guidObj
        $excelGuid = $GuidStr.Replace("System.__ComObject", "")
    
        $scriptingGuid = "{420B2830-E718-11CF-893D-00A0C9054228}"        
        $srcFiles = Get-ChildItem -Path .\Source -File
        $macros = $srcFiles | Where-Object { $_.Name -ne "mTodo.bas" }

        if (! (Test-Path -PathType container .\Distribution)) {
            New-Item -Path .\Distribution -ItemType "directory"
        }

        if (! (Test-Path -PathType container .\test_files)) {
            New-Item -Path .\test_files -ItemType "directory"
        }

        $testFiles = Get-ChildItem -Path .\test_files -File

        $distFiles = Get-ChildItem -Path .\Distribution -File

        foreach ($file in $srcFiles) {
            if ($file.Extension -eq ".doccls" -or $file.Extension -eq ".doccls" -or $file.Name -like "mTodo.bas") {
                $file.Delete()
            }
        }
    
        foreach ($file in $distfiles) {
            $file.delete()
        }

        foreach ($file in $testFiles) {
            $file.delete()
        }
    
        $GUIDs = @()
        $GUIDs += $scriptingGuid

        #This function requires "Trust access to the VBA project object model" to be enabled in Excel
        get-excel -outputPath $parentDir -macros $macros -GUIDs $GUIDs -removePersonalInfo
        $GUIDs += $excelGuid

        # #This function requires "Trust access to the VBA project object model" to be enabled in Word
        get-word -outputPath $parentDir -macros $macros -GUIDs $GUIDs -removePersonalInfo

        # #This function requires "Trust access to the VBA project object model" to be enabled in PowerPoint
        get-powerpoint -outputPath $parentDir -macros $macros -GUIDs $GUIDs -removePersonalInfo

        # #"Trust access to the VBA project object model" is not required for this function
        get-access -outputPath $parentDir -macros $macros -GUIDs $GUIDs -removePersonalInfo
    
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