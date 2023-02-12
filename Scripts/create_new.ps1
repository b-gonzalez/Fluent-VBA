try {
    $curDir = $PSScriptRoot
    $parentDir = (get-item $curDir).parent.FullName
    Set-Location $parentDir
    $outputPath = "$parentDir\Distribution\Fluent VBA"
    # Write-Output $parentDir
    # $outputPath = "$parentDir\Distribution\"
    # Write-Output $outputPath
    
    # $guidObj = Get-ExcelGuid
    # $GuidStr = Out-String -NoNewline -InputObject $guidObj
    # $xlGuid = $GuidStr.Replace("System.__ComObject","")


    $ScriptingGuid = "{420B2830-E718-11CF-893D-00A0C9054228}"
    $macros = Get-ChildItem -Path .\Source -File
    # $distFiles = Get-ChildItem -Path .\Distribution -File

    foreach ($file in $distfiles) {
        $file.delete()
    }

    get-excel -outputPath $outputPath -macros $macros -ScriptingGuid $ScriptingGuid
}

Catch {
    Write-Host "An error occurred:"
    Write-Host $_
}

Finally {
    # $workbook.RemovePersonalInformation = $true
    # $doc.Save()
    # $doc.RemovePersonalInformation = $true
    # $presentation.RemovePersonalInformation = $true
    # $acc.CurrentProject.RemovePersonalInformation = $true

    # $workbook.Save()
    # $presentation.Save()
    

    # $workbook.Close()
    # $doc.Close()
    # $presentation.Close()
    # $acc.CloseCurrentDatabase()
    
    # $excel.Quit()
    # $word.Quit()
    # $powerpoint.Quit()
    # $acc.Quit()

    # [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    # [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
    # [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation)
    # [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint)
    # [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($acc)
    # [GC]::Collect()

    # $fluentName = "Fluent VBA 1.65"
    # $distFiles = Get-ChildItem -Path .\Distribution -File

    # foreach ($file in $distfiles) {
    #     if ($file.FullName -like '*~$uent*') {
    #         Remove-Item $file.FullName -Force
    #     } else {
    #         Rename-Item $file.FullName -NewName "$($fluentName)$($file.Extension)"
    #     }
    # }
}

function get-excel {
    param (
        [string]$outputPath,
        [string]$macros,
        [string]$ScriptingGuid
    )
    $excel = New-Object -ComObject excel.application
    $workbook = $excel.Workbooks.Add()
    $xlOpenXMLWorkbookMacroEnabled = 52
    $workbook.SaveAs($outputPath, $xlOpenXMLWorkbookMacroEnabled)
    # $macros = Get-ChildItem -Path .\Source -File

    $Major = 0
    $Minor = 0

    foreach ($macro in $macros) {
        if ($macro.Extension -ne ".doccls" -and $macro.Extension -ne ".ps1" -and $macro.BaseName -ne "mTODO") {
            $workbook.VBProject.VBComponents.Import($macro.FullName) | Out-Null
        }
    }

    $workbook.VBProject.References.AddFromGuid($ScriptingGuid,$Major, $Minor)
    $workbook.Save()
    $workbook.Close()
    $excel.Quit()
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
}

function get-word {
    param (
        [string]$outputPath,
        [string]$ExcelGuid,
        [string]$ScriptingGuid,
        [string]$minorVersion,
        [string]$majorVersion,
        [string]$macros
    )

    $distFiles = Get-ChildItem -Path .\Distribution -File

    $word = New-Object -ComObject word.application
    $doc = $word.documents.add()
    $wdFormatFlatXMLMacroEnabled = 13
    $doc.SaveAs($outputPath,$wdFormatFlatXMLMacroEnabled)
    $macros = Get-ChildItem -Path .\Source -File

    $Major = 0
    $Minor = 0

    foreach ($macro in $macros) {
        if ($macro.Extension -ne ".doccls" -and $macro.Extension -ne ".ps1" -and $macro.BaseName -ne "mTODO") {
            $doc.VBProject.VBComponents.Import($macro.FullName)
        }
    }

    $doc.VBProject.References.AddFromGuid($xlGuid,$Major, $Minor)
    $doc.VBProject.References.AddFromGuid($ScriptingGuid,$Major, $Minor)
}

function get-powerpoint {
    $powerpoint = New-Object -ComObject powerpoint.application
    $presentation = $powerpoint.Presentations.Add()
    $ppSaveAsOpenXMLPresentationMacroEnabled = 25
    $presentation.SaveAs($outputPath,$ppSaveAsOpenXMLPresentationMacroEnabled)
    $macros = Get-ChildItem -Path .\Source -File

    $Major = 0
    $Minor = 0

    foreach ($macro in $macros) {
        if ($macro.Extension -ne ".doccls" -and $macro.Extension -ne ".ps1" -and $macro.BaseName -ne "mTODO") {
            $presentation.VBProject.VBComponents.Import($macro.FullName)
        }
    }

    $presentation.VBProject.References.AddFromGuid($xlGuid,$Major, $Minor)
    $presentation.VBProject.References.AddFromGuid($ScriptingGuid,$Major, $Minor)
}

function get-access {
    $acc = New-Object -ComObject Access.Application
    $acFileFormatAccess2007 = 12
    $acc.NewCurrentDataBase($outputPath,$acFileFormatAccess2007)

    $acCmdCompileAndSaveAllModules = 126
    $acModule = 5
    $macros = Get-ChildItem -Path .\Source -File

    $Major = 0
    $Minor = 0

    foreach ($macro in $macros) {
        if ($macro.Extension -ne ".doccls" -and $macro.Extension -ne ".ps1" -and $macro.BaseName -ne "mTODO") {
            $acc.VBE.ActiveVBProject.VBComponents.Import($macro.FullName)
            $acc.VBE.ActiveVBProject.VBComponents($acc.VBE.ActiveVBProject.VBComponents.Count).Name = $macro.BaseName
            $acc.DoCmd.RunCommand($acCmdCompileAndSaveAllModules)
            $acc.DoCmd.Save($acModule, $macro.BaseName)
        }
    }

    $acc.VBE.ActiveVBProject.References.AddFromGuid($xlGuid,$Major, $Minor)
    $acc.VBE.ActiveVBProject.References.AddFromGuid($ScriptingGuid,$Major, $Minor)
}