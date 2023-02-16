function get-word {
    param (
        [string]$outputPath,
        [Object[]]$macros,
        [string]$scriptingGuid,
        [string]$excelGuid
    )

    try {
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
    
        $doc.VBProject.References.AddFromGuid($excelGuid,$Major, $Minor)
        $doc.VBProject.References.AddFromGuid($ScriptingGuid,$Major, $Minor)
    }  catch {
        Write-Host "An error occurred:"
        Write-Host $_
    } finally {
    $doc.Save()
    $doc.RemovePersonalInformation = $true
    $doc.Close()
    $word.Quit()

    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
    [GC]::Collect()
    }

}

function get-powerpoint {
    param (
        [string]$outputPath,
        [Object[]]$macros,
        [string]$ScriptingGuid,
        [string]$excelGuid
    )

    try {
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
    
        $presentation.VBProject.References.AddFromGuid($excelGuid,$Major, $Minor)
        $presentation.VBProject.References.AddFromGuid($ScriptingGuid,$Major, $Minor)
    }  catch {
        Write-Host "An error occurred:"
        Write-Host $_
    } finally {
        $presentation.RemovePersonalInformation = $true
        $presentation.Save()
        $presentation.Close()
        $powerpoint.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation)
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint)
        [GC]::Collect()
    }
}

function get-access {
    param (
        [string]$outputPath,
        [Object[]]$macros,
        [string]$ScriptingGuid,
        [string]$excelGuid
    )

    try {
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
    
        $acc.VBE.ActiveVBProject.References.AddFromGuid($excelGuid,$Major, $Minor)
        $acc.VBE.ActiveVBProject.References.AddFromGuid($ScriptingGuid,$Major, $Minor)
    }   catch {
        Write-Host "An error occurred:"
        Write-Host $_
    } finally {
        $acc.CurrentProject.RemovePersonalInformation = $true
        $acc.CloseCurrentDatabase()
        $acc.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($acc)
        [GC]::Collect()
    }
}

function get-excel {
    param (
        [string]$outputPath,
        [Object[]]$macros,
        [string]$ScriptingGuid
    )
    try {
        $excel = New-Object -ComObject excel.application
        $workbook = $excel.Workbooks.Add()
        $xlOpenXMLWorkbookMacroEnabled = 52
        $workbook.SaveAs($outputPath, $xlOpenXMLWorkbookMacroEnabled)
        # $macros = Get-ChildItem -Path .\Source -File
    
        $Major = 0
        $Minor = 0
    
        foreach ($macro in $macros) {
            if ($macro.Extension -ne ".doccls" -and $macro.Extension -ne ".ps1" -and $macro.BaseName -ne "mTODO") {
                # $workbook.VBProject.VBComponents.Import($macro.FullName)
                $workbook.VBProject.VBComponents.Import($macro.FullName) | Out-Null
            }
        }
    
        $workbook.VBProject.References.AddFromGuid($ScriptingGuid,$Major, $Minor)
    } catch {
        Write-Host "An error occurred:"
        Write-Host $_
    } finally {
        $workbook.Save()
        $workbook.Close()
        $excel.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        [GC]::Collect()
    }
}


function Get-ExcelGuid {
    try {
        $excel = New-Object -ComObject excel.application
        # $excel.Application.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $GUID = "{0002E157-0000-0000-C000-000000000046}"
        $Major = 0
        $Minor = 0

        $workbook.VBProject.References.AddFromGuid($GUID,$Major, $Minor)
        $vbe = $excel.application.VBE
        $vbProj = $vbe.ActiveVBProject
        $references = $vbProj.References

        foreach ($ref in $references) {
            if ($ref.name -like "*Excel*") {
                $guidObj = $ref.GUID
                break
            }
        }

        return $guidObj
    }

    Catch {
        Write-Host "An error occurred:"
        Write-Host $_
    }

    Finally {
        $excel.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
        [GC]::Collect()
    }
}

try {
    $curDir = $PSScriptRoot
    $parentDir = (get-item $curDir).parent.FullName
    $outputPath = "$parentDir\Distribution\Fluent VBA"
    
    Set-Location $parentDir
    
    $guidObj = Get-ExcelGuid
    $GuidStr = Out-String -NoNewline -InputObject $guidObj
    $excelGuid = $GuidStr.Replace("System.__ComObject","")

    $ScriptingGuid = "{420B2830-E718-11CF-893D-00A0C9054228}"
    $macros = Get-ChildItem -Path .\Source -File
    $distFiles = Get-ChildItem -Path .\Distribution -File

    foreach ($file in $distfiles) {
        $file.delete()
    }

    get-word -outputPath $outputPath -macros $macros -ScriptingGuid $ScriptingGuid -excelGuid $excelGuid
    get-powerpoint -outputPath $outputPath -macros $macros -ScriptingGuid $ScriptingGuid -excelGuid $excelGuid
    get-access -outputPath $outputPath -macros $macros -ScriptingGuid $ScriptingGuid -excelGuid $excelGuid
    get-excel -outputPath $outputPath -macros $macros -ScriptingGuid $ScriptingGuid

    foreach ($file in $distfiles) {
        if ($file.FullName -like '*~$uent*') {
            Remove-Item $file.FullName -Force
        } else {
            Rename-Item $file.FullName -NewName "$($fluentName)$($file.Extension)"
        }
    }
}

Catch {
    Write-Host "An error occurred:"
    Write-Host $_
}