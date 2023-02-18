function get-word {
    param (
        [Parameter(Mandatory=$true)][string]$outputPath,
        [Parameter(Mandatory=$true)][Object[]]$macros,
        [Parameter(Mandatory=$false)][string[]]$GUIDs,
        [Parameter(Mandatory=$false)][bool]$removePersonalInfo
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
        
        if ($PSBoundParameters.ContainsKey('GUIDs')) {
            foreach ($GUID in $GUIDs) {
                $doc.VBProject.References.AddFromGuid($GUID,$Major, $Minor)
            }
        }


    }  catch {
        Write-Host "An error occurred:"
        Write-Host $_
    } finally {
    $doc.Save()
    if ($removePersonalInfo) {
        $doc.RemovePersonalInformation = $true
    }

    $doc.Close()
    $word.Quit()

    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
    [GC]::Collect()
    }
}

function get-powerpoint {
    param (
        [Parameter(Mandatory=$true)][string]$outputPath,
        [Parameter(Mandatory=$true)][Object[]]$macros,
        [Parameter(Mandatory=$false)][string[]]$GUIDs,
        [Parameter(Mandatory=$false)][bool]$removePersonalInfo
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
        if ($PSBoundParameters.ContainsKey('GUIDs')) {
            foreach ($GUID in $GUIDs) {
                $presentation.VBProject.References.AddFromGuid($GUID,$Major, $Minor) 
            }
        }
    
    }  catch {
        Write-Host "An error occurred:"
        Write-Host $_
    } finally {
        if ($removePersonalInfo) {
            $presentation.RemovePersonalInformation = $true
        }

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
        [Parameter(Mandatory=$true)][string]$outputPath,
        [Parameter(Mandatory=$true)][Object[]]$macros,
        [Parameter(Mandatory=$false)][string[]]$GUIDs,
        [Parameter(Mandatory=$false)][bool]$removePersonalInfo
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

        if ($PSBoundParameters.ContainsKey('GUIDs')) {
            foreach ($GUID in $GUIDs) {
                $acc.VBE.ActiveVBProject.References.AddFromGuid($GUID,$Major, $Minor)
            }
        }
    
    }   catch {
        Write-Host "An error occurred:"
        Write-Host $_
    } finally {
        if ($removePersonalInfo) {
            $acc.CurrentProject.RemovePersonalInformation = $true
        }

        $acc.CloseCurrentDatabase()
        $acc.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($acc)
        [GC]::Collect()
    }
}

function get-excel {
    param (
        [Parameter(Mandatory=$true)][string]$outputPath,
        [Parameter(Mandatory=$true)][Object[]]$macros,
        [Parameter(Mandatory=$false)][string[]]$GUIDs,
        [Parameter(Mandatory=$false)][bool]$removePersonalInfo
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

        if ($PSBoundParameters.ContainsKey('GUIDs')) {
            foreach ($GUID in $GUIDs) {
                $workbook.VBProject.References.AddFromGuid($GUID,$Major, $Minor)
            }
        }
        
    } catch {
        Write-Host "An error occurred:"
        Write-Host $_
    } finally {
        if ($removePersonalInfo) {
            $workbook.RemovePersonalInformation = $true
        }
        
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