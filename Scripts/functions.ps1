function get-word {
<#
.SYNOPSIS
    Create a Word Doc file from a 1 to n .bas or .cls files.

.DESCRIPTION
    This function create a Microsoft Word Doc file from a 1 to n .bas or .cls files.
    It supports the ability to add libraries if provided a string array of GUIDs.
    And it supports the ability to remove personal info from the filenames as well.


.PARAMETER outputPath
    The output path where the created output file will be saved. 

.PARAMETER macros
    The input path containing the .bas and .cls files

.PARAMETER GUIDs
    Optional - An array of GUIDs to be installed in the output file

.PARAMETER removePersonalInfo
    Optional - A parameter that will remove personal information from the output file.

.EXAMPLE
    $outputPath = "directory:\for\output\path"
    $macros = "directory:\for\macros"
    get-word -outputPath $outputPath -macros $macros

.EXAMPLE
    $outputPath = "directory:\for\output\path"
    $macros = "directory:\for\macros"
    get-word -outputPath $outputPath -macros $macros removePersonalInfo

.EXAMPLE
    $outputPath = "directory:\for\output\path"
    $macros = "directory:\for\macros"
    $scriptingGuid = "{420B2830-E718-11CF-893D-00A0C9054228}"
    $GUIDs = @()
    $GUIDs += $scriptingGuid
    get-word -outputPath $outputPath -macros $macros -GUIDs $GUIDs -removePersonalInfo

.INPUTS
    [String]outputPath
    [String]macros
    [Object[]]GUIDs
    [switch]removePersonalInfo

.OUTPUTS
    [void]

.NOTES
    Author: Brian Gonzalez
    Email: b.gonzalez.programming@gmail.com
#>
    param (
        [Parameter(Mandatory=$true)][string]$outputPath,
        [Parameter(Mandatory=$true)][Object[]]$macros,
        [Parameter(Mandatory=$false)][string[]]$GUIDs,
        [Parameter(Mandatory=$false)][switch]$removePersonalInfo
    )

    try {
        $word = New-Object -ComObject word.application
        $doc = $word.documents.add()
        $wdFormatFlatXMLMacroEnabled = 13
        $doc.SaveAs($outputPath,$wdFormatFlatXMLMacroEnabled)
    
        $Major = 0
        $Minor = 0
    
        foreach ($macro in $macros) {
            if ($macro.Extension -ne ".doccls") {
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
        [Parameter(Mandatory=$false)][switch]$removePersonalInfo
    )

    try {
        $powerpoint = New-Object -ComObject powerpoint.application
        $presentation = $powerpoint.Presentations.Add()
        $ppSaveAsOpenXMLPresentationMacroEnabled = 25
        $presentation.SaveAs($outputPath,$ppSaveAsOpenXMLPresentationMacroEnabled)
    
        $Major = 0
        $Minor = 0
    
        foreach ($macro in $macros) {
            if ($macro.Extension -ne ".doccls") {
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
        [Parameter(Mandatory=$false)][switch]$removePersonalInfo
    )

    try {
        $acc = New-Object -ComObject Access.Application
        $acFileFormatAccess2007 = 12
        $acc.NewCurrentDataBase($outputPath,$acFileFormatAccess2007)
    
        $acCmdCompileAndSaveAllModules = 126
        $acModule = 5
    
        $Major = 0
        $Minor = 0
    
        foreach ($macro in $macros) {
            if ($macro.Extension -ne ".doccls") {
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
        [Parameter(Mandatory=$false)][switch]$removePersonalInfo
    )
    try {
        $excel = New-Object -ComObject excel.application
        $workbook = $excel.Workbooks.Add()
        $xlOpenXMLWorkbookMacroEnabled = 52
        $workbook.SaveAs($outputPath, $xlOpenXMLWorkbookMacroEnabled)
    
        $Major = 0
        $Minor = 0
    
        foreach ($macro in $macros) {
            if ($macro.Extension -ne ".doccls") {
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