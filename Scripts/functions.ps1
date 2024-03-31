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
        NOTE: This function requires "Trust access to the VBA project object model" to be enabled on Word
    #>
    [OutputType([void])]
    param (
        [Parameter(Mandatory = $true)][string]$outputPath,
        [Parameter(Mandatory = $true)][Object[]]$macros,
        [Parameter(Mandatory = $false)][string[]]$GUIDs,
        [Parameter(Mandatory = $false)][switch]$removePersonalInfo
    )

    try {
        $distPath = "$outputPath\Distribution\Fluent VBA"
        $testPath = "$outputPath\test_files\test_fluent"

        $word = New-Object -ComObject word.application
        $doc = $word.documents.add()
        $wdFormatFlatXMLMacroEnabled = 13
        $doc.SaveAs($distPath, $wdFormatFlatXMLMacroEnabled)
    
        $Major = 0
        $Minor = 0

        $doc.VBProject.name = "fluent_vba"

        #This section requires "Trust access to the VBA project object model" to be enabled.
        #If it is not enabled this section will fail.
        foreach ($macro in $macros) {
            if ($macro.Extension -ne ".doccls") {
                $doc.VBProject.VBComponents.Import($macro.FullName)
            }
        }
        
        if ($PSBoundParameters.ContainsKey('GUIDs')) {
            foreach ($GUID in $GUIDs) {
                $doc.VBProject.References.AddFromGuid($GUID, $Major, $Minor)
            }
        }

        $mInit = $macros | Where-Object name -Contains "mInit.bas"

        $doc2 = $word.documents.add()
        $doc2.SaveAs($testPath, $wdFormatFlatXMLMacroEnabled)
        $doc2.VBProject.name = "fluent_vba_test"
        $doc2.VBProject.VBComponents.Import($mInit.FullName) | Out-Null


    }
    catch {
        Write-Host "An error occurred:"
        Write-Host $_
    }
    finally {
        $doc.Save()
        $doc2.Save()

        if ($removePersonalInfo) {
            $doc.RemovePersonalInformation = $true
            $doc2.RemovePersonalInformation = $true
        }

        $doc.Close()
        $doc2.Close()
        $word.Quit()

        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
        [GC]::Collect()
    }
}

function get-powerpoint {
    <#
    .SYNOPSIS
        Create a PowerPoint Presentation file from a 1 to n .bas or .cls files.

    .DESCRIPTION
        This function create a Microsoft PowerPoint Presentation file from a 1 to n .bas or .cls files.
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
        get-powerpoint -outputPath $outputPath -macros $macros

    .EXAMPLE
        $outputPath = "directory:\for\output\path"
        $macros = "directory:\for\macros"
        get-powerpoint -outputPath $outputPath -macros $macros removePersonalInfo

    .EXAMPLE
        $outputPath = "directory:\for\output\path"
        $macros = "directory:\for\macros"
        $scriptingGuid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        $GUIDs = @()
        $GUIDs += $scriptingGuid
        get-powerpoint -outputPath $outputPath -macros $macros -GUIDs $GUIDs -removePersonalInfo

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
        NOTE: This function requires "Trust access to the VBA project object model" to be enabled on PowerPoint
    #>
    [OutputType([void])]
    param (
        [Parameter(Mandatory = $true)][string]$outputPath,
        [Parameter(Mandatory = $true)][Object[]]$macros,
        [Parameter(Mandatory = $false)][string[]]$GUIDs,
        [Parameter(Mandatory = $false)][switch]$removePersonalInfo
    )

    try {
        $distPath = "$outputPath\Distribution\Fluent VBA"
        $testPath = "$outputPath\test_files\test_fluent"

        $powerpoint = New-Object -ComObject powerpoint.application
        $presentation = $powerpoint.Presentations.Add()
        $ppSaveAsOpenXMLPresentationMacroEnabled = 25
        $presentation.SaveAs($distPath, $ppSaveAsOpenXMLPresentationMacroEnabled)
    
        $Major = 0
        $Minor = 0

        $presentation.VBProject.name = "fluent_vba"
    
        #This section requires "Trust access to the VBA project object model" to be enabled.
        #If it is not enabled this section will fail.
        foreach ($macro in $macros) {
            if ($macro.Extension -ne ".doccls") {
                $presentation.VBProject.VBComponents.Import($macro.FullName)
            }
        }
        if ($PSBoundParameters.ContainsKey('GUIDs')) {
            foreach ($GUID in $GUIDs) {
                $presentation.VBProject.References.AddFromGuid($GUID, $Major, $Minor) 
            }
        }

        $mInit = $macros | Where-Object name -Contains "mInit.bas"

        $pres2 = $powerpoint.Presentations.Add()
        $pres2.SaveAs($testPath, $ppSaveAsOpenXMLPresentationMacroEnabled)
        $pres2.VBProject.name = "fluent_vba_test"
        $pres2.VBProject.VBComponents.Import($mInit.FullName) | Out-Null
    
    }
    catch {
        Write-Host "An error occurred:"
        Write-Host $_
    }
    finally {
        if ($removePersonalInfo) {
            $presentation.RemovePersonalInformation = $true
            $pres2.RemovePersonalInformation = $true
        }

        $presentation.Save()
        $pres2.Save()
        $presentation.Close()
        $pres2.Close()
        $powerpoint.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation)
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint)
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres2)
        [GC]::Collect()
    }
}

function get-access {
    <#
    .SYNOPSIS
        Create an Access database file from a 1 to n .bas or .cls files.

    .DESCRIPTION
        This function create a Microsoft Access database file from a 1 to n .bas or .cls files.
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
        get-access -outputPath $outputPath -macros $macros

    .EXAMPLE
        $outputPath = "directory:\for\output\path"
        $macros = "directory:\for\macros"
        get-access -outputPath $outputPath -macros $macros removePersonalInfo

    .EXAMPLE
        $outputPath = "directory:\for\output\path"
        $macros = "directory:\for\macros"
        $scriptingGuid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        $GUIDs = @()
        $GUIDs += $scriptingGuid
        get-access -outputPath $outputPath -macros $macros -GUIDs $GUIDs -removePersonalInfo

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
    [OutputType([void])]
    param (
        [Parameter(Mandatory = $true)][string]$outputPath,
        [Parameter(Mandatory = $true)][Object[]]$macros,
        [Parameter(Mandatory = $false)][string[]]$GUIDs,
        [Parameter(Mandatory = $false)][switch]$removePersonalInfo
    )

    try {
        $distPath = "$outputPath\Distribution\Fluent VBA"
        $testPath = "$outputPath\test_files\test_fluent"

        $acc = New-Object -ComObject Access.Application
        $acFileFormatAccess2007 = 12
        $acc.NewCurrentDataBase($distPath, $acFileFormatAccess2007)
    
        $acCmdCompileAndSaveAllModules = 126
        $acModule = 5
    
        $Major = 0
        $Minor = 0

        $acc.VBE.ActiveVBProject.name = "fluent_vba"
    
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
                $acc.VBE.ActiveVBProject.References.AddFromGuid($GUID, $Major, $Minor)
            }
        }

        $acc.CloseCurrentDatabase()

        $mInit = $macros | Where-Object name -Contains "mInit.bas"

        $acc.NewCurrentDataBase($testPath, $acFileFormatAccess2007)
        $acc.VBE.ActiveVBProject.name = "fluent_vba_test"
        $acc.VBE.ActiveVBProject.VBComponents.Import($mInit.FullName)
        # $acc.VBE.ActiveVBProject.VBComponents($acc.VBE.ActiveVBProject.VBComponents.Count).Name = $mInit.BaseName
        $acc.DoCmd.RunCommand($acCmdCompileAndSaveAllModules)
        $acc.DoCmd.Save($acModule, $mInit.BaseName)
    
    }
    catch {
        Write-Host "An error occurred:"
        Write-Host $_
    }
    finally {
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
    <#
    .SYNOPSIS
        Create an Excel workbook file from a 1 to n .bas or .cls files.

    .DESCRIPTION
        This function create a Microsoft Excel workbook file from a 1 to n .bas or .cls files.
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
        get-excel -outputPath $outputPath -macros $macros

    .EXAMPLE
        $outputPath = "directory:\for\output\path"
        $macros = "directory:\for\macros"
        get-excel -outputPath $outputPath -macros $macros removePersonalInfo

    .EXAMPLE
        $outputPath = "directory:\for\output\path"
        $macros = "directory:\for\macros"
        $scriptingGuid = "{420B2830-E718-11CF-893D-00A0C9054228}"
        $GUIDs = @()
        $GUIDs += $scriptingGuid
        get-excel -outputPath $outputPath -macros $macros -GUIDs $GUIDs -removePersonalInfo

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
        NOTE: This function requires "Trust access to the VBA project object model" to be enabled on Excel
    #>
    [OutputType([void])]
    param (
        [Parameter(Mandatory = $true)][string]$outputPath,
        [Parameter(Mandatory = $true)][Object[]]$macros,
        [Parameter(Mandatory = $false)][string[]]$GUIDs,
        [Parameter(Mandatory = $false)][switch]$removePersonalInfo
    )
    try {
        $distPath = "$outputPath\Distribution\Fluent VBA"
        $testPath = "$outputPath\test_files\test_fluent"

        $excel = New-Object -ComObject excel.application
        $workbook = $excel.Workbooks.Add()
        $xlOpenXMLWorkbookMacroEnabled = 52
        $workbook.SaveAs($distPath, $xlOpenXMLWorkbookMacroEnabled)
    
        $Major = 0
        $Minor = 0

        $workbook.VBProject.name = "fluent_vba"
    
        #This section requires "Trust access to the VBA project object model" to be enabled.
        #If it is not enabled this section will fail.
        foreach ($macro in $macros) {
            if ($macro.Extension -ne ".doccls" -and $macro.Name -notlike "mInit.bas") {
                $workbook.VBProject.VBComponents.Import($macro.FullName) | Out-Null
            }
        }

        if ($PSBoundParameters.ContainsKey('GUIDs')) {
            foreach ($GUID in $GUIDs) {
                $workbook.VBProject.References.AddFromGuid($GUID, $Major, $Minor)
            }
        }

        $mInit = $macros | Where-Object name -Contains "mInit.bas"

        $wb2 = $excel.Workbooks.Add()
        $wb2.SaveAs($testPath, $xlOpenXMLWorkbookMacroEnabled)
        $wb2.VBProject.name = "fluent_vba_test"
        $wb2.VBProject.VBComponents.Import($mInit.FullName) | Out-Null
        
    }
    catch {
        Write-Host "An error occurred:"
        Write-Host $_
    }
    finally {
        if ($removePersonalInfo) {
            $workbook.RemovePersonalInformation = $true
            $wb2.RemovePersonalInformation = $true
        }
        
        $workbook.Save()
        $workbook.Close()
        $wb2.Save()
        $wb2.Close()
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
        $extensibilityGUID = "{0002E157-0000-0000-C000-000000000046}"
        $Major = 0
        $Minor = 0

        $workbook.VBProject.References.AddFromGuid($extensibilityGUID, $Major, $Minor)
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