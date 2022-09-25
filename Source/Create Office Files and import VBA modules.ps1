# NOTE: For this script to work, trust must be given to the VBA object model for this to work.
# This will be in the options -> trust center settings -> Macro Settings  for the Office 
# applications. If this is not enabled, the code will fail.

# Access does not have the ability to give trust this way. But this is a moot point since
# Access imports VBA modules a different way.

# NOTE: Get-ExcelGuid prompts Excel to save. I can't figure out how to prevent this.
# Application.DisplayAlerts does not appear to disable it.

function Get-ExcelGuid {
    try {

        $excel = New-Object -ComObject excel.application
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
    $parentDir = (get-item $curDir).parent
    Set-Location $parentDir
    $fluentName = "Fluent VBA 1.65"
    $outputPath = "$parentDir\Distribution\Fluent VBA"
    $guidObj = Get-ExcelGuid
    $GuidStr = Out-String -NoNewline -InputObject $guidObj
    $xlGuid = $GuidStr.Replace("System.__ComObject","")
    $ScriptingGuid = "{420B2830-E718-11CF-893D-00A0C9054228}"

    $distFiles = $macros = Get-ChildItem -Path .\Distribution -File

    foreach ($file in $distfiles) {
        $file.delete()
    }

    $excel = New-Object -ComObject excel.application
    $word = New-Object -ComObject word.application
    $powerpoint = New-Object -ComObject powerpoint.application
    $acc = New-Object -ComObject Access.Application

    $workbook = $excel.Workbooks.Add()
    $doc = $word.documents.add()
    $presentation = $powerpoint.Presentations.Add()

    $xlOpenXMLWorkbookMacroEnabled = 52
    $wdFormatFlatXMLMacroEnabled = 13
    $ppSaveAsOpenXMLPresentationMacroEnabled = 25
    $acFileFormatAccess2007 = 12

    $workbook.SaveAs($outputPath, $xlOpenXMLWorkbookMacroEnabled)
    $doc.SaveAs($outputPath,$wdFormatFlatXMLMacroEnabled)
    $presentation.SaveAs($outputPath,$ppSaveAsOpenXMLPresentationMacroEnabled)
    $acc.NewCurrentDataBase($outputPath,$acFileFormatAccess2007)

    $acCmdCompileAndSaveAllModules = 126
    $acModule = 5
    $macros = Get-ChildItem -Path .\Source -File

    $Major = 0
    $Minor = 0

    foreach ($macro in $macros) {
        if ($macro.Extension -ne ".doccls" -and $macro.Extension -ne ".ps1" -and $macro.BaseName -ne "mTODO") {
            
            $workbook.VBProject.VBComponents.Import($macro.FullName) | Out-Null
            $doc.VBProject.VBComponents.Import($macro.FullName)
            $presentation.VBProject.VBComponents.Import($macro.FullName)

            $acc.VBE.ActiveVBProject.VBComponents.Import($macro.FullName)
            $acc.VBE.ActiveVBProject.VBComponents($acc.VBE.ActiveVBProject.VBComponents.Count).Name = $macro.BaseName
            $acc.DoCmd.RunCommand($acCmdCompileAndSaveAllModules)
            $acc.DoCmd.Save($acModule, $macro.BaseName)
        }
    }

    $doc.VBProject.References.AddFromGuid($xlGuid,$Major, $Minor)
    $presentation.VBProject.References.AddFromGuid($xlGuid,$Major, $Minor)
    $acc.VBE.ActiveVBProject.References.AddFromGuid($xlGuid,$Major, $Minor)

    $workbook.VBProject.References.AddFromGuid($ScriptingGuid,$Major, $Minor)
    $doc.VBProject.References.AddFromGuid($ScriptingGuid,$Major, $Minor)
    $presentation.VBProject.References.AddFromGuid($ScriptingGuid,$Major, $Minor)
    $acc.VBE.ActiveVBProject.References.AddFromGuid($ScriptingGuid,$Major, $Minor)

}

Catch {
    Write-Host "An error occurred:"
    Write-Host $_
}

Finally {

    $workbook.Save()
    $doc.Save()
    $presentation.Save()
    

    $workbook.Close()
    $doc.Close()
    $presentation.Close()
    #$acc.CloseCurrentDatabase()
    
    $excel.Quit()
    $word.Quit()
    $powerpoint.Quit()
    $acc.Quit()

    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($acc)
    [GC]::Collect()

    foreach ($file in $distfiles) {
        Rename-Item $file.FullName -NewName "$($fluentName)$($file.Extension)"
        #Write-Output "$($fluentName)$($file.Extension)"
    }
}