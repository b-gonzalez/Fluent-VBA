# NOTE: For this script to work, trust must be given to the VBA object model for this to work.
# This will be in the options -> trust center settings -> Macro Settings  for the Office 
# applications. If this is not enabled, the code will fail.

# Access does not have the ability to give trust this way. But this is a moot point since
# Access imports VBA modules a different way.

try {
    $curDir = $PSScriptRoot
    $parentDir = (get-item $curDir).parent
    Set-Location $parentDir
    $outputPath = "$parentDir\Distribution\Fluent VBA"

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
    
    $acModule = 5
    $macros = Get-ChildItem -Path .\Source -File

    foreach ($macro in $macros) {
        if ($macro.Extension -ne ".doccls" -and $macro.Extension -ne ".ps1" -and $macro.BaseName -ne "mTODO") {
            
            $workbook.VBProject.VBComponents.Import($macro.FullName) | Out-Null
            $doc.VBProject.VBComponents.Import($macro.FullName)
            $presentation.VBProject.VBComponents.Import($macro.FullName)
            $acc.Application.LoadFromText($acModule, $macro.BaseName,$macro)
        }
    }

    $workbook.Save()
    $doc.Save()
    $presentation.Save()

}

Catch {
    Write-Host "An error occurred:"
    Write-Host $_
}

Finally {
    $workbook.Close()
    $doc.Close()
    $presentation.Close()
    $acc.CloseCurrentDatabase()
    
    $excel.Quit()
    $word.Quit()
    $powerpoint.Quit()
    $acc.Quit()

    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($acc)
    [GC]::Collect()
}
