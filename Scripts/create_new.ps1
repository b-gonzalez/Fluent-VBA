try {
    $curDir = $PSScriptRoot
    $parentDir = (get-item $curDir).parent
    Set-Location $parentDir
    $outputPath = "$parentDir\Distribution\Fluent VBA"
    $guidObj = Get-ExcelGuid
    $GuidStr = Out-String -NoNewline -InputObject $guidObj
    $xlGuid = $GuidStr.Replace("System.__ComObject","")
    $ScriptingGuid = "{420B2830-E718-11CF-893D-00A0C9054228}"

    $distFiles = Get-ChildItem -Path .\Distribution -File

    foreach ($file in $distfiles) {
        $file.delete()
    }
}

function make-excel {
    param (
        [string]$outputPath,
        [string]$macroPath,
        [string]$ScriptingGuid
        [string]$minorVersion,
        [string]$majorVersion
    )
    $excel = New-Object -ComObject excel.application
    $workbook = $excel.Workbooks.Add()
    $xlOpenXMLWorkbookMacroEnabled = 52
    $workbook.SaveAs($outputPath, $xlOpenXMLWorkbookMacroEnabled)
    $macros = Get-ChildItem -Path .\Source -File

    $Major = 0
    $Minor = 0

    foreach ($macro in $macros) {
        if ($macro.Extension -ne ".doccls" -and $macro.Extension -ne ".ps1" -and $macro.BaseName -ne "mTODO") {
            $workbook.VBProject.VBComponents.Import($macro.FullName) | Out-Null
        }
    }

    $workbook.VBProject.References.AddFromGuid($ScriptingGuid,$Major, $Minor)
}

function make-word {
    param (
        [string]$outputPath,
        [string]$ExcelGuid,
        [string]$ScriptingGuid
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

function make-powerpoint {
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

function make-access {
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