function Get-LastModifiedExcel {
    param (
        [Parameter(Mandatory=$true)][string]$ExcelDir
    )

    [string]$mostRecentFile = Get-ChildItem -Path $ExcelDir | Sort-Object LastWriteTime | Select-Object -last 1
    return $mostRecentFile
}

enum VbaType {
    bas = 1
    cls = 2
    userform = 3
    doccls = 100
}

$dups_arr = @(".htm, .html", ".pdf", ".txt", ".xml", ".xps", ".rtf")
$excel_arr = @(".csv", ".dbf", ".dif", ".mht, .mhtml", ".ods", ".prn", ".slk", ".xla", ".xlam", ".xls", ".xlsb", ".xlsm", ".xlsx", ".xlt", ".xltm", ".xltx", ".xlw")
$powerpoint_arr = @(".bmp", ".emf", ".gif", ".jpg", ".mp4", ".odp", ".png", ".pot", ".potm", ".potx", ".ppa", ".ppam", ".pps", ".ppsm", ".ppsx", ".ppt", ".pptm", ".pptx", ".thmx", ".tif", ".wmf", ".wmv")
$word_arr = @(".doc", ".docm", ".docx", ".dot", ".dotm", ".dotx", ".mht; .mhtml", ".odt", ".wps")
$access_arr = @(".adn", ".accdb", ".accdr", ".accdt", ".accda", ".mdw", ".accde", ".mam", ".maq", ".mar", ".mat", ".maf", ".laccdb", ".ade", ".adp", ".mdb", ".cdb", ".mda", ".mdn", ".mdf", ".mde", ".ldb")

function Export-ExcelModules {
    param (
        [Parameter(Mandatory=$true)][string]$ExcelFilePath,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [Parameter(Mandatory=$false)][switch]$DeleteOutputContents,
        [Parameter(Mandatory=$false)][switch]$ExcelVisible,
        [Parameter(Mandatory=$false)][VbaType[]]$TypesArr

    )

    [string[]]$fileArr = $ExcelFilePath.Split(".")
    $fileExt = $fileArr[$fileArr.GetUpperBound(0)]

    if ($fileExt -eq "xlsm" -or $fileExt -eq "xlsb") {
        try {

            $wbCopy = "$($fileArr[0]) - Copy.$fileExt"

            Copy-Item $wbPath -Destination $wbCopy

            $excel = New-Object -ComObject excel.application
            
            $workbook = $excel.Workbooks.Open($wbCopy)

            if ($ExcelVisible) {
                $excel.Visible = $true
            }
    
            if ($TypesArr.Length -eq 0) {
                $TypesArr += [VbaType]::bas
                $TypesArr += [VbaType]::cls
                $TypesArr += [VbaType]::doccls
                $TypesArr += [VbaType]::userform
            }
    
            if ($deleteOutputContents -and (Test-Path -Path $outputPath)) {
                Remove-Item -Path "$outputPath\*" -Recurse
            }
        
            $vbe = $excel.application.VBE
            $vbProj = $vbe.ActiveVBProject
            $vbComps = $vbProj.VBComponents
            $extension = ""
        
            foreach($comp in $vbComps) {
                if ($comp.Type -eq [VbaType]::bas) {
                    $extension = ".bas"
                } elseif ($comp.Type -eq [VbaType]::cls) {
                    $extension = ".cls"
                } elseif ($comp.Type -eq [VbaType]::userform) {
                    $extension = ".frm"
                } elseif ($comp.Type -eq [VbaType]::doccls) {
                    $extension = ".doccls"
                } else {
                    Write-Output "Comp name is $($comp.name) and ext is $($comp.Type)"
                }
    
                if ($TypesArr -contains $comp.Type) {
                    $comp.export("$OutputPath\$($comp.Name)$($extension)")
                }
            }
        }
        catch {
            Write-Host "An error occurred:"
            Write-Host $_
        }
        finally {
            $workbook.Close()
            $excel.Quit()
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
            [GC]::Collect()
            Remove-Item $wbCopy
        }
    }
}

function Export-WordModules {
    param (
        [Parameter(Mandatory=$true)][string]$OfficeFilePath,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [Parameter(Mandatory=$false)][switch]$DeleteOutputContents,
        [Parameter(Mandatory=$false)][switch]$AppVisible,
        [Parameter(Mandatory=$false)][VbaType[]]$TypesArr

    )

    [string[]]$fileArr = $OfficeFilePath.Split(".")
    $fileExt = $fileArr[$fileArr.GetUpperBound(0)]
    
    try {

        $docCopy = "$($fileArr[0]) - Copy.$fileExt"

        Copy-Item $OfficeFilePath -Destination $docCopy

        $word = New-Object -ComObject word.application
        
        $document = $word.documents.Open($docCopy)

        if ($AppVisible) {
            $word.Visible = $true
        }

        if ($TypesArr.Length -eq 0) {
            $TypesArr += [VbaType]::bas
            $TypesArr += [VbaType]::cls
            $TypesArr += [VbaType]::doccls
            $TypesArr += [VbaType]::userform
        }

        if ($deleteOutputContents -and (Test-Path -Path $outputPath)) {
            Remove-Item -Path "$outputPath\*" -Recurse
        }
    
        $vbe = $word.application.VBE
         #Needed to account for the normal.docm project in Word files
        if ($vbe.VBProjects.Count -gt 1) {
            $vbProj = $vbe.VBProjects(1)
        } else {
            $vbProj = $vbe.ActiveVBProject
        }
        $vbComps = $vbProj.VBComponents
        $vbProj = $vbe.ActiveVBProject
        $vbComps = $vbProj.VBComponents

        $extension = ""
    
        foreach($comp in $vbComps) {
            if ($comp.Type -eq [VbaType]::bas) {
                $extension = ".bas"
            } elseif ($comp.Type -eq [VbaType]::cls) {
                $extension = ".cls"
            } elseif ($comp.Type -eq [VbaType]::userform) {
                $extension = ".frm"
            } elseif ($comp.Type -eq [VbaType]::doccls) {
                $extension = ".doccls"
            } else {
                Write-Output "Comp name is $($comp.name) and ext is $($comp.Type)"
            }

            if ($TypesArr -contains $comp.Type) {
                $comp.export("$OutputPath\$($comp.Name)$($extension)")
            }
        }
    }
    catch {
        Write-Host "An error occurred:"
        Write-Host $_
    }
    finally {
        $document.Close()
        $word.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
        [GC]::Collect()
        Remove-Item $docCopy
    }
}

function Export-PowerPointModules {
    param (
        [Parameter(Mandatory=$true)][string]$OfficeFilePath,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [Parameter(Mandatory=$false)][switch]$DeleteOutputContents,
        [Parameter(Mandatory=$false)][switch]$AppVisible,
        [Parameter(Mandatory=$false)][VbaType[]]$TypesArr

    )

    [string[]]$fileArr = $OfficeFilePath.Split(".")
    $fileExt = $fileArr[$fileArr.GetUpperBound(0)]
    
    try {

        $presCopy = "$($fileArr[0]) - Copy.$fileExt"

        Copy-Item $OfficeFilePath -Destination $presCopy

        $powerpoint = New-Object -ComObject powerpoint.application
        
        $pres = $powerpoint.Presentations.Open($presCopy)

        if ($AppVisible) {
            $powerpoint.Visible = $true
        }

        if ($TypesArr.Length -eq 0) {
            $TypesArr += [VbaType]::bas
            $TypesArr += [VbaType]::cls
            $TypesArr += [VbaType]::doccls
            $TypesArr += [VbaType]::userform
        }

        if ($deleteOutputContents -and (Test-Path -Path $outputPath)) {
            Remove-Item -Path "$outputPath\*" -Recurse
        }
    
        $vbe = $powerpoint.VBE
        $vbProj = $vbe.ActiveVBProject
        $vbComps = $vbProj.VBComponents

        $extension = ""
    
        foreach($comp in $vbComps) {
            if ($comp.Type -eq [VbaType]::bas) {
                $extension = ".bas"
            } elseif ($comp.Type -eq [VbaType]::cls) {
                $extension = ".cls"
            } elseif ($comp.Type -eq [VbaType]::userform) {
                $extension = ".frm"
            } elseif ($comp.Type -eq [VbaType]::doccls) {
                $extension = ".doccls"
            } else {
                Write-Output "Comp name is $($comp.name) and ext is $($comp.Type)"
            }

            if ($TypesArr -contains $comp.Type) {
                $comp.export("$OutputPath\$($comp.Name)$($extension)")
            }
        }
    }
    catch {
        Write-Host "An error occurred:"
        Write-Host $_
    }
    finally {
        $pres.Close()
        $powerpoint.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint)
        [GC]::Collect()
        Remove-Item $presCopy
    }
}

function Export-AccessModules {
    param (
        [Parameter(Mandatory=$true)][string]$OfficeFilePath,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [Parameter(Mandatory=$false)][switch]$DeleteOutputContents,
        [Parameter(Mandatory=$false)][switch]$AppVisible,
        [Parameter(Mandatory=$false)][VbaType[]]$TypesArr
    )

    [string[]]$fileArr = $OfficeFilePath.Split(".")
    $fileExt = $fileArr[$fileArr.GetUpperBound(0)]
    
    try {

        $dbCopy = "$($fileArr[0]) - Copy.$fileExt"

        Copy-Item $OfficeFilePath -Destination $dbCopy

        $access = New-Object -ComObject access.application

        $db = $access.Application.OpenCurrentDatabase($dbCopy)

        if ($AppVisible) {
            $access.Visible = $true
        }

        if ($TypesArr.Length -eq 0) {
            $TypesArr += [VbaType]::bas
            $TypesArr += [VbaType]::cls
            $TypesArr += [VbaType]::doccls
            $TypesArr += [VbaType]::userform
        }

        if ($deleteOutputContents -and (Test-Path -Path $outputPath)) {
            Remove-Item -Path "$outputPath\*" -Recurse
        }
    
        $vbe = $access.VBE
        $vbProj = $vbe.ActiveVBProject
        $vbComps = $vbProj.VBComponents

        $extension = ""
    
        foreach($comp in $vbComps) {
            if ($comp.Type -eq [VbaType]::bas) {
                $extension = ".bas"
            } elseif ($comp.Type -eq [VbaType]::cls) {
                $extension = ".cls"
            } elseif ($comp.Type -eq [VbaType]::userform) {
                $extension = ".frm"
            } elseif ($comp.Type -eq [VbaType]::doccls) {
                $extension = ".doccls"
            } else {
                Write-Output "Comp name is $($comp.name) and ext is $($comp.Type)"
            }

            if ($TypesArr -contains $comp.Type) {
                $comp.export("$OutputPath\$($comp.Name)$($extension)")
            }
        }
    }
    catch {
        Write-Host "An error occurred:"
        Write-Host $_
    }
    finally {
        $access.CloseCurrentDatabase()
        $access.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
        [GC]::Collect()
        Remove-Item $dbCopy
    }
}