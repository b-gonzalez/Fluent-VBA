#testing 123

enum VbaType {
    bas = 1
    cls = 2
    userform = 3
    doccls = 100
}

enum OfficeApplication {
    Excel = 1
    Word = 2
    PowerPoint = 3
    Access = 4
}

function Get-LastModifiedExcel {
    param (
        [Parameter(Mandatory = $true)][string]$ExcelDir
    )

    [string]$mostRecentFile = Get-ChildItem -Path $ExcelDir | Sort-Object LastWriteTime | Select-Object -last 1
    return $mostRecentFile
}

function get-OfficeApplication {
    param (
        [Parameter(Mandatory = $true)][OfficeApplication]$OfficeApp
    )

    if ($OfficeApp -eq [OfficeApplication]::Excel) {
        $app = New-Object -ComObject excel.application
    }
    elseif ($OfficeApp -eq [OfficeApplication]::Word) {
        $app = New-Object -ComObject word.application
    }
    elseif ($OfficeApp -eq [OfficeApplication]::PowerPoint) {
        $app = New-Object -ComObject powerpoint.application
    }
    elseif ($OfficeApp -eq [OfficeApplication]::Access) {
        $app = New-Object -ComObject access.application
    }

    return $app

    Write-Output "Hello world!"
}

function Get-FileExnteionValid {
    param (
        [Parameter(Mandatory = $true)][OfficeApplication]$OfficeApp,
        [Parameter(Mandatory = $true)][string]$Extension
    )

    [bool]$validFileExtension = $false

    $dups_arr = @(".htm, .html", ".pdf", ".txt", ".xml", ".xps", ".rtf")

    $excel_arr = @(".csv", ".dbf", ".dif", ".mht, .mhtml", ".ods", ".prn", ".slk", ".xla", ".xlam", ".xls", ".xlsb", ".xlsm", ".xlsx", ".xlt", ".xltm", ".xltx", ".xlw")
    $excel_arr += $dups_arr

    $word_arr = @(".doc", ".docm", ".docx", ".dot", ".dotm", ".dotx", ".mht; .mhtml", ".odt", ".wps")
    $word_arr += $dups_arr

    $powerpoint_arr = @(".bmp", ".emf", ".gif", ".jpg", ".mp4", ".odp", ".png", ".pot", ".potm", ".potx", ".ppa", ".ppam", ".pps", ".ppsm", ".ppsx", ".ppt", ".pptm", ".pptx", ".thmx", ".tif", ".wmf", ".wmv")
    $powerpoint_arr += $dups_arr

    $access_arr = @(".adn", ".accdb", ".accdr", ".accdt", ".accda", ".mdw", ".accde", ".mam", ".maq", ".mar", ".mat", ".maf", ".laccdb", ".ade", ".adp", ".mdb", ".cdb", ".mda", ".mdn", ".mdf", ".mde", ".ldb")

    if ($OfficeApp -eq [OfficeApplication]::Excel) {
        $validFileExtension = $excel_arr.Contains(".$Extension")
    }
    elseif ($OfficeApp -eq [OfficeApplication]::Word) {
        $validFileExtension = $word_arr.Contains(".$Extension")
    }
    elseif ($OfficeApp -eq [OfficeApplication]::PowerPoint) {
        $validFileExtension = $powerpoint_arr.Contains(".$Extension")
    }
    elseif ($OfficeApp -eq [OfficeApplication]::Access) {
        $validFileExtension = $access_arr.Contains(".$Extension")
    }

    return $validFileExtension
}

function get-OfficeFile {
    param (
        [Parameter(Mandatory = $true)][OfficeApplication]$OfficeApp,
        [Parameter(Mandatory = $true)]$app,
        [Parameter(Mandatory = $true)][string]$fileCopyPath
    )

    if ($OfficeApp -eq [OfficeApplication]::Excel) {
        $file = $app.Workbooks.Open($fileCopyPath)
    }
    elseif ($OfficeApp -eq [OfficeApplication]::Word) {
        $file = $app.documents.Open($fileCopyPath)
    }
    elseif ($OfficeApp -eq [OfficeApplication]::PowerPoint) {
        $file = $app.Presentations.Open($fileCopyPath)
    }
    elseif ($OfficeApp -eq [OfficeApplication]::Access) {
        $file = $app.Application.OpenCurrentDatabase($fileCopyPath)
    }

    return $file
}

function Export-ExcelModules {
    param (
        [Parameter(Mandatory = $true)][string]$ExcelFilePath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $false)][switch]$DeleteOutputContents,
        [Parameter(Mandatory = $false)][switch]$ExcelVisible,
        [Parameter(Mandatory = $false)][VbaType[]]$TypesArr

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
        
            foreach ($comp in $vbComps) {
                if ($comp.Type -eq [VbaType]::bas) {
                    $extension = ".bas"
                }
                elseif ($comp.Type -eq [VbaType]::cls) {
                    $extension = ".cls"
                }
                elseif ($comp.Type -eq [VbaType]::userform) {
                    $extension = ".frm"
                }
                elseif ($comp.Type -eq [VbaType]::doccls) {
                    $extension = ".doccls"
                }
                else {
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
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($vbe)
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($vbProj)
            [GC]::Collect()
            Remove-Variable excel
            Remove-Item $wbCopy
        }
    }
}

function Export-WordModules {
    param (
        [Parameter(Mandatory = $true)][string]$OfficeFilePath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $false)][switch]$DeleteOutputContents,
        [Parameter(Mandatory = $false)][switch]$AppVisible,
        [Parameter(Mandatory = $false)][VbaType[]]$TypesArr

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
        }
        else {
            $vbProj = $vbe.ActiveVBProject
        }
        $vbComps = $vbProj.VBComponents
        $vbProj = $vbe.ActiveVBProject
        $vbComps = $vbProj.VBComponents

        $extension = ""
    
        foreach ($comp in $vbComps) {
            if ($comp.Type -eq [VbaType]::bas) {
                $extension = ".bas"
            }
            elseif ($comp.Type -eq [VbaType]::cls) {
                $extension = ".cls"
            }
            elseif ($comp.Type -eq [VbaType]::userform) {
                $extension = ".frm"
            }
            elseif ($comp.Type -eq [VbaType]::doccls) {
                $extension = ".doccls"
            }
            else {
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
        [Parameter(Mandatory = $true)][string]$OfficeFilePath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $false)][switch]$DeleteOutputContents,
        [Parameter(Mandatory = $false)][switch]$AppVisible,
        [Parameter(Mandatory = $false)][VbaType[]]$TypesArr

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
    
        foreach ($comp in $vbComps) {
            if ($comp.Type -eq [VbaType]::bas) {
                $extension = ".bas"
            }
            elseif ($comp.Type -eq [VbaType]::cls) {
                $extension = ".cls"
            }
            elseif ($comp.Type -eq [VbaType]::userform) {
                $extension = ".frm"
            }
            elseif ($comp.Type -eq [VbaType]::doccls) {
                $extension = ".doccls"
            }
            else {
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
        [Parameter(Mandatory = $true)][string]$OfficeFilePath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $false)][switch]$DeleteOutputContents,
        [Parameter(Mandatory = $false)][switch]$AppVisible,
        [Parameter(Mandatory = $false)][VbaType[]]$TypesArr
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
    
        foreach ($comp in $vbComps) {
            if ($comp.Type -eq [VbaType]::bas) {
                $extension = ".bas"
            }
            elseif ($comp.Type -eq [VbaType]::cls) {
                $extension = ".cls"
            }
            elseif ($comp.Type -eq [VbaType]::userform) {
                $extension = ".frm"
            }
            elseif ($comp.Type -eq [VbaType]::doccls) {
                $extension = ".doccls"
            }
            else {
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

function Export-Modules {
    param (
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][OfficeApplication]$OfficeApp,
        [Parameter(Mandatory = $false)][switch]$DeleteOutputContents,
        [Parameter(Mandatory = $false)][switch]$AppVisible,
        [Parameter(Mandatory = $false)][VbaType[]]$TypesArr
    )

    [string[]]$fileArr = $filePath.Split(".")
    $fileExt = $fileArr[$fileArr.GetUpperBound(0)]

    if (Get-FileExnteionValid -OfficeApp $OfficeApp -Extension $fileExt) {
        try {
            & {
                $fileCopy = "$($fileArr[0]) - Copy.$fileExt"

                Copy-Item $FilePath -Destination $fileCopy

                $app = get-OfficeApplication -OfficeApp $OfficeApp

                $file = get-OfficeFile -OfficeApp $OfficeApp -app $app -fileCopyPath $fileCopy

                if ($appVisible) {
                    $app.Visible = $true
                }

                if ($TypesArr.Length -eq 0) {
                    $TypesArr = @()
                    $TypesArr += [VbaType]::bas
                    $TypesArr += [VbaType]::cls
                    $TypesArr += [VbaType]::doccls
                    $TypesArr += [VbaType]::userform
                }

                if ($deleteOutputContents -and (Test-Path -Path $outputPath)) {
                    Remove-Item -Path "$outputPath\*" -Recurse
                }

                if ($OfficeApp -eq [OfficeApplication]::PowerPoint) {
                    $vbe = $app.VBE
                }
                else {
                    $vbe = $app.application.VBE
                }
                #Needed to account for the normal.docm project in Word files
                if ($OfficeApp -eq [OfficeApplication]::Word) {
                    $vbProj = $vbe.VBProjects(1)
                }
                else {
                    $vbProj = $vbe.ActiveVBProject
                }

                $vbComps = $vbProj.VBComponents
                $extension = ""
            
                foreach ($comp in $vbComps) {
                    if ($comp.Type -eq [VbaType]::bas) {
                        $extension = ".bas"
                    }
                    elseif ($comp.Type -eq [VbaType]::cls) {
                        $extension = ".cls"
                    }
                    elseif ($comp.Type -eq [VbaType]::userform) {
                        $extension = ".frm"
                    }
                    elseif ($comp.Type -eq [VbaType]::doccls) {
                        $extension = ".doccls"
                    }
                    else {
                        Write-Output "Comp name is $($comp.name) and ext is $($comp.Type)"
                    }

                    if ($TypesArr -contains $comp.Type) {
                        $comp.export("$OutputPath\$($comp.Name)$($extension)")
                    }
                }

                if ($OfficeApp -eq [OfficeApplication]::Access) {
                    $app.CloseCurrentDatabase()
                    $app.Quit()
                }
                else {
                    $file.Close()
                    $app.Quit()
                }

                Remove-Item $fileCopy
            }
        }
        catch {
            Write-Host "An error occurred:"
            Write-Host $_
        }
        finally {
            [GC]::Collect()
        }
    }
}