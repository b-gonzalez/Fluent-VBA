#Excel memory bug issue fixed. Do testing on other office apps to make sure that they're working as expected and with no memory leaks.

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

function Get-LastModifiedFile {
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