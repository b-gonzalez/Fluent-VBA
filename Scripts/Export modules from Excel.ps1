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