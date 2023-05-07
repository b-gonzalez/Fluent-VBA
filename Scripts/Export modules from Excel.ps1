#Need to update file to deal with userforms

#This code requires a reference to the VB extensibility library. Consider adding a check to see if that library is available and if not add it dynamically.

#Write code to call this script from VBA. This should be possible by opening a copy of the workbook as read-only and then executing the script.

#Add functionality to Export-ExcelModules to optionally delete all files in the folder

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
    doccls = 100
}

function Export-ExcelModules {
    param (
        [Parameter(Mandatory=$true)][string]$ExcelFilePath,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [Parameter(Mandatory=$false)][VbaType[]]$TypesArr

    )
    try {
        $excel = New-Object -ComObject excel.application
        $excel.Visible = $true
        $workbook = $excel.Workbooks.Open($ExcelFilePath)
        $GUID = "{0002E157-0000-0000-C000-000000000046}"
        $Major = 0
        $Minor = 0

        if ($TypesArr.Length -eq 0) {
            $TypesArr += [VbaType]::bas
            $TypesArr += [VbaType]::cls
            $TypesArr += [VbaType]::doccls
        }
    
        $vbe = $excel.application.VBE
        $vbProj = $vbe.ActiveVBProject
        $vbComps = $vbProj.VBComponents
        $extension = ""
    
        foreach($comp in $vbComps) {
            if ($comp.Type -eq [VbaType]::bas) {
                $extension = ".bas"
            } elseif ($comp.Type -eq [VbaType]::cls) {
                $extension = ".cls"}
            #  elseif ($comp.Type -eq 3) {
            #     $extension = ".frm"}
             elseif ($comp.Type -eq [VbaType]::doccls) {
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
        $workbook.Save()
        $workbook.Close()
        $excel.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        [GC]::Collect()
    }
}