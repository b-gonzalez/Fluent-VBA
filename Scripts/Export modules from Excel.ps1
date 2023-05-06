#Need to update file to deal with userforms

#Convert this to a function that can accept a directory as a parameter

try {
    $curDir = $PSScriptRoot
    Set-Location $curDir
    $devFiles = Get-ChildItem -Path .\Development
    $exp = Resolve-Path .\export
    $mostRecentFile = $devFiles| Sort-Object LastWriteTime | Select-Object -last 1
    $excel = New-Object -ComObject excel.application
    $excel.Visible = $true
    $workbook = $excel.Workbooks.Open($mostRecentFile)
    $GUID = "{0002E157-0000-0000-C000-000000000046}"
    $Major = 0
    $Minor = 0

    $vbe = $excel.application.VBE
    $vbProj = $vbe.ActiveVBProject
    $vbComps = $vbProj.VBComponents
    $extension = ""

    foreach($comp in $vbComps) {
        if ($comp.Type -eq 1) {
            $extension = ".bas"
        } elseif ($comp.Type -eq 2) {
            $extension = ".cls"}
        #  elseif ($comp.Type -eq 3) {
        #     $extension = ".frm"}
         elseif ($comp.Type -eq 100) {
            $extension = ".doccls"
        } else {
            Write-Output "Comp name is $($comp.name) and ext is $($comp.Type)"
        }
        if ($comp.Type -eq 1 -or $comp.Type -eq 2) {
            $comp.export("$exp\$($comp.Name)$($extension)")
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