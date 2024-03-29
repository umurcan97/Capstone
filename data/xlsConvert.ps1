param(
    $downloadfolder
)

#$downloadfolder="C:\Users\SnorlaX\Downloads\Yeni klas�r\Capstone\data"         # folder where the .xls files are

$uploadfolder   = "$downloadfolder/Upload"  # folder that uploads the .xlsx files
$backupfolder   = "$downloadfolder/Backup"  # folder that has .xls files as backup

if((Test-Path -Path "$uploadfolder") -eq $false)
{
    mkdir $uploadfolder
}

if((Test-Path -Path "$backupfolder") -eq $false)
{
    mkdir $backupfolder
}


# open and convert xls to xlsx
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false  # it is much faster if Excel is not visible

# loop through the .xls files and process them
Get-ChildItem -Path $downloadfolder -Filter '*.xls' -File | 
ForEach-Object {
    try {
        $xlsfilename = $_.FullName
        #copy file to backup folder
        Copy-Item -Path $xlsfilename -Destination $backupfolder -Force
        # open the xls
        Write-Host "Converting $xlsfilename"
        $workbook = $excel.Workbooks.Open($xlsfilename)
        # save converted file (as xlsx) directly to the upload folder
        $newfilename = Join-Path -Path $uploadfolder -ChildPath ('{0}.xlsx' -f $_.BaseName)
        $name2 = $xlsfilename.Substring(0,$xlsfilename.IndexOf('('))
        $name2 = $name2.Substring($name2.IndexOf("data\") + 5)
        if ($name2.Length -gt 31)
        {
            $name2 = $name2.Substring(0,31)
        }
        echo $name2
        $worksheet = $workbook.worksheets.item(1)
        $worksheet.name = "$name2"
        $workbook.SaveAs($newfilename, $xlFixedFormat)
		$workbook.Close()
        #remove old file
        Write-Host "Delete old file '$xlsfilename'"
        Remove-Item -Path $xlsfilename -Force
    }
    catch {
        # write out a warning as to why something went wrong
        Write-Warning "Could not convert '$xlsfilename':`r`n$($_.Exception.Message)"
    }
}
# close excel
$excel.Quit()
# garbage collection
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()