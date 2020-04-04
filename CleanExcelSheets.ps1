<# Author: Yikai#>

 <#
 Reference: 
 1. https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/import-csv?view=powershell-7
 2. https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/where-object?view=powershell-7
 #>


 Write-Host "---------------------------------------------------------------------------------------------------------"
 $originName = Read-Host "Please type in th file name, ex.288831"
 $saveAsName = Read-Host "We need to save this file as .csv just for now. Please type in your temporary for this csv file"
 $newName = Read-Host "Okay, One more step. Please type in the actual name of this file, ex.2020Q1-ANISCI."
 Write-Host "---------------------------------------------------------------------------------------------------------"
function ConvertExcelToCSV {
    #Replace this path with actual file location: ex. warehouse/2020Q1..., will improve it later
    #First thing is to open the excel sheet
    #then use the save as function to save the files as .cvs
    $file = "C:\Downloads\2888831.xlsx"
    #start Excel
    $ExcelFile = New-Object -ComObject Excel.Application
    #open file
    $workBook = $ExcelFile.WorkBooks.Open($file)
    #make it visible
    $ExcelFile.Visible=$false
    $workBook.SaveAs("C:\Downloads\new.csv", 6) #second number specifies what tyoe of file it is to save!!!
    $workBook.close()
    $ExcelFile.Quit()
    
}

#Use .csv files to do the following work

#Where-Object is to find all violations that satisfy the following conditions
#CNotMatch contains condictions aka all the violaions that can be ignored, 
#there should be a more effective way to do it i guess? 
#but right now I am just copying everything from the Violation Not Sent file in Google Drive


function Delete () {
    $P = Import-Csv -Path ("C:\Downloads\" + $saveAsName + ".csv")-Delimiter "," | Where-Object  {
        $_.Note -CNotMatch "This HEAD does not contain a title element" -and $_.Note -CNotMatch "This LINK has an id attribute of 'font-awesome-5-kit-css', which is not unique" -and $_.Note -CNotMatch "This INPUT has an id attribute of" -and $_.Note -CNotMatch "carouselSS"} | Export-Csv -Path ("C:\Downloads\" + $newName +".csv")
                           #replace this path with actual file location  #delect all items that has "" 
}

function DoubleCheckViolations {
    $P = Import-Csv -Path ("C:\Downloads\" + $saveAsName + ".csv") -Delimiter ","| Where-Object  {
        $_.Note -Match "This P contains" -or $_.Note -Match "This H2"} | Export-Csv -Path ("C:\Downloads\" + $newName +".csv")
}

function ConvertCSVToExcel {
    $csvFile = "C:\Downloads\" + $newName + ".csv"
    $Excel = New-Object -ComObject Excel.Application
    $wb = $Excel.WorkBooks.Open($csvFile)
    $worksheet = $wb.Sheets.Item(1)
    $worksheet.Cells.Item(1,1).EntireRow.Delete()
    $wb.SaveAs("C:\Downloads\" + $newName + ".xlsx", 61)
    $wb.close()
    $Excel.Quit()
}

#--------------MAIN-------------#
ConvertExcelToCSV
$Ans = Read-Host "Violation/Needs Review"
if ($Ans = "Violation"){
    Delete
} elseif ($Ans = "Needs Review") {
    DoubleCheckViolations
}
ConvertCSVToExcel
Write-Host "---------------------------------------------------------------------------------------------------------"
