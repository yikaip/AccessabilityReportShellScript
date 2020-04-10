<#
Author: Yikai Peng

Please replace file path with the warehouse drive path

Reference: 
 1. https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/import-csv?view=powershell-7
 2. https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/where-object?view=powershell-7
#>

#Start with asking for the file name. We need the filename when it's just got downloaded, 
#when it turns to .csv and the actual name we want.
Write-Host "---------------------------------------------------------------------------------------------------------"
$originName = Read-Host "Please type in th filename, ex.288831"
$saveAsName = Read-Host "We need to save this file as .csv just for now. Please type in the temporary name for this csv file"
$newName = Read-Host "Okay, One more step. Please type in the actual name of this file, ex.2020Q1-ANISCI."
Write-Host "---------------------------------------------------------------------------------------------------------"

#This function is to convert the excel files to .csv
#open the excel sheet
#then use the save as function to save the files as .cvs
function ConvertExcelToCSV {
    $file = "C:\Downloads\" + $originName + ".xlsx"
    #start Excel
    $ExcelFile = New-Object -ComObject Excel.Application
    #open file
    $workBook = $ExcelFile.WorkBooks.Open($file)
    #make it unvisible
    $ExcelFile.Visible=$false
    #SaveAs allows us to change the file extension and file type, 
    #second parameter (number) is IMPORTANT since it specifies what type of file it is to save
    $workBook.SaveAs("C:\Downloads\" + $saveAsName + ".csv", 6) 
    #Quit and close excel
    $workBook.close()
    $ExcelFile.Quit()
    
}

#Use .csv files to do the following work

#Where-Object is to find all violations that satisfy the following conditions like Regular Expression
#CNotMatch contains conditions aka all the violaions that can be ignored (From Google Drive "Violation Not Send")
#Likewise, CMatch does the same thing, except that it only extract all violations that we need to double check before they get sent out

#This function is to delete all unnecessary template level errors
function Delete () {
    $P = Import-Csv -Path ("C:\Downloads\" + $saveAsName + ".csv")-Delimiter "," | Where-Object  {
        $_.Note -CNotMatch "This HEAD does not contain a title element" -and $_.Note -CNotMatch "This LINK has an id attribute of 'font-awesome-5-kit-css', which is not unique" -and $_.Note -CNotMatch "This INPUT has an id attribute of" -and $_.Note -CNotMatch "carouselSS" -and $_.Note -CNotMatch "slick-slide" -and $_."Module Location" -CMatch "html"} | Export-Csv -Path ("C:\Downloads\" + $newName +".csv")
}

#This function is to extract all violations that need a second-time review
function DoubleCheckViolations {
    $filterName = Read-Host "This is the file that stores all violations that needs us to double check. Give it a name"
    $P = Import-Csv -Path ("C:\Downloads\" + $saveAsName + ".csv") | Where-Object  {$_.Note -Match "This P contains" -or $_.Note -Match "This H2" -and $_."Module Location" -CMatch "html"} | Export-Csv -Path ("C:\Downloads\" + $filterName + ".csv")
    $S = Import-Csv -Path ("C:\Downloads\" + $saveAsName + ".csv") | Where-Object {$_.Note -CNotMatch "This P contains" -and $_.Note -CNotMatch "This H2" -and $_."Module Location" -CMatch "html"} | Export-Csv -Path ("C:\Downloads\" + $newName + ".csv")
}

#This function is to convert the .csv file back to excel so it's easier to send out
function ConvertCSVToExcel {
    $csvFile = "C:\Downloads\" + $newName + ".csv"
    $Excel = New-Object -ComObject Excel.Application
    $wb = $Excel.WorkBooks.Open($csvFile)
    $worksheet = $wb.Sheets.Item(1)
    #delete the first row since it's contains info we dont need
    $worksheet.Cells.Item(1,1).EntireRow.Delete()
    $wb.SaveAs("C:\Downloads\" + $newName + ".xlsx", 61)
    $wb.close()
    $Excel.Quit()
}

#--------------MAIN-------------#
ConvertExcelToCSV
#Have the user tells us what it should be. Note: Type in violation or v for "Violation Sheet" and anything else for "Needs Review Sheet"
$Ans = Read-Host "Violation/Needs Review"
if ($Ans -Match "Violation"){
    Delete
} 
else{
    DoubleCheckViolations
    Write-Host "Note: After checking those items, don't forget to copy and paste violations that are valid back to" $newName -ForegroundColor Red -BackgroundColor White
}
ConvertCSVToExcel
Write-Host "---------------------------------------------------------------------------------------------------------"
