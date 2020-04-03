<# Author: Yikai#>

 <#
 Reference: 
 1. https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/import-csv?view=powershell-7
 2. https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/where-object?view=powershell-7
 #>



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
    $P = Import-Csv -Path "C:\Downloads\new.csv" -Delimiter "," | Where-Object  {
        $_.Note -CNotMatch "This HEAD does not contain a title element" -and $_.Note -CNotMatch "This LINK has an id attribute of 'font-awesome-5-kit-css', which is not unique" -and $_.Note -CNotMatch "This INPUT has an id attribute of" -and $_.Note -CNotMatch "carouselSS"} | Export-Csv -Path "C:\Downloads\New1.csv" 
                           #replace this path with actual file location  #delect all items that has "" 
}

function DoubleCheckViolations {
    $P = Import-Csv -Path "C:\Downloads\new.csv" -Delimiter ","| Where-Object  {
        $_.Note -Match "This P contains" -or $_.Note -Match "This H2"} | Export-Csv -Path "C:\Users\yikai\Downloads\New1.csv"
}


#--------------MAIN-------------#
ConvertExcelToCSV
$Ans = Read-Host "Violation/Needs Review"
if ($Ans = "Violation"){
    Delete
} elseif ($Ans = "Needs Review") {
    DoubleCheckViolations
}


