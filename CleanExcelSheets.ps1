<# --I am using VS Code and this is just a rough draft
 I saved all files as csv cuz it's easier, will improve it later
 also need to replace the import path to the actual file location, 
 replace the export path to the folder we want it to be --Yikai#>

 <#Reference: 
 1. https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/import-csv?view=powershell-7
 2. https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/where-object?view=powershell-7
 Microsoft docs are cooooooool#>

#Replace this path with actual file location: ex. warehouse/2020Q1..., will improve it later
#First thing is to open the excel sheet
#then use the save as function to save the files as .cvs

$file = "C:\Downloads\288883.xls"

#start Excel
$ExcelFile = New-Object -ComObject Excel.Application
#open file
$workBook = $ExcelFile.WorkBooks.Open($file)
#make it unvisible
$ExcelFile.Visible=$false
$workBook.SaveAs("C:\Downloads\288883.csv", 6)  #second number specifies what tyoe of file it is to save!!!

#This is where I got stuck. I am trying to highlight all text with "This P..." 
#but it looks like it's not going in to the loop
#pretty sure I did somthing wrong, any ideas?

#Get cells info, this excel sheet is called Violations
$col = 4 #Error message is in column D "Note" 
$row = 200 #initialize row
$worksheet = $workBook.Worksheets.item("Violations") #This lets me access the whole "Violations" sheet?
foreach ($worksheet in $workBook){
    if ($worksheet -match "This P contains text with a background color of #ffffff rgb(255, 255, 255) and foreground color of #ffc425 rgb(255, 196, 37) that is less than 18 point in size; or bold text less than 14 point in size that has a luminosity contrast ratio of 1.59, which is below 4.5:1"){
        $Range = $worksheet.Range($row,$col)
        $row++
        $Range.Interior.Color = RGB(0,0,255)
        $Range.Font.Bold = $true
    }
}

#close
$workBook.close()

#Use .csv files to do the following work

#Where-Object is to find all violations that satisfy the following conditions
#CNotMatch contains condictions aka all the violaions that can be ignored, 
#there should be a more effective way to do it i guess? 
#but right now I am just copying everything from the Violation Not Sent file in Google Drive

$P = Import-Csv -Path "C:\Downloads\288883.csv" | Where-Object  {$_.Note -CNotMatch "This HEAD does not contain a title element" -and 
$_.Note -CNotMatch "This LINK has an id attribute of 'font-awesome-5-kit-css', which is not unique" -and 
$_.Note -CNotMatch "This INPUT has an id attribute of" -and
$_.Note -CNotMatch "carouselSS"} | Export-Csv -Path "C:\Downloads\New1.csv"


#not sure if it's actully nessecary just throwing it out there for now
$P | Format-Table

