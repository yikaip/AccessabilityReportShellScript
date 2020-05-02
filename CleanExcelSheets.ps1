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
$newName = Read-Host "Is this sheet Violation or Needs Review"
if ($newName -Match "Violation" -or $newName -Match "Violations" -or $newName -Match "v"){
    $newName = "Violations"
    $tempName1 = $newName
}
elseif ($newName -Match "Needs Review" -or $newName -Match "Need Review" -or $newName -Match "n"){
    $newName = "Needs Review"
    $tempName2 = $newName
}
else {
    while($newName -NotMatch "Violation" -or $newName -NotMatch "Violations" -or $newName -NotMatch"Needs Review" -or $newName -NotMatch "Need Review"){
          $newName = Read-Host "That seems wrong, please check your spelling and only type in Violation or Needs Review"
          if ($newName -Match "Violation" -or $newName -Match "Violations" -or $newName -Match "v"){
              $newName = "Violations"
              $tempName1 = $newName
              break
          }
          elseif ($newName -Match "Needs Review" -or $newName -Match "Need Review" -or $newName -Match "n"){
              $newName = "Needs Review"
              $tempName2 = $newName
              break
          }
    }
}
Write-Host "---------------------------------------------------------------------------------------------------------"

#This function is to convert the excel files to .csv
#open the excel sheet
#then use the save as function to save the files as .cvs
function ConvertExcelToCSV {
    $file = "C:\Users\USERNAME\Downloads\" + $originName + ".xls"
    #start Excel
    $ExcelFile = New-Object -ComObject Excel.Application
    #open file
    $workBook = $ExcelFile.WorkBooks.Open($file)
    #make it unvisible
    $ExcelFile.Visible=$false
    #SaveAs allows us to change the file extension and file type, 
    #second parameter (number) is IMPORTANT since it specifies what type of file it is to save
    $workBook.SaveAs("C:\Users\USERNAME\Downloads\" + $saveAsName + ".csv", 6) 
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
    $P = Import-Csv -Path ("
    Downloads\" + $saveAsName + ".csv")-Delimiter "," | Where-Object  {
        $_.Note -CNotMatch "This HEAD does not contain a title element" -and $_.Note -CNotMatch "This LINK has an id attribute of 'font-awesome-5-kit-css', which is not unique" -and $_.Note -CNotMatch "This INPUT has an id attribute of" -and $_.Note -CNotMatch "carouselSS" -and $_.Note -CNotMatch "slick-slide" -and $_."Module Location" -CMatch "html"} | Export-Csv -Path ("C:\Users\USERNAME\Downloads\" + $newName +".csv")
}

#This function is to extract all violations that need a second-time review
function DoubleCheckViolations {
    $filterName = Read-Host "This is the file that stores all violations that needs us to double check. Give it a name"
    $P = Import-Csv -Path ("C:\Users\USERNAME\Downloads\" + $saveAsName + ".csv") | Where-Object  {$_.Note -Match "This P contains" -or $_.Note -Match "This H2" -and $_."Module Location" -CMatch "html"} | Export-Csv -Path ("C:\Users\USERNAME\Downloads\" + $filterName + ".csv")
    $S = Import-Csv -Path ("C:\Users\USERNAME\Downloads\" + $saveAsName + ".csv") | Where-Object {$_.Note -CNotMatch "This P contains" -and $_.Note -CNotMatch "This H2" -and $_."Module Location" -CMatch "html"} | Export-Csv -Path ("C:\Users\USERNAME\Downloads\" + $newName + ".csv")
}

#This function is to convert the .csv file back to excel so it's easier to send out
function ConvertCSVToExcel {
    $csvFile = "C:\Users\USERNAME\Downloads\" + $newName + ".csv"
    $Excel = New-Object -ComObject Excel.Application
    $wb = $Excel.WorkBooks.Open($csvFile)
    $worksheet = $wb.Sheets.Item(1)
    #delete the first row since it's contains info we dont need
    $worksheet.Cells.Item(1,1).EntireRow.Delete()
    #change the column width of the sheet so it fits with all characters 
    $ws = $wb.Sheets.Item($newName)
    $ws.columns.item(1).columnWidth = 40
    $ws.columns.item(2).columnWidth = 50
    $ws.columns.item(4).columnWidth = 70
    $ws.columns.item(5).columnWidth = 100
    $wb.SaveAs("C:\Users\USERNAME\Downloads\" + $newName + ".xlsx", 61)
    $wb.close()
    $Excel.Quit()
    Remove-Item -Path ("C:\Users\USERNAME\Downloads\" + $saveAsName + ".csv")
    Remove-Item -Path ("C:\Users\USERNAME\Downloads\" + $newName + ".csv")
}
function CombineTwoSheets {
    #import two excel files, 
    #assuming names are Violation and Needs Review,
    #then add Violation to Needs Review
    #Note: The file should be ready to send now. Before sending out change the filename in directory to 2020Q1-SITENAME

    $Excel = New-Object -ComObject Excel.Application
    $Excel.visible = $false
    $workbook1 = $Excel.WorkBooks.open("C:\Users\USERNAME\Downloads\" + $tempName1)
    $worksheet1 = $workbook1.WorkSheets.item("Violation")
    $range1 = $worksheet1.Range("A1:E1").EntireColumn
    $range1.Copy() | Out-Null
    
    $excel = New-Object -ComObject Excel.Application
    $excel.visible = $false
    $workbook2 = $excel.WorkBooks.open("C:\Users\USERNAME\Downloads\" + $tempName2)
    $worksheet2 = $workbook2.WorkSheets.Item("Needs Review")
    $worksheet2 = $workbook2.WorkSheets.Add()
    $worksheet2.Name = "Violations"
    $range2 = $worksheet2.Range("A1:E1").EntireColumn
    $worksheet2.Paste($range2)
    $worksheet2.columns.item(1).columnWidth = 40
    $worksheet2.columns.item(2).columnWidth = 50
    $worksheet2.columns.item(4).columnWidth = 70
    $worksheet2.columns.item(5).columnWidth = 100
    
    $workbook1.close()
    $workbook2.close()
    $Excel.Quit()
    $excel.Quit()
    
    
}

#--------------MAIN-------------#
#First convert excel sheets to csv files;
#if it is a violation sheet, run delete function, else fun double check function;
#convert csv files back to excel files;
#if the there are two excel files, combine them together, else end this program.

ConvertExcelToCSV
#Have the user tells us what it should be. Note: Type in violation for "Violation Sheet" and anything else for "Needs Review Sheet"
$Ans = Read-Host "Violation/Needs Review"
iif ($Ans -Match "Violation" -or $Ans -Match "Violations" -or $Ans -Match "v"){
    Delete
}
elseif ($Ans -Match "Needs Review" -or $Ans -Match "Need Review" -or $Ans -Match "n"){
    DoubleCheckViolations
    Write-Host "Note: After checking those items, don't forget to copy and paste violations that are valid back to" $newName -ForegroundColor Red -BackgroundColor White
}
else {
    while($Ans -NotMatch "Violation" -or $Ans -NotMatch "Violations" -or $Ans -NotMatch"Needs Review" -or $Ans -NotMatch "Need Review"){
        $Ans = Read-Host "That seems wrong, please check your spelling and only type in Violation or Needs Review"
        if ($Ans -Match "Violation" -or $Ans -Match "Violations"){
            Delete
            break
        }
        elseif ($Ans -Match "Needs Review" -or $Ans -Match "Need Review"){
            DoubleCheckViolations
            Write-Host "Note: After checking those items, don't forget to copy and paste violations that are valid back to" $newName -ForegroundColor Red -BackgroundColor White
            break
        }
    }
}
ConvertCSVToExcel
#Have the user tells us if there are two files yet.
$Run = Read-Host "Do you have both of the two sheets, Violation and Needs Review? If there is only one sheet, end here."
if ($Run -Match "Yes" -or $Run -Match "Y"){
    CombineTwoSheets
}
else{
    exit
}
Write-Host "---------------------------------------------------------------------------------------------------------"
