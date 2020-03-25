<# --I am using VS Code and this is just a rough draft
 I saved all files as csv cuz it's easier, will improve it later
 also replace the import path to the actual file location, 
 replace the export path to the folder we want it to be --Yikai#>

 <#Reference: 
 1. https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/import-csv?view=powershell-7
 2. https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/where-object?view=powershell-7
 Microsoft docs are cooooooool#>

#Replace this path with actual file location, will improve it later
#Where-Object is to find all violations that satisfy the following conditions
#CNotMatch contains condictions aka all the violaions that can be ignored, 
#there should be a more effective way to do it i guess? 
#but right now I am just copying everything from the Violation Not Sent file in Google Drive



$P = Import-Csv -Path "C:\Downloads\288883.csv" | Where-Object  {$_.Note -CNotMatch "This HEAD does not contain a title element" -and 
$_.Note -CNotMatch "This LINK has an id attribute of 'font-awesome-5-kit-css', which is not unique" -and 
$_.Note -CNotMatch "This INPUT has an id attribute of" -and
$_.Note -CNotMatch "carouselSS"} | Export-Csv -Path "C:\Downloads\New1.csv"


#literally, it's to format the excel table, not sure if it's actully nessecary 
#just throwing it out there for now
$P | Format-Table

