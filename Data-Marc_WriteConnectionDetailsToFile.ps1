<# 
    External Tools Power BI model documentor script version 1.0.1
    New in this version: 
        - Easier editing of personal preferred location. This is now defined in one single variable ($defaultlocation). 
        - Added logging for the script. 
        - Prepared for upcoming changes with $finalLocation. Soon more about this.
        - Better error handling in the pbit file which results in less errors while loading the data. 

    Full change log can be found here: https://data-marc.com/model-documenter/
#>

# Below you can define your personal preference for file saving and reading. 
# The default location can be changed and will be leverages througout the entire script. 
$defaultLocation = 'c:\temp\'
$finalLocation = $defaultLocation

#This part starts tracing to catch unfortunate errors and defines where to write the file. 
$Logfile = $finalLocation + 'PBI_DocumentModel_LogFile.txt'
Start-Transcript -Path $Logfile

# The PBITLocation defines where the templated PBIT file is saved. 
# In case you are using a different pbit file, you can define that in below variable.
$PBITLocation = $finalLocation + 'ModelDocumentationTemplate.pbit'

# Below section defines the server and databasename based on the input captured from the External tools integration. 
# This is defined as arguments \"%server%\" and \"%database%\" in the external tools json. 
$Server = $args[0]
$DatabaseName = $args[1]

# Write Server and Database information to screen. 
Write-Host $Server 
Write-Host $DatabaseName 

# Generate json array based on the received server and database information.
$json = @"
    {
    "Server": "$Server", 
    "DatabaseName": "$DatabaseName"
    }
"@

# Writes the output in json format to the defined file location. This is a temp location and will be overwritten next time. 
$OutputLocation = $finalLocation + 'ModelDocumenterConnectionDetails.json'
$json  | ConvertTo-Json  | Out-File $OutputLocation

# Open PBIT template file from PBITLocation as defined in the variable. 
Invoke-Item $PBITLocation

# Catching possible errors and writing to the log file
Write-Output 'Errors occured: ' $error.count
Write-Output $error[0]

Stop-Transcript