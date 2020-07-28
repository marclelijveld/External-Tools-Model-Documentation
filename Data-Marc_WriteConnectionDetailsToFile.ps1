# Below you can define your personal preference for file saving and reading. 
# The outputfolder will be used to dump a temporarily file with the connection to the model. 
# The PBITLocation defines where the templated PBIT file is saved. 
$OutputFolder = 'c:\temp\'
$PBITLocation = 'C:\temp\ModelDocumentationTemplate.pbit'

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
$OutputLocation = $OutputFolder + 'ModelDocumenterConnectionDetails.json'
$json  | ConvertTo-Json  | Out-File $OutputLocation

# Open PBIT template file from PBITLocation as defined in the variable. 
Invoke-Item $PBITLocation