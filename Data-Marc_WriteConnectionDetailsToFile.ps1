<# 
    External Tools Power BI model documentor script version 1.1.0
    New in this version: 
        - Support for installation via PowerBI.tips Business Ops. 
        - Automatic detection of installer location vs default location. 
        - Automatic download of the pbit file if it cannot be found.  

    Full change log can be found here: https://data-marc.com/model-documenter/
#>

# Below you can define your personal preference for file saving and reading. 
# The default location can be changed and will be leverages througout the entire script. 
# InstallerLocation only applies to installation via PowerBI.tips Business Ops. 
$InstallerLocation = '__TOOL_INSTALL_DIR__'
$defaultLocation = 'c:\BusinessOpsTemp\'
$finalLocation = if($InstallerLocation -like '*TOOL_INSTALL_DIR*') 
{$defaultLocation} else {$InstallerLocation}

#This part starts tracing to catch unfortunate errors and defines where to write the file. 
$Logfile = $finalLocation + 'PBI_DocumentModel_LogFile.txt'
Start-Transcript -Path $Logfile

# Function to automatically download the pbit file if it cannot be found on the defined location. 
# Function based on https://gist.github.com/chrisbrownie/f20cb4508975fb7fb5da145d3d38024a 
function DownloadFilesFromRepo {
Param(
    $Owner = 'marclelijveld',
    $Repository = 'External-Tools-Model-Documentation',
    $Path = 'ModelDocumentationTemplate.pbit',
    $DestinationPath = 'C:\BusinessOpsTemp'
    )

    $baseUri = "https://api.github.com/"
    $args = "repos/$Owner/$Repository/contents/$Path"
    $wr = Invoke-WebRequest -Uri $($baseuri+$args)
    $objects = $wr.Content | ConvertFrom-Json
    $files = $objects | where {$_.type -eq "file"} | Select -exp download_url
    $directories = $objects | where {$_.type -eq "dir"}
    
    $directories | ForEach-Object { 
        DownloadFilesFromRepo -Owner $Owner -Repository $Repository -Path $_.path -DestinationPath $($DestinationPath+$_.name)
    }

    
    if (-not (Test-Path $DestinationPath)) {
        # Destination path does not exist, let's create it
        try {
            New-Item -Path $DestinationPath -ItemType Directory -ErrorAction Stop
        } catch {
            throw "Could not create path '$DestinationPath'!"
        }
    }

    foreach ($file in $files) {
        $fileDestination = Join-Path $DestinationPath (Split-Path $file -Leaf)
        try {
            Invoke-WebRequest -Uri $file -OutFile $fileDestination -ErrorAction Stop -Verbose
            "Grabbed '$($file)' to '$fileDestination'"
        } catch {
            throw "Unable to download '$($file.path)'"
        }
    }

}

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
try {
    Invoke-Item $PBITLocation  -ErrorAction Stop 
    }
catch {
         DownloadFilesFromRepo
         Invoke-Item $PBITLocation
      }

# Stop tracing errors
Stop-Transcript