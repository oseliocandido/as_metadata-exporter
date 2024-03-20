# Function to retrieve environment settings based on user choice
function Get-EnvironmentSettings {
    Write-Host "Select an environment:"
    Write-Host "1. Dev"
    Write-Host "2. Staging"
    Write-Host "3. Prod"

    $env_choice = Read-Host "Enter your choice (1, 2, 3)"

    # Settings based on user choice
    switch ($env_choice) {
        '1' { 
            @{
                serverNames = @("DevServer")
                prefixpathfile = "Dev"
                pattern = "*PATTERN1*"
            }
        }
        '2' { 
            @{
                serverNames = @("PT1550YW")
                prefixpathfile = "Staging"
                pattern = "*PATTERN2*"
            }
        }
        '3' { 
            @{
                serverNames = @("ProdServer")
                prefixpathfile = "Prod"
                pattern = "*PATTERN3*"
            }
        }
    }
}


function Main {
    [Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices")

    # Retrieve settings
    $selectedEnv = Get-EnvironmentSettings
    $serverNames = $selectedEnv.serverNames
    $prefixpathfile = $selectedEnv.prefixpathfile
    $pattern = $selectedEnv.pattern
    $targetBasePath = "C:\PATH_HERE\"

    # Create arrays to store file paths
    $filePaths_XMLA = @()
    $filePaths_DMV_Hierarchies = @()

    foreach ($serverName in $serverNames) {
        $server = New-Object Microsoft.AnalysisServices.Server
        $server.Connect($serverName)

        foreach ($database in $server.Databases) {
            $databaseName = $database.Name
            if ($databaseName -like $pattern) {
                #XMLA
                $stringWriter = New-Object System.IO.StringWriter
                $stringbuilder = New-Object System.Text.StringBuilder
                $stringwriter = New-Object System.IO.StringWriter($stringbuilder)
                $xmlOut = New-Object System.Xml.XmlTextWriter($stringwriter)
                $xmlOut.Formatting = [System.Xml.Formatting]::Indented
                $scriptObject = New-Object Microsoft.AnalysisServices.Scripter
                $MSASObject = [Microsoft.AnalysisServices.MajorObject[]] @($database)
                $ScriptObject.ScriptCreate($MSASObject, $xmlOut, $false)
                $filenameXMLA = Join-Path -Path $targetBasePath -ChildPath "$serverName.$databaseName.xmla"
                $stringbuilder.ToString() | Out-File -FilePath $filenameXMLA
                $filePaths_XMLA += $filenameXMLA

                #DMV with Hierachies Ordinal in Levels
                $filenameDMV_Hierarchies = Join-Path -Path $targetBasePath -ChildPath "$serverName.$databaseName.hierarchies.xml"
                $dmvQuery = "SELECT CUBE_NAME, LEVEL_NAME, [LEVEL_NUMBER],[HIERARCHY_UNIQUE_NAME] FROM [$databaseName].[$system].mdschema_levels WHERE LEVEL_NAME <> 'MeasuresLevel' AND LEVEL_NAME <> '(All)' AND LEVEL_ORIGIN = 1"
                $dmvResults = Invoke-ASCmd -Server $serverName -Database $databaseName -Query $dmvQuery 
                $dmvResults | Out-File -FilePath $filenameDMV_Hierarchies  -Encoding UTF8
                $filePaths_DMV_Hierarchies += $filenameDMV_Hierarchies
             
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Write-Host "$databaseName - ($timestamp)"
            }
        }
        $server.Disconnect()
    }

 
    $zipFilePathXMLA = Join-Path -Path $targetBasePath -ChildPath "$prefixpathfile.XMLA_Metadata.zip"
    $zipFilePathDMV_Hierarchies = Join-Path -Path $targetBasePath -ChildPath "$prefixpathfile.XML_Hierarchies_DMV_Metadata.zip"

 
    Write-Host "Zipping XMLA files..."
    Compress-Archive -Path $filePaths_XMLA -DestinationPath $zipFilePathXMLA

    Write-Host "Zipping DMV files..."
    Compress-Archive -Path $filePaths_DMV_Hierarchies -DestinationPath $zipFilePathDMV_Hierarchies

    Write-Host "Deleting files"
    Remove-Item $filePaths_XMLA
    Remove-Item $filePaths_DMV_Hierarchies
}

# Call the main function
Main
