# Analysis Services Metadata Exporter

PowerShell script to automate the extraction and archiving of metadata from Analysis Services databases. It generates XMLA scripts and queries to gather detailed hierarchys metadata information from dimensions.

## Features

- Allows selection of environment (Dev, Staging, Prod) to target specific Analysis Services instances;
- Extracts all metadata from SSAS instances;
- Archives metadata into ZIP files for easy storage and distribution.

## Prerequisites

- PowerShell 5.1 or higher;
- Microsoft.AnalysisServices Assembly

## Usage

1. Clone or download the script file [`ASMetadataExtractor.ps1`](ASMetadataExtractor.ps1) to your local machine.

2. Open PowerShell and navigate to the directory where its located and run the script with

```
.\ASMetadataExtractor.ps1
```

3. The script will connect to the specified Analysis Services instance(s), extract metadata and generate ZIP files containing the exported metadata.

4. Once completed, individual XMLA and DMV files will be deleted, leaving only the ZIP archives.

## Customization

- You can customize the script by modifying the predefined patterns and server names for each environment in the `Get-EnvironmentSettings` function.
- Adjust the target base path (`$targetBasePath`) to specify the directory where the exported metadata ZIP files will be stored.
