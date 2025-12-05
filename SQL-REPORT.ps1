#////////////////////////////////////////////
# SQL-REPORT.ps1 by Tim Fischbach
#////////////////////////////////////////////
# Check if Import-Excel module is installed, if not install it for current user
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "[INFO] ImportExcel module not found. Installing for current user..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
    Write-Host "[INFO] ImportExcel module installed successfully." -ForegroundColor Green
}

$scriptname = "SQL-REPORT.ps1"
$author = "Tim Fischbach"
$ver = "v1.0.3"
$builddate = "16.10.2025"

<# Changelog:
1.0.3: Added .txt cleanup functionality
1.0.2: Made Excel table filterable
1.0.1: Excel gets deleted before recreation
1.0.0: Init Version
#>

Clear-Host
$host.UI.RawUI.ForegroundColor = "Cyan"
Write-Host "///////////////////////////////////////////////////////"
Write-Host "///// Script: $scriptname"
Write-Host "///// Author: $author"
Write-Host "///// Version: $ver"
Write-Host "///// Build Date: $builddate"
Write-Host "///////////////////////////////////////////////////////"
Write-Host ""
$host.UI.RawUI.ForegroundColor = "White"
Start-Sleep 2
Write-Host "[INFO] Starting SQL Server inventory collection..."

# Update folder path to include txt subfolder
$folderPath = Join-Path (Get-Location).Path "txt"

Write-Host "[INFO] Scanning folder: $folderPath"

# Initialize an empty array to store the data
$data = @()

# Define the MSSQL version mapping based on the provided reference table with version numbers
$versionMapping = @(
    @{ Version = 16.0; Name = "SQL Server 2022" }
    @{ Version = 15.0; Name = "SQL Server 2019" }
    @{ Version = 14.0; Name = "SQL Server 2017" }
    @{ Version = 13.0; Name = "SQL Server 2016" }
    @{ Version = 12.0; Name = "SQL Server 2014" }
    @{ Version = 11.0; Name = "SQL Server 2012" }
    @{ Version = 10.50; Name = "SQL Server 2008 R2" }
    @{ Version = 10.0; Name = "SQL Server 2008" }
)

Write-Host "[INFO] Processing text files..."
# Iterate over each file in the folder
Get-ChildItem -Path $folderPath -Filter *.txt | ForEach-Object {
    Write-Host "[INFO] Processing file: $($_.Name)"
    
    $filename = $_.Name
    $hostname = $filename.Split('_')[0]
    
    # Variables to store system info
    $cpuDetails = ""
    $cpuCores = 0
    $ram = "0 GB"
    
    # Read the content of the file
    $content = Get-Content -Path $_.FullName
    
    # Process the first line for system info
    if ($content[0] -match 'System Info: CPU: (.*?)\s*,\s*Cores: (\d+),\s*RAM: (\d+) GB') {
        $cpuDetails = $matches[1].Trim()
        $cpuCores = $matches[2]
        $ram = "$($matches[3]) GB"
    }
    
    # Skip the header line and separator, start processing from line 3
    $content | Select-Object -Skip 2 | ForEach-Object {
        if ($_ -match 'Instance: (.*), Edition: (.*), Version: (.*)') {
            $instance = $matches[1]
            $edition = $matches[2]
            $version = $matches[3]
            
            # Convert version string to number for comparison
            $versionNum = [double]($version.Split('.')[0..1] -join '.')
            
            # Determine MSSQL version based on the version number
            $mssqlVersion = $null
            foreach ($mapping in $versionMapping) {
                if ($versionNum -ge $mapping.Version) {
                    $mssqlVersion = $mapping.Name
                    break
                }
            }
            
            # Create a custom object and add it to the array
            $data += [PSCustomObject]@{
                Hostname        = $hostname
                Instance        = $instance
                Edition         = $edition
                Version         = 'v' + $version
                "MSSQL Version" = $mssqlVersion
                "CPU Cores"     = $cpuCores
                "RAM"           = $ram
                "CPU Name"      = $cpuDetails
            }
        }
    }
}

Write-Host "[INFO] Data collection completed. Preparing summary..."
# Create summary data
$versionSummary = $data | Group-Object "MSSQL Version" | Select-Object Name, Count
$editionSummary = $data | Group-Object Edition | Select-Object Name, Count
# Calculate totals
$totalVersions = ($versionSummary | Measure-Object Count -Sum).Sum
Write-Host "[INFO] Creating Excel data..."
$TimeStamp = Get-Date -Format "HHmmss"
$excelFilePath = "SQL-REPORT" + "_" + (Get-Date).ToString("yyyyMMdd") + "_" + $TimeStamp + ".xlsx"
# Get current date in DD.MM.YYYY format
$currentDate = Get-Date -Format "dd.MM.yyyy"
$summarySheetName = "$currentDate-MAIN"
$dataSheetName = "$currentDate-DATA"
Write-Host "[INFO] Creating Summary worksheet ($summarySheetName)..."
# Export summary data first
$summaryData = @(
    [PSCustomObject]@{
        "Category" = "Total SQL Server Instances"
        "Count"    = $totalVersions
    }
)
$summaryData | Export-Excel -Path $excelFilePath -WorksheetName $summarySheetName -StartRow 1 -AutoSize
# Calculate the maximum row position based on the longer of the two data sets
$maxDataRows = [Math]::Max($versionSummary.Count, $editionSummary.Count)
$chartStartRow = $maxDataRows + 8  # Position both charts at the same row
Write-Host "[INFO] Generating Version Distribution chart..."
# Export version summary for chart
$versionSummary | Export-Excel -Path $excelFilePath -WorksheetName $summarySheetName -StartRow 5 -AutoSize `
    -ExcelChartDefinition @{
    Title     = "SQL Server Version Distribution"
    ChartType = "BarClustered"
    XRange    = "A6:A$($versionSummary.Count + 5)"
    YRange    = "B6:B$($versionSummary.Count + 5)"
    Row       = $chartStartRow
    Column    = 0
    Width     = 400
    Height    = 300
}

Write-Host "[INFO] Generating Edition Distribution chart..."
# Export edition summary for chart
$editionSummary | Export-Excel -Path $excelFilePath -WorksheetName $summarySheetName -StartRow 5 -StartColumn 5 -AutoSize `
    -ExcelChartDefinition @{
    Title     = "SQL Server Edition Distribution"
    ChartType = "BarClustered"
    XRange    = "E6:E$($editionSummary.Count + 5)"
    YRange    = "F6:F$($editionSummary.Count + 5)"
    Row       = $chartStartRow
    Column    = 6
    Width     = 600
    Height    = 300
}

Write-Host "[INFO] Creating SQL Inventory worksheet ($dataSheetName)..."
# Export main data with updated columns
$data | Select-Object Hostname, Instance, Edition, Version, "MSSQL Version", "CPU Cores", RAM, "CPU Name" | 
Write-Host "[INFO] Saving and opening Excel file in Excel..."
Export-Excel -Path $excelFilePath -WorksheetName $dataSheetName -AutoSize -TableName "SQLInventory" -TableStyle Medium2 -Show
Write-Host -ForegroundColor Green "[SUCCESS] Excel file '$excelFilePath' created successfully!"
Write-Host -ForegroundColor Magenta "[INFO] Summary includes:"
Write-Host -ForegroundColor Magenta "       - $totalVersions total SQL Server instances"
Write-Host -ForegroundColor Green "[SUCCESS] Script execution completed successfully!"
# Count existing .txt files
$files = Get-ChildItem -Path $folderPath -Filter "*.txt"
$fileCount = ($files | Measure-Object).Count
if ($fileCount -eq 0) {
    Write-Host -ForegroundColor Yellow "[INFO] No .txt files for deletion found in the folder."
}
else {
    Write-Host "[INFO] Found $fileCount .txt files ready to cleanup."
    Write-Host -NoNewline "Do you want to clean up the .txt files? (Y/N): "
    $confirmation = Read-Host
    if ($confirmation -notin @('Y', 'y')) {
        Write-Host -ForegroundColor Yellow "[INFO] Cleanup Operation denied by user."
    }
    else {
        try {
            # Delete all .txt files with progress
            $counter = 0
            $total = $files.Count
    
            foreach ($file in $files) {
                $counter++
                Write-Host -NoNewline "`rDeleting files... ($counter/$total)"
                Remove-Item $file.FullName -Force
            }
            Write-Host "`n"  # Add newline after progress
            Write-Host -ForegroundColor Green "[SUCCESS] Successfully deleted $fileCount .txt files!"
        }
        catch {
            Write-Host "`n"  # Add newline if error occurs during progress
            Write-Host -ForegroundColor Red "[ERROR] Failed to delete files: $($_.Exception.Message)"
        }
    }
}
Write-Host -ForegroundColor Cyan "[INFO] DONE! Cya! :)"
Write-Host -ForegroundColor Yellow "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")