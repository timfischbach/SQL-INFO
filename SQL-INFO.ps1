#////////////////////////////////////////////
# SQL-INFO.ps1 by Tim Fischbach
#////////////////////////////////////////////
$hostname = $env:COMPUTERNAME
$instances = (get-itemproperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server').InstalledInstances
$Output = @()
# Only proceed with system info gathering if instances are found
if ($instances) {
    # Get CPU Count from Registry
    $CPUDetails = (Get-ItemProperty "HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0").ProcessorNameString
    $CPUCores = (Get-ChildItem "HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor").Count

    # Get RAM in GB from Registry
    $TotalRAM = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1gb
    
    # Add system info as first line
    $Output += "System Info: CPU: $CPUDetails, Cores: $CPUCores, RAM: $TotalRAM GB"
    $Output += "----------------------------------------"

    # Process SQL instances
    foreach ($i in $instances) {
        $instanceRegPath = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL').$i
        $Edition = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$instanceRegPath\Setup").Edition
        $Version = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$instanceRegPath\Setup").Version
        $Output += "Instance: $i, Edition: $Edition, Version: $Version"
    }
}

# Only create file if we have data
if ($Output.Count -gt 0) {
    $TimeStamp = Get-Date -Format "HHmmss"
    $filename = $hostname + "_" + (Get-Date).ToString("yyyyMMdd") + "_" + $TimeStamp + ".txt"
    $FilePath = "\\MY\NETWORK\DRIVE\txt\$filename"

    $Output | Out-File -FilePath $FilePath

    # Output statistics
    Write-Output "[INFO] SQL Server instances found!"
} else {
    Write-Output "[INFO] No SQL Server instances found."
}