# Get user inputs
$changeNumber = Read-Host "Enter change number"
$accountName = Read-Host "Enter account name"
$serversInput = Read-Host "Enter list of servers (comma-separated) or path to file"

# Determine if input is a file path or a list of servers
if (Test-Path $serversInput) {
    $serverList = Get-Content $serversInput
} else {
    $serverList = $serversInput -split ','
}

# Start a separate PowerShell process for each server
foreach ($server in $serverList) {
    $server = $server.Trim()  # Remove any whitespace
    Start-Process powershell.exe -ArgumentList "-File UpdateServer.ps1 -server $server -changeNumber $changeNumber -accountName $accountName" -RedirectStandardOutput "log_$server.txt"
}

# Wait for user to confirm all processes are complete
Write-Host "Started update processes for all servers. Please monitor the individual windows."
Write-Host "When all update windows are closed, press Enter to compile the status CSV."
Read-Host

# Compile the status CSVs into one file
$allStatus = @()
foreach ($server in $serverList) {
    $server = $server.Trim()
    $statusFile = "status_$server.csv"
    if (Test-Path $statusFile) {
        $status = Import-Csv $statusFile
        $allStatus += $status
    } else {
        Write-Warning "Status file for $server not found."
    }
}

# Export to final CSV
if ($allStatus) {
    $allStatus | Export-Csv -Path "patch_status.csv" -NoTypeInformation
    Write-Host "Patch status compiled in patch_status.csv"
} else {
    Write-Warning "No status files were found to compile."
}