param(
    [string]$server,
    [string]$changeNumber,
    [string]$accountName
)

# Output to console for visibility in the separate window
Write-Host "Starting update process for $server..."

try {
    # Record start time to filter update history
    $startTime = Get-Date

    # Connect to the server and apply updates
    Write-Host "Connecting to $server and installing updates..."
    Invoke-Command -ComputerName $server -ScriptBlock {
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-Host "Installing PSWindowsUpdate module on $server..."
            Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser
        }
        Import-Module PSWindowsUpdate
        Write-Host "Applying updates on $server..."
        Get-WUInstall -AcceptAll -IgnoreReboot -Confirm:$false
    }

    # Get update history since start time
    Write-Host "Retrieving update history for $server..."
    $history = Invoke-Command -ComputerName $server -ScriptBlock {
        Get-WUHistory | Where-Object { $_.Date -ge $using:startTime }
    }

    # Process update history
    $installedUpdates = $history | Where-Object { $_.ResultCode -eq "Succeeded" } | Select-Object -ExpandProperty KBArticleID
    $failedUpdates = $history | Where-Object { $_.ResultCode -ne "Succeeded" } | Select-Object -ExpandProperty KBArticleID

    # Check if reboot is required
    Write-Host "Checking reboot status for $server..."
    $rebootRequired = Invoke-Command -ComputerName $server -ScriptBlock {
        (Get-WURebootStatus -Silent).RebootRequired
    }

    # Determine status
    $statusValue = if ($failedUpdates) { "Partial Success" } elseif ($installedUpdates) { "Success" } else { "No Updates" }
    $errorMessage = if ($failedUpdates) { "Failed updates: $($failedUpdates -join ',')" } else { "" }

    # Create status object
    $status = [PSCustomObject]@{
        ChangeNumber     = $changeNumber
        AccountName      = $accountName
        ServerName       = $server
        Status           = $statusValue
        ErrorMessage     = $errorMessage
        RebootRequired   = if ($rebootRequired) { "Yes" } else { "No" }
        UpdatesInstalled = $installedUpdates -join ","
    }

    # Export status to CSV
    $status | Export-Csv -Path "status_$server.csv" -NoTypeInformation
    Write-Host "Status for $server written to status_$server.csv"

    # Prompt for reboot if required
    if ($rebootRequired) {
        Write-Host "Updates installed on $server. Reboot required. Press Enter to reboot now or close this window to skip."
        Read-Host
        Write-Host "Rebooting $server..."
        Invoke-Command -ComputerName $server -ScriptBlock { Restart-Computer -Force }
    } else {
        Write-Host "No reboot required for $server. Update process complete."
    }
} catch {
    # Handle errors (e.g., server unreachable)
    Write-Host "Error occurred while processing $server : $($_.Exception.Message)"
    $status = [PSCustomObject]@{
        ChangeNumber     = $changeNumber
        AccountName      = $accountName
        ServerName       = $server
        Status           = "Error"
        ErrorMessage     = $_.Exception.Message
        RebootRequired   = "No"
        UpdatesInstalled = ""
    }
    $status | Export-Csv -Path "status_$server.csv" -NoTypeInformation
}