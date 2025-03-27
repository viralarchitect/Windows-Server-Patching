# Windows Server Patching Script

This is a robust PowerShell solution for patching Windows Servers based on your requirements. The script will take a list of IPs or hostnames, apply Windows updates asynchronously, report the status in a CSV file (including a user-provided change number and account name), prompt for reboots, display activity in separate PowerShell windows for each server, and handle errors gracefully. The solution consists of two scripts: a main script to orchestrate the process and a child script to handle updates on each server.

## Technical Requirements

Based on your query, here are the detailed technical requirements for the PowerShell script:

### 1. Input Handling:

Accept an arbitrary list of server IPs or hostnames, either as a comma-separated string or from a file path.

### 2. Asynchronous Execution:

Apply updates on multiple servers simultaneously using separate PowerShell processes.

### 3. Update Application:

Install all available Windows updates on each server using the PSWindowsUpdate module.

### 4. Reboot Prompt:

Prompt the user in each server's PowerShell window before rebooting, allowing the user to confirm or skip the reboot.

### 5. Status Reporting:

1. Report the patch installation status for each server, including:
   1. Server name
   2. Status (e.g., Success, Partial Success, Error, No Updates)
   3. Error message (if applicable)
   4. List of installed updates
   5. Reboot requirement (Yes/No)
2. Tie the status to a user-provided change number and account name, included as columns in the CSV output.

### 6. Separate PowerShell Windows:

Open a new PowerShell window for each server to display its update activity.

### 7. Error Handling:

Gracefully handle errors (e.g., unreachable servers, update failures) without stopping the script.

Log errors in the status output and continue processing other servers.

### 8. Logging and Output:

Generate a CSV file consolidating the status of all servers, including the change number and account name.

## Main Script (Start-Patching.ps1)

This script collects user inputs, initiates the update process for each server asynchronously, and compiles the results into a single CSV file.

## Child Script (UpdateServer.ps1)

This script handles the update process for a single server, including applying updates, checking for reboots, and logging the status.

## How to Use the Scripts

### Preparation:

Save the main script as Start-Patching.ps1 and the child script as UpdateServer.ps1 in the same directory.

Ensure you have administrative privileges and that PowerShell remoting is enabled on all target servers (Enable-PSRemoting).

### Execution:

Run Start-Patching.ps1 in a PowerShell console.

Enter the change number (e.g., "CHG12345").

Enter the account name (e.g., "adminuser").

Enter the server list (e.g., "server1,server2,server3" or a file path like "C:\servers.txt" containing one server per line).

### Monitoring:

A new PowerShell window will open for each server, showing the update activity (logged to log_<server>.txt).

If a reboot is required, the window will prompt you to press Enter to reboot or close the window to skip.

### Completion:

Once all update windows are closed, return to the main script window and press Enter.

The script will compile all status files (status_<server>.csv) into patch_status.csv.