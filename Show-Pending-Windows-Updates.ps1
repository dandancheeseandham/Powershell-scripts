# File     : Show-Pending-Windows-Updates.ps1
# Effect   : Shows pending Windows updates in full, or as KB numbers only.
# Run As   : Administrator
# Author   : Dan White

# Set to 'Full' for full description, or 'KB' for KB numbers only
$outputOption = 'Full'

# Create a new Windows Update session
$windowsUpdateSession = New-Object -ComObject Microsoft.Update.Session

# Set the Client Application ID for the session
$windowsUpdateSession.ClientApplicationID = 'MSDN Sample Script'

# Create an Update Searcher object for the session
$updateSearcher = $windowsUpdateSession.CreateUpdateSearcher()

# Define the search criteria for pending updates
$searchCriteria = "IsInstalled=0 and Type='Software' and IsHidden=0"

try {
    # Search for pending updates using the defined criteria
    $pendingUpdatesResult = $updateSearcher.Search($searchCriteria)
} catch [System.Runtime.InteropServices.COMException] {
    if ($_.Exception.ErrorCode -eq 0x80240438) {
        Write-Host "Error HRESULT: 0x80240438 - Issue connecting to the Windows Update service."
        Write-Host "We suggest that you ensure you are running as an administrator, disable AV, firewalls, or other security software, and reset the Windows Update components, then reboot."
        exit
    } else {
        throw
    }
}

# Get the titles or KB numbers of the pending updates based on the output option
if ($outputOption -eq 'Full') {
    $pendingUpdatesOutput = $pendingUpdatesResult.Updates | Select-Object -ExpandProperty Title
} elseif ($outputOption -eq 'KB') {
    $pendingUpdatesOutput = $pendingUpdatesResult.Updates | ForEach-Object {
        if ($_.Title -match 'KB(\d+)') {
            'KB' + $matches[1]
        }
    }
} else {
    Write-Host "Invalid output option. Please set the output option to 'Full' or 'KB'."
    exit
}

# Output the list of pending updates titles or KB numbers
$pendingUpdatesOutput
