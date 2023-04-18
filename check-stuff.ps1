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

# get the computername and date so it can be part of the output file
$compname = $env:computername
$Today = get-date -Format "dd.mm.yyyy"
# change working directory and create output folder
c:
cd\
$path = "$env:TEMP"
If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
}

Write-Output ("Check all the stuff script")
Write-Output ($compname)
Get-Date

# Check Last Boot Time
Write-DecoratedString -InputString "Last boot time"
Get-CimInstance -ClassName Win32_OperatingSystem | Select -Exp LastBootUpTime


# Check for driver issues
Write-DecoratedString -InputString "Checking for driver issues"
Write-Output ("Result:")
$DeviceState = Get-WmiObject -Class Win32_PnpEntity -ComputerName localhost -Namespace Root\CIMV2 | Where-Object {$_.ConfigManagerErrorCode -gt 0
}

$DevicesInError = foreach($Device in $DeviceState){
 $Errortext = switch($device.ConfigManagerErrorCode){
0 {"This device is working properly."}
1 {"This device is not configured correctly."}
2 {"Windows cannot load the driver for this device."}
3 {"The driver for this device might be corrupted, or your system may be running low on memory or other resources."}
4 {"This device is not working properly. One of its drivers or your registry might be corrupted."}
5 {"The driver for this device needs a resource that Windows cannot manage."}
6 {"The boot configuration for this device conflicts with other devices."}
7 {"Cannot filter."}
8 {"The driver loader for the device is missing."}
9 {"This device is not working properly because the controlling firmware is reporting the resources for the device incorrectly."}
10 {"This device cannot start."}
11 {"This device failed."}
12 {"This device cannot find enough free resources that it can use."}
13 {"Windows cannot verify this device's resources."}
14 {"This device cannot work properly until you restart your computer."}
15 {"This device is not working properly because there is probably a re-enumeration problem."}
16 {"Windows cannot identify all the resources this device uses."}
17 {"This device is asking for an unknown resource type."}
18 {"Reinstall the drivers for this device."}
19 {"Failure using the VxD loader."}
20 {"Your registry might be corrupted."}
21 {"System failure: Try changing the driver for this device. If that does not work, see your hardware documentation. Windows is removing this device."}
22 {"This device is disabled."}
23 {"System failure: Try changing the driver for this device. If that doesn't work, see your hardware documentation."}
24 {"This device is not present, is not working properly, or does not have all its drivers installed."}
25 {"Windows is still setting up this device."}
26 {"Windows is still setting up this device."}
27 {"This device does not have valid log configuration."}
28 {"The drivers for this device are not installed."}
29 {"This device is disabled because the firmware of the device did not give it the required resources."}
30 {"This device is using an Interrupt Request (IRQ) resource that another device is using."}
31 {"This device is not working properly because Windows cannot load the drivers required for this device."}
}
[PSCustomObject]@{
ErrorCode = $device.ConfigManagerErrorCode
ErrorText = $Errortext
Device = $device.Caption
Present = $device.Present
Status = $device.Status
StatusInfo = $device.StatusInfo
}
}

if(!$DevicesInError){
write-host "Healthy. No driver issues detected."
} else {
$DevicesInError
}


# Check Last CHKDSK
Write-DecoratedString -InputString "Last CHKDSK"

try {
    $eventschkdsk = Get-EventLog -LogName Application -InstanceId 26226 -Source Chkdsk -ErrorAction Stop
    if ($eventschkdsk) {
        $eventschkdsk | Select-Object -ExpandProperty Message
    } else {
        Write-Output "No CHKDSK id 26226 in Application Log"
    }
} catch {
    Write-Output "No CHKDSK id 26226 in Application Log"
}


$eventschkdskwin = Get-Winevent -FilterHashTable @{logname="Application"; id="1001"}| ?{$_.providername –match "wininit"} | fl timecreated, message
if ($eventschkdskwin) {
    $eventschkdskwin | Format-Table -AutoSize -Wrap
} else {
    Write-Output "No CHKDSK id 1001 in Application Log"
}

# Show all wifi profiles and passwords
# Would be nice to send this to IT Glue
Write-DecoratedString -InputString "Wifi profiles and passwords"
(netsh wlan show profiles) | Select-String "\:(.+)$" | %{$name=$_.Matches.Groups[1].Value.Trim(); $_} | %{(netsh wlan show profile name="$name" key=clear)} | Select-String "Key Content\W+\:(.+)$" | %{$pass=$_.Matches.Groups[1].Value.Trim(); $_} | %{[PSCustomObject]@{ PROFILE_NAME=$name;PASSWORD=$pass }} | Format-Table -AutoSize


# Would be nice to send this to IT Glue
Write-DecoratedString -InputString "Why did the system reboot last?"
$events = Get-WinEvent -FilterHashtable @{logname = 'System'; id = 1074, 6005, 6006, 6008} -MaxEvents 6

if ($events) {
    $events | Format-Table -Wrap
} else {
    Write-Output "No matches found"
}


# Last errors in System Log
Write-DecoratedString -InputString "Last 32 errors in System Log"
$events = Get-EventLog -LogName System -EntryType Error -Newest 32

if ($events) {
    $events | Format-Table -AutoSize -Wrap
} else {
    Write-Output "No errors found"
}

# Last errors in Application Log
Write-DecoratedString -InputString "Last 32 errors in Application Log"
$eventsapp =  Get-EventLog -LogName Application -EntryType Error -Newest 32
if ($eventsapp) {
    $eventsapp | Format-Table -AutoSize -Wrap
} else {
    Write-Output "No errors found"
}


# Any other notable events in the Event Log
Write-DecoratedString -InputString "Any other notable events in the Event Log"
try {Get-WinEvent -FilterHashtable @{logname = 'System'; id = 1001, 4740, 4724, 4728,4732,4756, 4724, 4625} -ErrorAction Stop -MaxEvents 24 | Format-Table -wrap
    }
catch [Exception] {
        if ($_.Exception -match "No events were found that match the specified selection criteria") {
        Write-Output "No events found";
                 }
    }

# Show BSOD's using BlueScreenView
Write-DecoratedString -InputString "BSOD's"
###bluescreen
 $scriptName = "Blue Screen View"
 $computerName = (get-wmiObject win32_computersystem).name
 $computerDomain = (get-wmiObject win32_computersystem).domain
 if($computerdomain -notlike '*.*'){ #if there's no period in the domain, (workgroup)
	$computerDomain = "$computerDomain.local"	
 }
  $messageBody = "----Blue Screen View Results----`r`n"
 $url = "http://www.runpcrun.com/files/BlueScreenView.exe"
 $filename = "BlueScreenView.exe"
 $client = New-Object System.Net.WebClient
 $client.DownloadFile($url, "$env:temp\$filename")
 Start-Process -FilePath "$env:temp\$filename" -ArgumentList "/stab","$env:temp\crashes.txt","/sort 2","/sort ~1"""
 Get-Content $env:temp\crashes.txt


# Pending Patches
Write-DecoratedString -InputString "Patches Pending"
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateupdateSearcher()
$Updates = @($UpdateSearcher.Search("IsHidden=0 and IsInstalled=0").Updates)
$Updates | Select-Object Title



# Is a Reboot Pending?
Write-DecoratedString -InputString "Is a Reboot Pending? (and why)"
$ErrorActionPreference = 'Stop'

function Test-RegistryKey {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Key
    )

    if (Get-Item -Path $Key -ErrorAction Ignore) {
        $true
    }
}

function Test-RegistryValue {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Key,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Value
    )

    if (Get-ItemProperty -Path $Key -Name $Value -ErrorAction Ignore) {
        $true
    }
}

function Test-RegistryValueNotNull {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Key,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Value
    )

    if (($regVal = Get-ItemProperty -Path $Key -Name $Value -ErrorAction Ignore) -and $regVal.($Value)) {
        $true
    }
}

    $tests = @(
        { Test-RegistryKey -Key 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending' }
        { Test-RegistryKey -Key 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress' }
        { Test-RegistryKey -Key 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' }
        { Test-RegistryKey -Key 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending' }
        { Test-RegistryKey -Key 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting' }
        { Test-RegistryValueNotNull -Key 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Value 'PendingFileRenameOperations' }
        { Test-RegistryValueNotNull -Key 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Value 'PendingFileRenameOperations2' }
        { (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Updates' -Name 'UpdateExeVolatile' | Select-Object -ExpandProperty UpdateExeVolatile) -ne 0 }
        { Test-RegistryValue -Key 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -Value 'DVDRebootSignal' }
        { Test-RegistryKey -Key 'HKLM:\SOFTWARE\Microsoft\ServerManager\CurrentRebootAttemps' }
        { Test-RegistryValue -Key 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon' -Value 'JoinDomain' }
        { Test-RegistryValue -Key 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon' -Value 'AvoidSpnSet' }
        {
            (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName').ComputerName -ne
            (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName').ComputerName
        }
        {
            if (Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services\Pending') {
                $true
            }
        }
    )

$isPendingReboot = $false

Write-Output "Pending Reboot"

foreach ($test in $tests) {
    if (& $test) {
        Write-Output $test
        $isPendingReboot = $true
        break
    }
}

[pscustomobject]@{
    ComputerName    = $env:COMPUTERNAME
    IsPendingReboot = $isPendingReboot
}

# Programs installed in the last week
Write-DecoratedString -InputString "Programs installed in the last week"

$currentDate = Get-Date
$weekAgo = $currentDate.AddDays(-7)
# Get the list of installed programs from the registry
$installedPrograms = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* `
    -ErrorAction SilentlyContinue `
    | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
# Filter the programs installed in the last week
$recentlyInstalledPrograms = $installedPrograms | Where-Object {
    $_.InstallDate -ne $null -and `
    [datetime]::ParseExact($_.InstallDate, "yyyyMMdd", $null) -ge $weekAgo
}
# Display the recently installed programs
if ($recentlyInstalledPrograms -ne $null) {
    Write-Host "Programs installed in the last week:" -ForegroundColor Green
    $recentlyInstalledPrograms | Format-Table -AutoSize
} else {
    Write-Host "No programs were installed in the last week." -ForegroundColor Yellow
}


# Showing any higher (>5%) CPU processes 
Write-DecoratedString -InputString "Any high CPU usage? Set to 5% of any logical core."
$Threshold = 5
$ProcessUsage = Get-Counter -Counter "\Processor(_Total)\% Processor Time" -SampleInterval 1 -MaxSamples 1
$CPUCount = (Get-WmiObject -Class Win32_ComputerSystem).NumberOfLogicalProcessors
$HighUsageProcesses = Get-WmiObject -Class Win32_PerfFormattedData_PerfProc_Process |
    Where-Object { ($_.PercentProcessorTime / $CPUCount) -gt $Threshold } |
    Select-Object -Property IDProcess, Name, @{Name="PercentProcessorTimePerCore"; Expression={($_.PercentProcessorTime / $CPUCount)}}

if ($HighUsageProcesses) {
    $HighUsageProcesses | Format-Table -AutoSize
} else {
    Write-Output "No processes found using more than $Threshold% of any logical CPU."
}
