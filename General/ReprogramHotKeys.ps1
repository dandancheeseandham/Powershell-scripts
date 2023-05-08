# File     : ReprogramKeys.ps1
# Effect   : Reprograms Hotkeys on an Microsoft Natural Keyboard Pro
# Use-case : Keeping alive old Microsoft keyboards, including my 22 year old Microsoft Natural Keyboard Pro.
# Run As   : Administrator
# Author   : Dan White

# This script requires administrator privileges.
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Exit
}

# Hashtable to store the key mappings.
$keyMappings = @{
#    "Back"        = "C:\Path\To\Back.exe"
#    "Forward"     = "C:\Path\To\Forward.exe"
#    "Stop"        = "C:\Path\To\Stop.exe"
#    "Refresh"     = "C:\Path\To\Refresh.exe"
#    "Search"      = "C:\Path\To\Search.exe"
#    "Favorites"   = "C:\Path\To\Favorites.exe"
#    "WebHome"     = "C:\Path\To\WebHome.exe"
#    "MyComputer"  = "explorer.exe"
#    "Mail"        = "C:\Path\To\Mail.exe"
#    "Media"       = "C:\Path\To\Media.exe"
#    "Calculator"  = "calc.exe"
#    "Mute"        = "C:\Path\To\Mute.exe"
#    "VolumeDown"  = "C:\Path\To\VolumeDown.exe"
#    "VolumeUp"    = "C:\Path\To\VolumeUp.exe"
#    "PlayPause"   = "C:\Path\To\PlayPause.exe"
#    "StopMedia"   = "C:\Path\To\StopMedia.exe"
#    "PrevTrack"   = "C:\Path\To\PrevTrack.exe"
#    "NextTrack"   = "C:\Path\To\NextTrack.exe"
}

# Registry keys for the Microsoft Natural Keyboard Pro buttons.
$registryKeys = @{
    "Back"        = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\1"
    "Forward"     = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\2"
    "Stop"        = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\3"
    "Refresh"     = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\4"
    "Search"      = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\5"
    "Favorites"   = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\6"
    "WebHome"     = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\7"
    "MyComputer"  = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\17"
    "Mail"        = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\15"
    "Media"       = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\16"
    "Calculator"  = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\18"
    "Mute"        = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\20"
    "VolumeDown"  = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\21"
    "VolumeUp"    = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\22"
    "PlayPause"   = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\14"
    "StopMedia"   = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\13"
    "PrevTrack"   = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\11"
    "NextTrack"   = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AppKey\12"
}

foreach ($key in $keyMappings.Keys) {
    $registryKeyPath = $registryKeys[$key]
    $newProgramPath = $keyMappings[$key]

    # Check if the specified program exists.
    if (-not (Test-Path $newProgramPath) -and $newProgramPath -ne "explorer.exe" -and $newProgramPath -ne "calc.exe") {
        Write-Host "Program not found at the specified path for $key. Skipping." -ForegroundColor Yellow
        continue
    }

    # Set the value of the specified key.
    Set-ItemProperty -Path $registryKeyPath -Name "ShellExecute" -Value $newProgramPath
    Write-Host "$key button has been successfully reprogrammed to open $($keyMappings[$key])." -ForegroundColor Green
}
