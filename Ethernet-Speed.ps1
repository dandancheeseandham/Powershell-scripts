# File     : Ethernet-Speed.ps1
# Effect   : Check the connection speed on connected Ethernet
# Use-case : Quick check to see if 100Mbit / 1Gbit
# Run As   : Administrator
# Author   : Dan White

$wmi = Get-WmiObject -Class Win32_NetworkAdapter -Filter "NetConnectionID='Ethernet'" | Select-Object Speed
$linkSpeed = [Math]::Round(($wmi.Speed / 1e+6), 2)
Write-Host "The link speed of the active Ethernet connection is $linkSpeed Mbit."