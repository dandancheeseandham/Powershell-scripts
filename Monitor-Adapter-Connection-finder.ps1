$adapterTypes = @{ #https://www.magnumdb.com/search?q=parent:D3DKMDT_VIDEO_OUTPUT_TECHNOLOGY
    '-2' = 'Unknown'
    '-1' = 'Unknown'
    '0' = 'VGA'
    '1' = 'S-Video'
    '2' = 'Composite'
    '3' = 'Component'
    '4' = 'DVI'
    '5' = 'HDMI'
    '6' = 'LVDS'
    '8' = 'D-Jpn'
    '9' = 'SDI'
    '10' = 'DisplayPort (external)'
    '11' = 'DisplayPort (internal)'
    '12' = 'Unified Display Interface'
    '13' = 'Unified Display Interface (embedded)'
    '14' = 'SDTV dongle'
    '15' = 'Miracast'
    '16' = 'Internal'
    '2147483648' = 'Internal'
}

$monitors = Get-WmiObject -Namespace root/wmi -Class WmiMonitorID
$connections = Get-WmiObject -Namespace root/wmi -Class WmiMonitorConnectionParams
$videoControllers = Get-WmiObject -Class Win32_VideoController

$result = @()

foreach ($controller in $videoControllers) {
    foreach ($monitor in $monitors) {
        $connection = $connections | Where-Object { $_.InstanceName -eq $monitor.InstanceName }
        $connectionType = $adapterTypes["$($connection.VideoOutputTechnology)"]

        if ($connectionType -eq $null) { $connectionType = 'Unknown' }

        $manufacturer = $monitor.ManufacturerName
        $name = $monitor.UserFriendlyName

        if ($manufacturer -ne $null) { $manufacturer = [System.Text.Encoding]::ASCII.GetString($manufacturer -ne 0) }
        if ($name -ne $null) { $name = [System.Text.Encoding]::ASCII.GetString($name -ne 0) }

        if (($manufacturer -ne $null) -or ($name -ne $null)) {
            $result += New-Object -TypeName PSObject -Property @{
                "MonitorName" = "$manufacturer $name"
                "AdapterName" = $controller.Name
                "Connection"  = $connectionType
            }
        }
    }
}

$result
