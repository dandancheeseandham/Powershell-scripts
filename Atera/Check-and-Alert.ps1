[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Send-AteraAlert {
    param (
        [string]$AlertTitle,
        [string]$Message,
        [string]$Severity
    )

    $apiKey = "APIKEY"
    $headers = @{
        "X-API-KEY" = $apiKey
    }

    $deviceInfoUrl = "https://app.atera.com/api/v3/agents/machine/$ENV:COMPUTERNAME"
    $deviceInfo = Invoke-RestMethod -Uri $deviceInfoUrl -Headers $headers

    $alertPostUrl = "https://app.atera.com/api/v3/alerts"
    $alertBody = @{
        DeviceGuid       = $deviceInfo.Items.DeviceGuid
        CustomerID       = $deviceInfo.Items.CustomerID
        Title            = $AlertTitle
        MessageTemplate  = $Message
        Severity         = $Severity
        AlertCategoryID  = "Performance"
        Code             = 1
    }

    Invoke-RestMethod -Method Post -Uri $alertPostUrl -Headers $headers -Body $alertBody
}

function Test-SystemDriveEncrypted {
    $encryptionStatus = manage-bde -status C:
    $encryptionStatus -match "Protection Status: Protection (On|Off)"
}

function Test-RdpEnabled {
    $rdpRegPath = 'HKLM:\System\CurrentControlSet\Control\Terminal Server'
    (Get-ItemProperty -Path $rdpRegPath).fDenyTSConnections -eq 0
}

function Test-RdpPortOpen {
    param (
        [string]$IpAddress,
        [int]$Port = 3389
    )

    $tcpClient = New-Object Net.Sockets.TcpClient
    $tcpClient.Connect($IpAddress, $Port)
    $tcpClient.Connected
}

if (-not (Test-SystemDriveEncrypted)) {
    Write-Output "System Drive Not Encrypted"
    $alertTitle = "System Drive Not Encrypted on $ENV:COMPUTERNAME"
    $message = "The system drive is not encrypted, please enable BitLocker or another encryption solution. Remove this check if required."
    Send-AteraAlert -AlertTitle $alertTitle -Message $message -Severity "Critical"
}


if (Test-RdpEnabled) {
    Write-Output "RDP Enabled"
    $alertTitle = "RDP Enabled on $ENV:COMPUTERNAME"
    $message = "RDP is enabled please check to see if this is required if not disable it. Remove this check if required"
    Send-AteraAlert -AlertTitle $alertTitle -Message $message -Severity "Critical"
}

$ipInfoUrl = "http://ipinfo.io/json"
$ipAddress = (Invoke-RestMethod -Uri $ipInfoUrl).ip
$rdpPort = 3389

if (Test-RdpPortOpen -IpAddress $ipAddress -Port $rdpPort) {
    Write-Output "Port $rdpPort is operational"
    $alertTitle = "RDP Port open on router from $ENV:COMPUTERNAME"
    $message = "RDP port is open please check to see if this is required if not disable it. Remove this check if required"
    Send-AteraAlert -AlertTitle $alertTitle -Message $message -Severity "Critical"
} else {
    Write-Output "Port $rdpPort is closed, You may need to contact your IT team to open it."
}
