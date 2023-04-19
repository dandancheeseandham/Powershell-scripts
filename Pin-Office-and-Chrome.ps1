# File     : Pin-Office-And-Chrome.ps1
# Effect   : Pins various applications on the current user.
# Use-case : Best used on new computers, or at least without these apps already pinned  
#            so remove any first if you are using it on a computer that's already had use.
# Run As   : Current User
# Author   : Dan White

$url = "http://www.technosys.net/download.aspx?file=syspin.exe"
$output = "$env:temp\syspin.exe"
Invoke-WebRequest -Uri $url -OutFile $output

function PinToTaskbar($appPath) {
    $sysPinPath = $output
    if (-not (Test-Path $sysPinPath)) {
        
    }
    & $sysPinPath $appPath 5386
}

$applications = @(
    "C:\Program Files\Google\Chrome\Application\chrome.exe",   			    # Google Chrome
    "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",   	# 32-bit Microsoft Word
    "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",     	# 32-bit  Microsoft Excel
    "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",    # 32-bit  Microsoft Outlook
    "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",   	    # 64-bit Microsoft Word
    "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",     	    # 64-bit  Microsoft Excel
    "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"    	    # 64-bit  Microsoft Outlook

)

foreach ($app in $applications) {
    $appPath = $app
    if (Test-Path $appPath) {
        PinToTaskbar $appPath
    } else {
        Write-Host "Application not found: $appPath"
    }
}