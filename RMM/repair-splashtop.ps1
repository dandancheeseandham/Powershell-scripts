# File     : repair-splashtop.ps1
# Effect   : Remote deployment script to download and restore Splashtop Streamer via https://my.splashtop.com/csrs/win
# Use-case : Where the Atera (or other RMM) is still working, but Splashtop is not, or has been removed.
# Run As   : Administrator
# Author   : Dan White

# Function to download the file and return the downloaded file's path
function Download-File {
    param (
        [string]$url,
        [string]$tempFolderPath = $env:TEMP
    )
    
    # Resolve redirect
    $req = [System.Net.HttpWebRequest]::Create($url)
    $req.AllowAutoRedirect = $false
    $resp = $req.GetResponse()
    $destinationUrl = $resp.GetResponseHeader("Location")
    $fileName = [System.IO.Path]::GetFileName($destinationUrl)
    
    $tempFilePath = Join-Path -Path $tempFolderPath -ChildPath $fileName
    Invoke-WebRequest -Uri $destinationUrl -OutFile $tempFilePath -Verbose
    
    return $tempFilePath
}

# Function to install the downloaded file
function Install-SplashtopStreamer {
    param (
        [string]$filePath
    )
    
    Write-Host "Installing Splashtop Streamer..."
    Start-Process -FilePath $filePath -ArgumentList "/silent" -Wait -Verbose
    Write-Host "Installation complete."
}

# Main script
try {
    $splashtopUrl = "https://my.splashtop.com/csrs/win"
    
    Write-Host "Downloading Splashtop Streamer..."
    $downloadedFile = Download-File -url $splashtopUrl
    Write-Host "File downloaded: $($downloadedFile)"
    
    Install-SplashtopStreamer -filePath $downloadedFile

} catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
}
