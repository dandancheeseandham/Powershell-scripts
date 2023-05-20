function GetProg {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ProgDir,

        [Parameter(Mandatory=$false)]
        [string]$downloadUri
    )

    # Create the directory if it doesn't exist
    if (!(Test-Path -Path $ProgDir)) {
        New-Item -ItemType Directory -Path $ProgDir | Out-Null
        Write-Host "Created directory $ProgDir" -ForegroundColor Green
    }

    # If downloadUri is not null or empty, download and extract the program
    if (![string]::IsNullOrEmpty($downloadUri)) {
        $fileName = Split-Path -Path $downloadUri -Leaf
        Write-Host "Downloading $fileName from $downloadUri" -ForegroundColor Green
        Invoke-WebRequest -Uri $downloadUri -OutFile "$ProgDir\$fileName"
        
        Write-Host "Decompressing $fileName" -ForegroundColor Green
        Expand-Archive -Path "$ProgDir\$fileName" -DestinationPath $ProgDir
        
        Write-Host "Removing downloaded file $fileName" -ForegroundColor Green
        Remove-Item "$ProgDir\$fileName"
    }
}


Clear
Write-Host "These two programs automate the creation of a VHD, putting Ventoy on it and copying ISO's"
Write-Host "This program only needs to be run once, it creates folders and gets the required programs for the operation and is safe"
Write-Warning "THE NEXT SCRIPT CREATES VHDS AND USES VENTOY IN /NOUSBCheck MODE."
Write-Warning "THIS IS A POTENTIALLY DANGEROUS PROCESS"
Write-Warning "DO NOT USE THIS ON ANY MACHINE YOU CARE ABOUT."
Write-Warning "Use it on a test rig or blank Hyper-V virtual machine ONLY. You have been warned."

# Create the Programs directory and download Ventoy if it doesn't exist
$ventoyUri = "https://github.com/ventoy/Ventoy/releases/download/v1.0.91/ventoy-1.0.91-windows.zip"
GetProg -ProgDir (Join-Path -Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) -ChildPath "Programs") -downloadUri $ventoyUri

# Create the qemu-img directory and download qemu-img if it doesn't exist
$qemuUri = "https://cloudbase.it/downloads/qemu-img-win-x64-2_3_0.zip"
GetProg -ProgDir (Join-Path -Path (Join-Path $PSScriptRoot "Programs") -ChildPath "qemu-img") -downloadUri $qemuUri

# Create the ISO directory if it doesn't exist
GetProg -ProgDir (Join-Path -Path $PSScriptRoot -ChildPath "ISO")

Write-Host "To use, place ISO's in the ISO folder and then run the create VHD program."