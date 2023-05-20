Clear
Write-Warning "THIS SCRIPT CREATES VHDS AND USES VENTOY IN /NOUSBCheck MODE."
Write-Warning "THIS IS A POTENTIALLY DANGEROUS PROCESS"
Write-Warning "DO NOT USE THIS ON ANY MACHINE YOU CARE ABOUT."
Write-Warning "Use it on a test rig or blank Hyper-V virtual machine ONLY. You have been warned."
Write-Host "You should have already placed your ISO's in the ISO folder for this step to work."

$filename = "Ventoy"
$convertDir = Join-Path $PSScriptRoot "Programs"

# Get all ISO files
$isoFiles = Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath "ISO") -Filter "*.iso"

# Initialize total size
$totalSizeGB = 0

# Loop through each ISO file and add its size to the total
foreach ($isoFile in $isoFiles) {
    $totalSizeGB += $isoFile.Length / 1GB
}

# Round up the total size and add 1 GB for overhead
$sizeGB = [math]::Ceiling($totalSizeGB) + 1

# Create a new VHD
$vhdPath = Join-Path -Path $PSScriptRoot -ChildPath "$filename.vhd"
New-VHD -Path $vhdPath -SizeBytes ($sizeGB * 1GB) -Fixed

# Mount the VHD
$disk = Mount-VHD -Path $vhdPath -PassThru

# Initialize the disk, default is Online
Initialize-Disk $disk.Number -PartitionStyle MBR

# Create a new partition and assign a drive letter
$partition = New-Partition -DiskNumber $disk.Number -UseMaximumSize -AssignDriveLetter

Format-Volume -DriveLetter $partition.DriveLetter -FileSystem NTFS -NewFileSystemLabel "Ventoy" -AllocationUnitSize 4096 -Confirm:$false

$convertDir = Join-Path $PSScriptRoot "Programs"
$ventoy = (Get-ChildItem -Path "$convertDir\ventoy*" -Filter "Ventoy2Disk.exe" -Recurse).FullName

# Warn the user about the Ventoy process
Write-Warning "You are about to write Ventoy to a disk. It should be the VHD just created."
Write-Warning "Please check this in Disk Management. I am opening it for you now."
Start-Process -FilePath "diskmgmt.msc"

Write-Host "Disk Number:" $disk.Number
Write-Host "Size in GB:" ($disk.Size / 1GB)
Write-Warning "As this is a potentially damaging operation, please agree that this is the drive you wish to Ventoy"
 Write-Warning "Press 'Y' to continue. Any other key will exit the program."
$choice = Read-Host "Do you want to continue? (Y/N)"
if ($choice -ne 'Y') {
    Write-Host "Exiting..."
    Dismount-VHD -Path $vhdPath -Confirm:$false
    Remove-Item -Path $vhdPath
    exit
}

# Start Ventoy process and wait for it to complete
$physDrive = "/Drive:" + $partition.DriveLetter + ":"
$ventoyProcess = Start-Process -FilePath $ventoy -ArgumentList "VTOYCLI", "/I", $physDrive, "/NOUSBCheck" , "/FS:NTFS" -PassThru
$ventoyProcess.WaitForExit()



# Copy all ISO files to the new drive
$driveLetter = $partition.DriveLetter + ":\"
foreach ($isoFile in $isoFiles) {
    Copy-Item -Path $isoFile.FullName -Destination $driveLetter
    Write-Host "Copying:" $isoFile.FullName "to" $driveLetter
}

# Detach the VHD and wait for it to complete
Dismount-VHD -Path $vhdPath -Confirm:$false

# run qemu to convert to img
$qemu = (Get-ChildItem -Path "$convertDir\qemu-img*" -Filter "qemu-img.exe" -Recurse).FullName
$command = "& `"$qemu`" convert -O raw `"$vhdPath`" `"$imgPath`""
Invoke-Expression $command

Remove-Item -Path $vhdPath