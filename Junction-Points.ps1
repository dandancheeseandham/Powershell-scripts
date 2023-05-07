function Set-JunctionPoints {
    param (
        [string]$OneDriveFolderName = "OneDrive",
        [hashtable]$SourceToTarget
    )

    $usr = Get-WmiObject -Class Win32_ComputerSystem | Select -ExpandProperty UserName
    $UserName = $usr.Split('\')[-1]

    function Test-Junction {
        param (
            [string]$Path
        )

        return (Get-Item -Path $Path -Force -ErrorAction SilentlyContinue).Attributes -match 'ReparsePoint'

    }

    foreach ($source in $SourceToTarget.Keys) {
        $target = "C:\Users\$UserName\$OneDriveFolderName\$($SourceToTarget[$source])"

        if (Test-Path -Path $source -PathType Container) {
            if (Test-Path -Path $target -PathType Leaf) {
                Write-Host "ERROR: File exists at target path: $target" -ForegroundColor Red
                continue
            }
			if (Test-Path -Path $target -PathType Container) -and (-not (Test-Junction -Path $target)) {
                Write-Host "ERROR: Directory exists at target path: $target" -ForegroundColor Red
                continue
            }

            if (Test-Junction -Path $target) {
                Write-Host "WARNING: Junction point already exists at target path: $target" -ForegroundColor Yellow
            } else {
                New-Item -ItemType Junction -Path $target -Value $source
                Write-Host "Created junction point: $source -> $target"
            }
        } else {
            Write-Host "ERROR: Source directory not found: $source" -ForegroundColor Red
        }
    }
}

function Remove-JunctionPoints {
    param (
        [string]$OneDriveFolderName = "OneDrive",
        [array]$Targets
    )

    $usr = Get-WmiObject -Class Win32_ComputerSystem | Select -ExpandProperty UserName
    $UserName = $usr.Split('\')[-1]

    function Test-Junction {
        param (
            [string]$Path
        )

        return (Get-Item -Path $Path -Force).Attributes -match 'ReparsePoint'
    }

    foreach ($target in $Targets) {
        $fullTargetPath = "C:\Users\$UserName\$OneDriveFolderName\$target"

        if (Test-Path -Path $fullTargetPath) {
            if (Test-Junction -Path $fullTargetPath) {
                Remove-Item -Path $fullTargetPath -Force
                Write-Host "Removed junction point: $fullTargetPath"
            } else {
                Write-Host "ERROR: The path is not a junction point: $fullTargetPath" -ForegroundColor Red
            }
        } else {
            Write-Host "ERROR: Target path not found: $fullTargetPath" -ForegroundColor Red
        }
    }
}

$OneDriveFolderName = "OneDrive - Gurr Johns"
$SourceToTarget = @{
    "C:\Tempvaluation" = "Tempvaluation";
    "C:\BackupValuation" = "BackupValuation"
}
Set-JunctionPoints -OneDriveFolderName $OneDriveFolderName -SourceToTarget $SourceToTarget

#$OneDriveFolderName = "OneDrive - Company Name"
#$Targets = @("Tempvaluation", "BackupValuation")
#Remove-JunctionPoints -OneDriveFolderName $OneDriveFolderName -Targets $Targets

