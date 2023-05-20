$searchStrings = @("35312232809", "2045425661","1212278571","07380899670") # Replace with your search strings
$folderPath = "G:\Shared drives\Data\Phone\recordings" # Replace with the path to the folder where you want to search

# Function to convert filename to human-readable format
function ConvertTo-HumanReadable {
    param([string]$filename)

    $filename = $filename -replace "_", " " -replace "-", " "
    $matches = [regex]::Match($filename, "on (\d{8}) at (\d{6}) from")
    $date = [datetime]::ParseExact($matches.Groups[1].Value, "yyyyMMdd", $null)
    $time = [datetime]::ParseExact($matches.Groups[2].Value, "HHmmss", $null)

    $filename = $filename -replace $matches.Groups[1].Value, $date.ToString("dd/MM/yyyy")
    $filename = $filename -replace $matches.Groups[2].Value, $time.ToString("HH:mm:ss")

    return $filename.Trim(".mp3")
}

# Get files that match the search strings
#$matchingFiles = Get-ChildItem -Path $folderPath -Recurse -File | Where-Object { $file = $_.Name; $searchStrings | ForEach-Object { $file.Contains($_) } }

$matchingFiles = Get-ChildItem -Path $folderPath -Recurse -File | Where-Object { $file = $_.Name; ($searchStrings | Where-Object { $file.Contains($_) }).Count -gt 0 }


# Display original filenames
$originalFilenames = "original-filenames.txt"
Remove-Item -Path $originalFilenames -ErrorAction Ignore

Write-Host "Original filenames:"
foreach ($file in $matchingFiles) {
    Write-Host $file.Name
    Add-Content -Path $originalFilenames -Value $file.Name
}

Write-Host ""

# Display human-readable format
$humanReadableFilenames = "human-readable.txt"
Remove-Item -Path $humanReadableFilenames -ErrorAction Ignore

Write-Host "Human-readable format:"
foreach ($file in $matchingFiles) {
    $humanReadable = ConvertTo-HumanReadable -filename $file.Name
    Write-Host $humanReadable
    Add-Content -Path $humanReadableFilenames -Value $humanReadable
}
