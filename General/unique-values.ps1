# File     : UniqueValues.ps1
# Effect   : Takes a single column of data with repeated data, removes all duplicated values and outputs to console (which can be redirected)
# Run As   : Current User
# Author   : Dan White

param (
    [Parameter(Mandatory=$true, HelpMessage="Please provide the input file path.")]
    [string]$InputFilePath
)

if (!(Test-Path -Path $InputFilePath)) {
    Write-Host "The input file does not exist. Please provide a valid file path." -ForegroundColor Red
    exit 1
}

try {
    # Read the content of the input file and filter unique values
    $UniqueValues = Get-Content $InputFilePath | Sort-Object | Get-Unique -AsString

    # Output the unique values
    Write-Host "Unique values from the input file:" -ForegroundColor Green
    $UniqueValues
} catch {
    Write-Host "An error occurred while processing the input file." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
