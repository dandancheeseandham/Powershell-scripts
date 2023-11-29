# Import CSV file
$data = Import-Csv -Path "uniquevalues-in-example.csv"

# Group by both columns and select the first item from each group
$uniqueData = $data | Group-Object Model_or_Version, CleanedModel | ForEach-Object { $_.Group[0] }

# Export unique pairings to new CSV file
$uniqueData | Export-Csv -Path "uniquevalues-out-example.csv" -NoTypeInformation
