# The input string
$inputString = "example, name <name@example.com>; Othername <anothername@otherexample.com>; "

# Common and shared email domains
$commonDomains = @('hotmail.com', 'outlook.com', 'gmail.com')

# Split the string into individual entries
$entries = $inputString -split ';'

# Create a hashtable to hold the unique results
$uniqueResults = @{}

foreach ($entry in $entries) {
    # Trim white spaces
    $entry = $entry.Trim()

    # Skip empty entries
    if (-not $entry) { continue }

    # Extract the email address
    $email = if ($entry -match '<(.*?)>') { $matches[1] } else { $entry -replace '.* ', '' }

    # Extract the domain from the email
    $domain = $email -replace '.*@', ''

    # Determine the output (full email or just domain)
    $output = if ($domain -in $commonDomains) { $email } else { $domain }

    # Add to the hashtable if not already present
    $uniqueResults[$output] = $true
}

# Output the results to a file
$uniqueResults.Keys | Out-File -FilePath "unique_domains.txt"
