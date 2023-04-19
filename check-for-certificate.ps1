# File     : check-stuff.ps1
# Effect   : Check for the existence of a certificate called "ISRG Root X1" in the "Trusted Root Certification Authorities" store
# Use-case : Made for https://community.letsencrypt.org/t/email-from-slack-regarding-isrg-root-x1/196186
#          : Can be used to check for any certificate in the store.
# Run As   : Administrator
# Author   : Dan White


# Define the certificate store and certificate name
$storeName = "Root"
$certificateName = "ISRG Root X1"

# Open the certificate store
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName, "LocalMachine")
$store.Open("ReadOnly")

# Check if the certificate exists in the store
$certificateExists = $store.Certificates | Where-Object { $_.Subject -match $certificateName }

# Close the certificate store
$store.Close()

if ($certificateExists) {
    Write-Host "Certificate '$certificateName' found in the '$storeName' certificate store."
} else {
    Write-Host "Certificate '$certificateName' not found in the '$storeName' certificate store."
}