# Check for the existence of a certificate called "ISRG Root X1" in the "Trusted Root Certification Authorities" store

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