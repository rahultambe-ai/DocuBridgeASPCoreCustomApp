<#
 Creates the certificate the portal uses to sign in to Microsoft Entra ID as an APPLICATION.
 SharePoint REST (which the DMS uses) refuses app-only tokens that were obtained with a client secret;
 a certificate is required.

 1. Run (PowerShell, Windows):   .\New-AppCertificate.ps1 -Password "choose-a-strong-password"
 2. Entra admin centre > App registrations > <your app> > Certificates & secrets > Certificates > Upload  ->  DmsPortal.cer
 3. Put the .pfx somewhere outside the web root and set, with user-secrets or environment variables:
        dotnet user-secrets set "SharePoint:CertificatePath"     "C:\secure\DmsPortal.pfx"
        dotnet user-secrets set "SharePoint:CertificatePassword" "<the password from step 1>"
 Never commit the .pfx.
#>
param(
    [Parameter(Mandatory = $true)][string]$Password,
    [string]$Name = "DmsPortal",
    [string]$OutDir = ".",
    [int]$Years = 2
)

$cert = New-SelfSignedCertificate `
    -Subject "CN=$Name" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears($Years)

$secure = ConvertTo-SecureString -String $Password -Force -AsPlainText
Export-Certificate    -Cert $cert -FilePath (Join-Path $OutDir "$Name.cer") | Out-Null
Export-PfxCertificate -Cert $cert -FilePath (Join-Path $OutDir "$Name.pfx") -Password $secure | Out-Null

Write-Host "Created $Name.cer (upload to Entra) and $Name.pfx (keep secret). Thumbprint: $($cert.Thumbprint)"
