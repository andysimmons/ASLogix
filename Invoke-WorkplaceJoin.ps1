[CmdletBinding()]
param (
    [IO.FileInfo]
    $Transcript = 'C:\temp\wpjoin-transcript.txt',

    [int]
    $MaxAttempts = 100,

    [int]
    $TimeoutSeconds = 5
)

if ($Transcript) { Start-Transcript -Path $Transcript }

function Invoke-WorkplaceJoin
{
    & "${env:PROGRAMFILES}\Microsoft Workplace Join\AutoWorkplace.exe" /join
}

function Get-WorkplaceJoinCertificate
{
    param (
        [Security.Cryptography.X509Certificates.X509Store]
        $Store = 'Cert:\CurrentUser\My',

        [String]
        $IssuerPattern = 'MS-Organization-Access'
    )

    Get-ChildItem -Path $Store.Name | Where-Object { $_.Issuer -Match $IssuerPattern }
}

if (-not (Get-WorkplaceJoinCertificate))
{
    $attemptCounter = 0
    do
    {
        $attemptCounter++
        Write-Output "[$((Get-Date).ToLongTimeString())] Hybrid Azure AD Domain Join (attempt $attemptCounter/$MaxAttempts)..."
        Invoke-WorkplaceJoin
        Start-Sleep -Seconds $TimeoutSeconds
    } until (Get-WorkplaceJoinCertificate -or ($attemptCounter -ge $MaxAttempts))
}

if (Get-WorkplaceJoinCertificate)
{ 
    Write-Output 'Workplace Join completed successfully' 
}
else
{ 
    throw [TimeoutException] "Workplace Join failed after $MaxAttempts attempts. Giving up." 
}