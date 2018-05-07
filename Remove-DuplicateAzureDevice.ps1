#requires -Version 4
#requires -Module AzureADPreview
[CmdletBinding(SupportsShouldProcess)]
param (
    [ValidateScript({$_ -ge 0})]
    [int]
    $DeviceLimit = 0,
    
    [ValidateScript({$_ -ge 1})]
    [int]
    $MaxAgeInDays = 14
)

#region functions
function Get-NonPersistentAzureADDevice
{
    [CmdletBinding()]
    param (
        [string]
        $Pattern = 'XD[BT][PNDT]\d{2}(?!PERS).*',

        [string]
        $ODataFilter = "startswith(displayName,'xd')"
    )

    
    Write-Verbose "[$(Get-Date -format G)] Pulling list of Azure AD devices. This may take a few minutes."
    $aadDevices = Get-AzureADDevice -Filter $ODataFilter -All:$true

    # return the interesting devices 
    $aadDevices.Where({$_.DisplayName -match $Pattern})
}

function Select-StaleAzureADDevice
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [int]
        $MaxAgeInDays,

        [Parameter(Mandatory)]
        [Microsoft.Open.AzureAD.Model.Device[]] 
        $Device
    )
    $now = Get-Date

    Write-Verbose "[$(Get-Date -format G)] Grouping duplicate device objects. This may also take a few minutes."
    $groups = $Device | Group-Object -Property 'DisplayName'
    $i = 0

    foreach ($g in $groups)
    {
        $i++
        $wpParams = @{
            Activity        = "Checking AAD device freshness"
            Status          = ("Group {0}/{1} ({2})" -f $i, $groups.Count, $g.Group[0].DisplayName)
            PercentComplete = ($i / $groups.Count) * 100
        }
        Write-Progress @wpParams

        $devices = @($g.Group | Sort-Object -Property 'ApproximateLastLogonTimestamp' -Descending)

        if (($now - $devices[0].ApproximateLastLogonTimestamp).TotalDays -le $MaxAgeInDays)
        {
            # if the most recent record isn't stale, skip it
            $devices | Select-Object -Skip 1
        }
        else { $devices }
    }
}

# Wrap Remove-AzureADDevice to implement ShouldProcess (-WhatIf/-Confirm)
function Remove-AzureADDeviceSP
{
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [guid[]]
        $ObjectId
    )
    
    process
    {
        foreach ($id in $ObjectId)
        {
            if ($PSCmdlet.ShouldProcess($id, 'Remove')) 
            {
                # Remove-AzureADDevice -ObjectId $id
                Write-Warning "Placeholder: Remove-AzureADDevice -ObjectId $id"
            }
        }
    }
}
#endregion functions

#region main
$npDevices = Get-NonPersistentAzureADDevice
$staleNonPersistents = @(Select-StaleAzureADDevice -Device $npDevices -MaxAgeInDays $MaxAgeInDays)

if (!$staleNonPersistents)
{
    'Nothing to do.'
    exit
}

# enforce device limit if specified
if ($DeviceLimit -and ($DeviceLimit -lt $staleNonPersistents.Count)) 
{
    $remainder = $staleNonPersistents.Count - $DeviceLimit
    Write-Warning "Ignorning $remainder stale AAD devices! Cleanup limited to the first $DeviceLimit."
    $staleNonPersistents = @($staleNonPersistents | Select-Object -First $DeviceLimit)
}

Write-Verbose "[$(Get-Date -format G)] Found $($staleNonPersistents.Count) device(s) eligible for removal."

# remove devices (via wrapper function)
$staleNonPersistents | Remove-AzureADDeviceSP
#endregion main