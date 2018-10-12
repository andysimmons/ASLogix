#requires -Version 4
#requires -Module AzureADPreview
<#
.NOTES
    Name:   Remove-DuplicateAzureDevice.ps1
    Author: Andy Simmons
    Date:   10/12/2018

.SYNOPSIS
    Removes stale desktop devices from Azure AD.

.DESCRIPTION
    Non-persistent VDI machines are re-joined to Azure each time a user logs
    in. Each join creates a new object in Azure AD. We limit the number of devices
    each user is allowed to have joined to Azure at any given time, so periodic 
    cleanup is required.
    
    Stale records are identified by retrieving every non-persistent device from
    Azure AD, comparing the approximate last logon timestamp for devices with
    identical display names, and comparing the approximate last logon timestamp
    to the current time. 

    In October 2018, the scope of this cleanup was expanded to include persistent
    desktop devices, both physical and virtual, and still includes non-persistent
    VDI. The logic is very similar, but the initial scope is limited to devices
    with a trust type of "ServerAD", which are SL1 domain-joined boxes that have 
    hybrid domain joined to AzureAD.

    By default, self-imposed throttling is used because Remove-AzureADDevice
    doesn't take a collection as input, and using the pipeline results in one
    job per device removal. At the time of this writing, only 100 jobs can be
    submitted per 30 seconds per automation account.

    See the Azure Subscription Limits document for current information:
    https://docs.microsoft.com/en-us/azure/azure-subscription-service-limits

.PARAMETER DeviceLimit
    Maximum number of devices to be removed. The default value is 0 (no limit).

.PARAMETER MaxAgeInDays
    Maximum number of days a device is allowed to go without logging in, before
    it's considered stale.

.PARAMETER NewJobLimit
    Maximum number of jobs to submit at a time.

.PARAMETER NewJobIntervalSeconds
    Interval between submitting new device removal jobs.

.EXAMPLE
    Remove-DuplicateAzureDevice.ps1 -MaxAgeInDays 28 -NewJobLimit 80 -WhatIf -Verbose

    Analyzes devices. In addition
    to comparing relative login timestamps across duplicate device records, it will return every 
    device that hasn't logged in within the past 7 days. 
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [ValidateScript({$_ -ge 0})]
    [int]
    $DeviceLimit = 0,
    
    [ValidateScript({$_ -ge 1})]
    [int]
    $MaxAgeInDays = 30,

    [ValidateScript({($_ -ge 1) -and ($_ -le 100)})]
    [int]
    $NewJobLimit = 33,

    [ValidateScript({$_ -ge 1})]
    [int]
    $NewJobIntervalSeconds = 30
)

#region functions
<#
.SYNOPSIS
    Retrieves non-persistent desktop devices from Azure.
.NOTES
    This function is not called as of 10/2018, but I'm leaving it here
    for now. It's replaced by Get-AzureADDeviceByTrustType.
#> 
function Get-NonPersistentAzureADDevice
{
    [CmdletBinding()]
    param (
        [string]
        $Pattern = 'XD[BT][PNDT]\d{2}(?!PERS).*',

        [string]
        $ODataFilter = "startswith(displayName,'xd')"
    )
  
    $aadDevices = Get-AzureADDevice -Filter $ODataFilter -All:$true

    # return the interesting devices 
    $aadDevices.Where({$_.DisplayName -match $Pattern})
}

<#
.SYNOPSIS
    The filter support on these is pretty limited. Can't filter by deviceTrustType
    at the time of this writing.

#>
function Get-AzureADDeviceByTrustType
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string[]]
        $TrustType
    )
        
    $allDevices = Get-AzureADDevice -All:$true

    $allDevices.Where({$_.DeviceTrustType -in $TrustType})
}

<#
.SYNOPSIS
    Identifies stale device records in Azure AD.
#>
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

    Write-Verbose "[$(Get-Date -f G)] Grouping duplicate device objects."
    $groups = $Device | Sort-Object -Property 'DisplayName' | Group-Object -Property 'DisplayName'
    $i = 0

    Write-Verbose "[$(Get-Date -f G)] Identifying stale devices."
    foreach ($g in $groups)
    {
        $i++
        $wpParams = @{
            Activity        = "Checking AAD device freshness"
            Status          = ("Group {0}/{1} ({2})" -f $i, $groups.Count, $g.Group[0].DisplayName)
            PercentComplete = ($i / $groups.Count) * 100
        }
        Write-Progress @wpParams

        $devices = @($g.Group.Where({$_.ApproximateLastLogonTimestamp}) | Sort-Object -Property 'ApproximateLastLogonTimestamp' -Descending)

        if ($devices)
        {
            if (($now - $devices[0].ApproximateLastLogonTimestamp).TotalDays -le $MaxAgeInDays)
            {
                # if the most recent record isn't stale, skip it
                $devices | Select-Object -Skip 1
            }
            else { $devices }
        }
    }
}

<#
.SYNOPSIS
    Wraps Remove-AzureADDevice to implement ShouldProcess().

.DESCRIPTION
    Remove-AzureADDevice is destructive by design, and the AzureADPreview module has not
    yet implemented -WhatIf/-Confirm functionality, so we'll wrap this on.

.EXAMPLE
    Get-AzureADDevice | Remove-AzureADDeviceSP

    Removes the devices retrieved by Get-AzureADDevice, implementing the standard
    PowerShell confirmation behavior (prompting, honoring $WhatIfPreference, etc)
#> 
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
                try
                {
                    Remove-AzureADDevice -ObjectId "$id" -ErrorAction Stop
                    Write-Verbose "Device removed: $id"
                }
                catch 
                {
                    Write-Warning "Device removal failed: $id"
                    Write-Warning $_.Exception.Message
                }
            }
        }
    }
}

<#
.SYNOPSIS
    Throttles the rate at which objects are passed down the pipeline.
#>
function Limit-Pipeline
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Object[]]
        $InputObject,

        [Parameter(Mandatory)]
        [int]
        $Limit,
    
        [Parameter(Mandatory, ParameterSetName = 'Seconds')]
        [int]
        $Seconds,

        [Parameter(Mandatory, ParameterSetName = 'Milliseconds')]
        $Milliseconds
    )

    begin 
    {
        switch ($PSCmdlet.ParameterSetName)
        {
            'Seconds'
            { 
                $timesliceMs = 1000 * $Seconds
                $timesliceText = "$Seconds sec"
            }
            'Milliseconds'
            { 
                $timesliceMs = $Milliseconds
                $timesliceText = "$Milliseconds ms" 
            }
        }
        Write-Verbose "[$(Get-Date -f G)] Throttling pipeline to $Limit objects per $timesliceText."
        $counter = 0
        $thisBatchStart = Get-Date
        $nextBatchStart = ($thisBatchStart).AddMilliseconds($timesliceMs)
    }

    process
    {
        foreach ($object in $InputObject)
        {
            # check throughput vs. bandwidth after each batch
            if ($counter -ge $Limit)
            {
                $now = Get-Date
                $delayMs = ($nextBatchStart - $now).TotalMilliseconds -as [int]
                $velocityPct = [Math]::Round(100 * $timesliceMs / ($now - $thisBatchStart).TotalMilliseconds, 1)

                if ($delayMs -gt 0) 
                {
                    Write-Verbose "[$(Get-Date -f G)] Pipeline velocity at $velocityPct% capacity. Blocked for $delayMs ms."
                    Start-Sleep -Milliseconds $delayMs
                    Write-Verbose "[$(Get-Date -f G)] Pipeline open."
                }
                else { Write-Verbose "[$(Get-Date -f G)] Pipeline velocity at $velocityPct% capacity." }
             
                # reset counter/timer for next batch of objects
                $thisBatchStart = (Get-Date)
                $nextBatchStart = ($thisBatchStart).AddMilliseconds($timesliceMs)
                $counter = 0
            }
            $object
            $counter++
        }
    }
}

<#
.SYNOPSIS
    Determines if the current PowerShell session is connected to AzureAD.
#>
function Test-AzureADConnection 
{
    try 
    { 
        Get-AzureADDomain -ErrorAction Stop | Out-Null
        return $true
    }
    catch { return $false }
}
#endregion functions

#region main
$alreadyConnected = Test-AzureADConnection
if (!$alreadyConnected) 
{ 
    try
    {
        # Ha. Connect-AzureAD implements ShouldProcess() while Remove-AzureADDevice does not.
        # We need a connection to simulate everything else. Overriding $WhatIfPreference, and will
        # disconnect before exiting.
        Connect-AzureAD -WhatIf:$false -ErrorAction Stop
    }
    catch { throw "Couldn't connect to Azure AD. Bailing.`n$($_.Exception.Message)" }
}

Write-Output "[$(Get-Date -f G)] Pulling list of Azure AD devices with 'ServerAD' trust type. This may take a few minutes."
$sadDevices = Get-AzureADDeviceByTrustType -TrustType 'ServerAD'

Write-Output "[$(Get-Date -f G)] Analyzing $($sadDevices.Count) devices."
$staleSadDevices = @(Select-StaleAzureADDevice -Device $sadDevices -MaxAgeInDays $MaxAgeInDays)

if (!$staleSadDevices)
{
    Write-Output "[$(Get-Date -f G)] Nothing to do."
    if (!$alreadyConnected) { Disconnect-AzureAD -WhatIf:$false }
    exit
}

# enforce device limit if specified
if ($DeviceLimit -and ($DeviceLimit -lt $staleSadDevices.Count)) 
{
    $remainder = $staleSadDevices.Count - $DeviceLimit
    Write-Warning "Ignorning $remainder stale AAD devices! Cleanup limited to the first $DeviceLimit."
    $staleSadDevices = @($staleSadDevices | Select-Object -First $DeviceLimit)
}

Write-Output "[$(Get-Date -f G)] Found $($staleSadDevices.Count) device(s) eligible for removal."
$staleBreakdown = $staleSadDevices.DisplayName | 
    Sort-Object | 
    Group-Object -NoElement | 
    Sort-Object -Property Count -Descending | 
    Out-String
Write-Verbose "[$(Get-Date -f G)] Stale device breakdown:`n`n$staleBreakdown"

# remove devices
if ($NewJobLimit -or $NewJobIntervalSeconds)
{
    $staleSadDevices |
        Limit-Pipeline -Limit $NewJobLimit -Seconds $NewJobIntervalSeconds |
        Remove-AzureADDeviceSP
}
else { $staleSadDevices | Remove-AzureADDeviceSP }

if (!$alreadyConnected) { Disconnect-AzureAD -WhatIf:$false }
Write-Output "[$(Get-Date -f G)] Done."
#endregion main
