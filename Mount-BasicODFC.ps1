#Requires -RunAsAdministrator
#Requires -Version 4.0
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param (
    [IO.DirectoryInfo]
    $ShareRoot = '\\slbctxaldfs01\AppLayeringTestFS\ODFC',

    [string]
    $UserDir = $env:USERNAME,

    [string]
    $FileName = 'ODFC.vhd',

    [IO.DirectoryInfo]
    $MountPoint = 'C:\ODFC',

    [IO.DirectoryInfo]
    $TempDir = 'C:\Temp',

    [int]
    $MaxSizeMB = 5120
)

$dpScriptFile = [IO.FileInfo] "$TempDir\dpScript"
$transcriptFile = [IO.FileInfo] "$TempDir\odfcLog-login.txt"
$dpLogFile = [IO.FileInfo] "$TempDir\dpLog.txt"

Start-Transcript -Path $transcriptFile

# need to split this up into simpler functions
function Initialize-ODFC
{
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [IO.FileInfo] 
        $FilePath,

        [IO.DirectoryInfo]
        $MountPoint,

        [string]
        $VolumeLabel = 'ODFC',
        
        [int]
        $DiskPartColWidth = 11
    )

    $mountAttempted = $false

    if (!$MountPoint.Exists) 
    {
        if ($PSCmdlet.ShouldProcess($MountPoint.FullName, 'CREATE DIRECTORY'))
        {
            try
            {
                $niParams = @{
                    Path        = $MountPoint.FullName
                    ItemType    = 'Directory'
                    ErrorAction = 'Stop'
                    Force       = $true
                    WhatIf      = $false
                }
                $MountPoint = New-Item @niParams
            }
            catch 
            {
                Write-Error "Couldn't create mount point '$MountPoint'!" 
                throw $_.Exception
            }
        }
    }
    else
    {
        $odfcVolume = Get-WmiObject -Class Win32_Volume -filter "label='$VolumeLabel'" | 
            Where-Object { $_.Name -like "$MountPoint\" } |  Select-Object -First 1

        if ($odfcVolume)
        { 
            "{0} volume already mounted at {1}" -f $VolumeLabel, $MountPoint 
            return
        }
        elseif ($PSCmdlet.ShouldProcess($MountPoint, 'DELETE CONTENTS'))
        {
            # Have to empty this directory to mount anything here
            try { Remove-Item -Path "$MountPoint\*" -Recurse -Force -WhatIf:$false }
            catch { throw "Error clearing ODFC mount point '$MountPoint'! $($_.Exception.Message)" }
        }
    }


    if (!$FilePath.Exists)
    {
        if ($PSCmdlet.ShouldProcess($FilePath.FullName, 'INITIALIZE NEW CONTAINER'))
        {
            try
            {
                # We'll lose all error handling once we get to diskpart, so
                # do a dry run of file creation beforehand (then nuke it)
                $niParams = @{
                    Path        = $FilePath.FullName
                    ItemType    = 'File'
                    ErrorAction = 'Stop'
                    Force       = $true
                    WhatIf      = $false
                }
                New-Item @niParams | Remove-Item -Force -WhatIf:$false -ErrorAction Stop
            }
            catch
            {
                Write-Error "Couldn't create Outlook data file container '$FilePath'!"
                throw $_.Exception.Message
            }
            
            try
            {
                "Preparing Outlook data file container for first use... this may take a minute."

                $dpScript = @(
                    "create vdisk file='$FilePath' maximum=$MaxSizeMB type=expandable",
                    "select vdisk file='$FilePath'",
                    "attach vdisk",
                    "create partition primary",
                    "active",
                    "attributes volume set nodefaultdriveletter",
                    "automount disable",
                    "assign mount='$MountPoint'",
                    "format quick fs=ntfs label='$VolumeLabel'"
                )

                $dpScript | Out-File -FilePath $dpScriptFile -Encoding ascii -Force -ErrorAction Stop
                diskpart.exe /s $dpScriptFile > $dpLogFile
                $mountAttempted = $true
            }
            catch
            { 
                Write-Error "Failed to create roaming container '$FilePath'!"
                throw $_.Exception 
            }
        }
    }

    if (!$mountAttempted -and $PSCmdlet.ShouldProcess($FilePath, 'MOUNT')) 
    {
        
        try
        {
            # Win7 can't run the storage management cmdlets, and diskpart is awful, so I'm sorry if you're
            # trying to read this. Remounting requires multiple diskpart scripts/invocations.

            # build some patterns we can use to parse volume information from diskpart output
            if ($VolumeLabel.Length -gt $DiskPartColWidth)
            { 
                $volNamePattern = $VolumeLabel.Substring(0, $DiskPartColWidth) 
            }
            else { $volNamePattern = $VolumeLabel }                                          
            $volNamePattern = "^[\s]+Volume [\d]+.+$([regex]::Escape($volNamePattern))"
            $volNumberPattern = '(?<=^[\s]+Volume )[\d]+'

            # attach the VHD
            "Attaching Outlook data file container '$FilePath'"
            $dpScript = @(
                "select vdisk file='$FilePath'",
                "attach vdisk"
            )
            Invoke-DiskPart -Script $dpScript -ScriptFile "$dpScriptFile.1" -LogFile $dpLogFile

            # parse volume info
            "Retrieving volume information"
            $volInfo = (Invoke-DiskPart -Script 'list volume' -ScriptFile "$dpScriptFile.2") -match $volNamePattern
            "Volume info: $volInfo"
            $volNumber = [int] ([regex]::Match($volInfo, $volNumberPattern)).Value
            "Volume number: $volNumber"

            # mount up
            "Mounting '$FilePath' to '$MountPoint'"
            $dpScript = @(
                "select vdisk file='$FilePath'",
                "select partition 1",
                "rescan",
                "select volume $volNumber",
                "rescan",
                "assign mount='$MountPoint'"
            )  
            Invoke-DiskPart -Script $dpScript -ScriptFile "$dpScriptFile.3" -LogFile $dpLogFile

            $mountAttempted = $true
        }
        catch
        { 
            Write-Error "Failed to mount roaming container '$FilePath' at '$MountPoint'!"
            throw $_.Exception 
        }
    }

    # Have to do some really generic post-invocation error handling/logging
    if ($mountAttempted)
    {
        $volume = Get-WmiObject -Class Win32_Volume -Filter "label='$VolumeLabel'"

        if ($volume)
        { 
            #placeholder - write success event
            #if ($dpResult) { $dpResult | Write-Verbose }
            return $true 
        }
        else
        { 
            #placeholder - write error event 
            #if ($dpResult) { $dpResult | Write-Warning }
            return $false
        }
    }
}

function Invoke-DiskPart 
{
    [CmdletBinding()]
    param (
        [string[]] $Script,

        [IO.FileInfo] $ScriptFile,

        [IO.FileInfo] $LogFile
    )

    try
    {
        $Script | Out-File -FilePath $ScriptFile -Encoding ascii -Force -ErrorAction Stop

        if ($LogFile) { diskpart.exe /s $ScriptFile >> $LogFile }
        else { diskpart.exe /s $ScriptFile }
    }
    catch { throw $_.Exception }
}

Initialize-ODFC -FilePath "$ShareRoot\$UserDir\$FileName" -MountPoint $MountPoint
