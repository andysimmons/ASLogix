#Requires -RunAsAdministrator
#Requires -Version 4.0
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param (
    [IO.DirectoryInfo]
    $ShareRoot = '\\slbctxdfs01\AppLayeringTestFS\ODFC',

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

$transcriptFile = [IO.FileInfo] "$TempDir\odfcLog-login.txt"
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

        [IO.DirectoryInfo]
        $TempDir = $TempDir,
        
        [int]
        $DiskPartColWidth = 11
    )

    $mountAttempted = $false
    $dpScriptFile = [IO.FileInfo] "$TempDir\dpScript"
    $dpLogFile = [IO.FileInfo] "$TempDir\dpLog.txt"
    
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
                Invoke-DiskPart -Script $dpScript -ScriptFile "$dpScriptFile.txt" -LogFile $dpLogFile
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
            # Win7 can't run the storage management cmdlets, meaning we need diskpart scripts. Apologies
            # to anyone reading this. Remounting requires multiple diskpart scripts/invocations, along
            # with some excessive rescanning to keep up with all the async operations.

            # build some patterns we can use to parse volume information from diskpart output
            if ($VolumeLabel.Length -gt $DiskPartColWidth)
            { 
                $volNamePattern = $VolumeLabel.Substring(0, $DiskPartColWidth) 
            }
            else { $volNamePattern = $VolumeLabel }                                          
            $volNamePattern = "^[\s]+Volume [\d]+.+$([regex]::Escape($volNamePattern))"
            $volNumberPattern = '(?<=^[\s]+Volume )[\d]+'

            # diskpart script 1 - attach the VHD
            "Attaching Outlook data file container '$FilePath'"
            $dpScript = @(
                "select vdisk file='$FilePath'",
                'attach vdisk',
                'rescan'
            )
            Invoke-DiskPart -Script $dpScript -ScriptFile "$dpScriptFile-1.txt" -LogFile $dpLogFile

            # diskpart script 2 - parse volume info 
            # These back-to-back rescans are deliberate. For some reason, diskpart
            # doesn't see the new volume without 2 of them ...
            "Retrieving volume information"
            $dpScript = @(
                'rescan',
                'list volume'
            )
            $volInfo = (Invoke-DiskPart -Script $dpScript -ScriptFile "$dpScriptFile-2.txt") -match $volNamePattern
            "Volume info: $volInfo"
            $volNumber = [int] ([regex]::Match($volInfo, $volNumberPattern)).Value
            "Volume number: $volNumber"

            # diskpart script 3 - mount up
            "Mounting '$FilePath' to '$MountPoint'"
            $dpScript = @(
                "select vdisk file='$FilePath'",
                "select volume $volNumber",
                'rescan',
                "assign mount='$MountPoint'"
            )  
            Invoke-DiskPart -Script $dpScript -ScriptFile "$dpScriptFile-3.txt" -LogFile $dpLogFile

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
