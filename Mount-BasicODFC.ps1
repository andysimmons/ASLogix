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

try
{
	$dpScriptFile = [IO.FileInfo] "$TempDir\dpScript.txt"
	$transcriptFile = [IO.FileInfo] "$TempDir\odfcLog-login.txt"
	$dpLogFile = [IO.FileInfo] "$TempDir\dpLog.txt"
}
catch { throw $_.Exception }

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
		$VolumeLabel = 'ODFC'
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
					"create vdisk file='$($FilePath.FullName)' maximum=$MaxSizeMB type=expandable",
					"select vdisk file='$($FilePath.FullName)'",
					"attach vdisk",
					"create partition primary",
					"active",
					"attributes volume set nodefaultdriveletter",
					"automount disable",
					"assign mount='$($MountPoint.FullName)'",
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
	if (!$mountAttempted -and $PSCmdlet.ShouldProcess($FilePath.FullName, 'MOUNT')) 
	{
        
		try
		{
			# Win7 can't run the storage management cmdlets, and diskpart is awful, so remounting
			# is a bit of a chore here.

			# First, get a list of all volumes and parse out their numbers
<#			'list volume' | Out-File -FilePath $dpScriptFile -Encoding ascii -Force -ErrorAction Stop
			$volumeList = (diskpart /s $dpScriptFile) | Select-Object -Skip 7
            $volumeNumbers = [int[]]([regex]::Matches($volumeList, '(?<=(^|[^.])[\s]+Volume )[\d]+')).Value

            # Determine next available volume number, 
            $newVolumeNumber = ($volumeNumbers | Measure-Object -Maximum).Maximum + 1
            Remove-Item $dpScriptFile -Force -ErrorAction SilentlyContinue
#>



			# DiskPart will just blow through these asynchronously, so mounting is hit 
			# or miss unless you launch diskpart, attach, quit, relaunch, and mount.
			"Attaching Outlook data file container '$FilePath'"
			$dpScript = @(
				"select vdisk file='$($FilePath.FullName)'",
				"attach vdisk"
			)
			$dpScript | Out-File -FilePath $dpScriptFile -Encoding ascii -Force -ErrorAction Stop
			diskpart.exe /s $dpScriptFile > $dpLogFile
			Move-Item -Path $dpScriptFile -Destination "$dpScriptFile.pass1" -ErrorAction SilentlyContinue
            
			# In the ghettoooooooooo
			"Mounting '$FilePath' to '$MountPoint'"
			$dpScript = @(
				"select vdisk file='$($FilePath.FullName)'",
				"select partition 1",
				"rescan",
				"select volume $newVolumeNumber",
				"rescan",
				"assign mount='$($MountPoint.FullName)'"
			)  
			$dpScript | Out-File -FilePath $dpScriptFile -Encoding ascii -Force -ErrorAction Stop
			diskpart.exe /s $dpScriptFile >> $dpLogFile
			Move-Item -Path $dpScriptFile -Destination "$dpScriptFile.pass2" -ErrorAction SilentlyContinue
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

Initialize-ODFC -FilePath "$ShareRoot\$UserDir\$FileName" -MountPoint $MountPoint
