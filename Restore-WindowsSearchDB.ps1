#Requires -RunAsAdministrator
#Requires -Version 4.0
[CmdletBinding(SupportsShouldProcess)]
param (
	[IO.DirectoryInfo]
	$ContainerPath = 'C:\ODFC',

	[IO.FileInfo]
	$Transcript = 'C:\Temp\RestoreWinSearchDB-transcript.txt'
)

if ($Transcript) { Start-Transcript -Path $Transcript }

function New-SymLink
{
	[CmdletBinding(SupportsShouldProcess)]
	param (
		[Parameter(Position = 0)]
		[IO.FileInfo] 
		$Link,

		[Parameter(Position = 1)]
		[IO.FileInfo]
		$Target
	)
	
	if ($PSCmdlet.ShouldProcess("$Link -> $Target", 'create symlink')) 
	{
		if ($PSVersionTable.PSVersion.Major -ge 5)
		{
			New-Item -Path $Link -ItemType SymbolicLink -Value $Target
		}
		else
		{
			if (Test-Path -Path $Target -PathType Container)
			{
				$command = 'cmd /c mklink /d'
			}
			else { $command = 'cmd /c mklink' }
			Invoke-Expression "$command $Link $Target"
		}
	}
}

Set-Service -Name 'WSearch' -StartupType 'Disabled'
Stop-Service -Name 'WSearch' -Force

[IO.DirectoryInfo] $winSearchDir = Split-Path (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows Search').DataDirectory

if ($winSearchDir.Exists -and $PSCmdlet.ShouldProcess($winSearchDir, 'RENAME'))
{
	Move-Item -Path "$winSearchDir" -Destination "$winSearchDir.old" -Force
}
if (-not $ContainerPath.Exists)
{ 
	New-Item -ItemType 'Directory' -Path $ContainerPath 
}
New-SymLink -Link $winSearchDir.FullName -Target $ContainerPath.FullName

Set-Service -Name 'WSearch' -StartupType 'Manual'
Start-Service -Name 'WSearch'