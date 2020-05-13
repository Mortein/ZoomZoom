function Update-CustomModule {
	[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]

	param (
		[Parameter(Mandatory = $true)] [string] $Name,
		[Parameter(Mandatory = $true)] [string] $Path
	)

	# Find out what the version should be
	[version] $Version = "1.0.0"
	$Module = Find-Module -Name "$Name" -ErrorAction SilentlyContinue
	if ($Module) {
		$Version = ("{0}.{1}.{2}" -f $Version.Major, $Version.Minor, $Version.Build + 1)
	}

	# Get the list of module functions
	$Functions = @()
	$Functions += Get-ChildItem -Path (Join-Path -Path "$Path" -ChildPath "Public") -File | ForEach-Object { $_.BaseName }

	if ($PSCmdlet.ShouldProcess("Manifest", "Update")) {
		# Update the manifest
		Update-ModuleManifest -Path (Join-Path -Path "$Path" -ChildPath "$Name.psd1") -ModuleVersion $Version -FunctionsToExport $Functions
	}
}

function Copy-CustomModule {
	param (
		[Parameter(Mandatory = $true)] [string] $Name,
		[Parameter(Mandatory = $true)] [string] $Directory
	)

	Remove-Item -Path "$Directory" -Recurse -Force -ErrorAction SilentlyContinue

	$TargetPath = Join-Path -Path "$Directory" -ChildPath "$Name"
	Copy-Item -Path "$Name" -Destination $TargetPath -Recurse

	return Get-Item -Path $TargetPath
}
