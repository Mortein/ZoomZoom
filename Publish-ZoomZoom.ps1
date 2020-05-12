param (
    [Parameter(Mandatory = $true)] [string] $NuGetApiKey
)

$ModuleName = "ZoomZoom"

# Find out what the version should be
[version] $Version = "1.0.0"
$Module = Find-Module -Name "$ModuleName" -ErrorAction SilentlyContinue
if ($Module) {
    $Version = ("{0}.{1}.{2}" -f $Version.Major, $Version.Minor, $Version.Build + 1)
}

# Get the list of module functions
$Functions = @()
$Functions += Get-ChildItem -Path (Join-Path -Path "$ModuleName" -ChildPath "Public") -File | ForEach-Object { $_.BaseName }

# Update the manifest
Update-ModuleManifest -Path (Join-Path -Path "$ModuleName" -ChildPath "$ModuleName.psd1") -ModuleVersion $Version -FunctionsToExport $Functions

Publish-Module -Path ZoomZoom -NuGetApiKey $NuGetApiKey -Verbose
