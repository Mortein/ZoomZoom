param (
    [Parameter(Mandatory = $true)] [string] $NuGetApiKey
)

# Import CustomModule for some functions
try {
    . .\CustomModule.ps1
}
catch {
    Write-Error $_
    exit 1
}

$ModuleName = "ZoomZoom"
$BuildDirectory = "Build"

$ModulePath = Copy-CustomModule -Name $ModuleName -Directory $BuildDirectory

# Update the module manifest
Update-CustomModule -Name $ModuleName -Path $ModulePath

Publish-Module -Path $ModulePath -NuGetApiKey $NuGetApiKey -Verbose
