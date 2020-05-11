param (
    [Parameter(Mandatory = $true)] [string] $NuGetApiKey
)

Publish-Module -Path ZoomZoom -NuGetApiKey $NuGetApiKey -Verbose
