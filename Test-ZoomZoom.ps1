param (
    [string] $NuGetApiKey
)

# Check for required modules
"PSScriptAnalyzer" | ForEach-Object {
    $Module = Get-Module -Name "$_" -ListAvailable
    if (!$Module) {
        Write-Error -Message "Please install module $_"
        break
    }
}

Test-ModuleManifest -Path (Join-Path -Path $PSScriptRoot -ChildPath ZoomZoom | Join-Path -ChildPath "ZoomZoom.psd1")

Invoke-ScriptAnalyzer -Path ZoomZoom -ReportSummary -Recurse

if ($null -eq $NuGetApiKey) {
    Publish-Module -Path ZoomZoom -NuGetApiKey $NuGetApiKey -WhatIf -Verbose
}
