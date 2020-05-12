param (
    [string] $NuGetApiKey
)

$ModuleName = "ZoomZoom"

# Check for required modules
"PSScriptAnalyzer" | ForEach-Object {
    $Module = Get-Module -Name "$_" -ListAvailable
    if (!$Module) {
        Write-Error -Message "Please install module $_"
        break
    }
}

Test-ModuleManifest -Path (Join-Path -Path $PSScriptRoot -ChildPath "$ModuleName" | Join-Path -ChildPath "$ModuleName.psd1")

Invoke-ScriptAnalyzer -Path "$ModuleName" -ReportSummary -Recurse

if ($null -eq $NuGetApiKey) {
    Publish-Module -Path $ModuleName -NuGetApiKey $NuGetApiKey -WhatIf -Verbose
}
