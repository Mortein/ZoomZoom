param (
    [string] $NuGetApiKey
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

$Errors = 0

# Check for and install required modules for testing
"PSScriptAnalyzer" | ForEach-Object {
    $Module = Get-Module -Name "$_" -ListAvailable
    if (!$Module) {
        Install-Module -Name $_ -Scope CurrentUser -Force
    }
}

# Module manifest
try {
    Test-ModuleManifest -Path (Join-Path -Path $ModulePath -ChildPath "$ModuleName.psd1") -ErrorAction Stop | Out-Null
}
catch {
    Write-Error $_
    $Errors += 1
}

# Script analyzer
$AnalyzerResults = Invoke-ScriptAnalyzer -Path $ModulePath -Severity 'Information', 'Warning', 'Error' -Recurse
$AnalyzerResults | ForEach-Object {
    [PSCustomObject]@{
        RuleName   = $_.RuleName
        ScriptName = $_.ScriptName
        Line       = $_.Line
    }

    if ($_.Severity -in 'Error') {
        $Errors += 1
    }
} | Sort-Object -Property 'ScriptName', 'Line'

# Publish
if ($NuGetApiKey) {
    try {
        Publish-Module -Path $ModulePath -NuGetApiKey $NuGetApiKey -WhatIf -Verbose -ErrorAction Stop
    }
    catch {
        Write-Error $_
        $Errors += 1
    }
}

exit $Errors
