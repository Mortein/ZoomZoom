function Get-ChildScripts {
    param (
        [Parameter(Mandatory = $true)] [string] $Directory
    )

    $Children = @( )

    $Children += if (Test-Path -Path (Join-Path -Path "$PSScriptRoot" -ChildPath "$Directory") -PathType Container) {
        @( Get-ChildItem -Path (Join-Path -Path "$PSScriptRoot" -ChildPath "$Directory" | Join-Path -ChildPath "*.ps1") )
    }

    return $Children
}

$Public = Get-ChildScripts -Directory "Public"
$Private = Get-ChildScripts -Directory "Private"

@($Public + $Private) | ForEach-Object {
    if ($_.Name -cnotmatch "\.Tests\.ps1$") {
        try {
            . $_.FullName
        }
        catch {
            Write-Error -Message "Failed to import function $($_.FullName): $_"
        }
    }
}

Export-ModuleMember -Function $Public.BaseName
