<#
.SYNOPSIS
    Searches a string for a Zoom Meeting invite URL and returns the first instance found
.DESCRIPTION
    Searches a string of text for any Zoom URL's based off the standard Zoom uses. Returns the first line found and splits this incase
    there are multiple URLS in 1 line (Common for hyperlink formats)
.INPUTS
    String
.OUTPUTS
    String
.NOTES
    Used to scrape meeting object bodies for Zoom meeting URLs
#>
function Get-ZoomStringFromBody {
    [CmdletBinding()]
    param (
        # A body of text that will contain a Zoom URL
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Body
    )

    begin {
    }

    process {
        $SplitString = $Body -Split ([Environment]::NewLine)
        Foreach ($String in $SplitString) {
            if ($String -match "\D{5,6}\/\/\S*\/j\/\d{10,11}\?{1}\S*") {
                $Results = (Select-String -InputObject $String -Pattern "\D{5,6}\/\/\S*\/j\/\d{10,11}\?{1}\S*").Matches.Value
                break
            }
            elseif ($String -match "\D{5,6}\/\/\S*\/j\/\d{10,11}") {
                $results = (Select-String -InputObject $String -Pattern "\D{5,6}\/\/\S*\/j\/\d{10,11}").Matches.Value
                break
            }
            else {
                $results = $null
            }
        }
    }

    end {
        return $results
    }
}
