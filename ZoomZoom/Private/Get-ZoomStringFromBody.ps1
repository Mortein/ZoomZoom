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

    From: https://support.zoom.us/hc/en-us/articles/201362373-What-is-a-Meeting-ID-
    The meeting ID is the meeting number associated with an instant or scheduled meeting. The meeting ID can be a 10 or 11-digit number.
    The 11-digit number is used for instant, scheduled or recurring meetings. The 10-digit number is used for Personal Meeting IDs.
    Meetings scheduled prior to April 12, 2020 may be 9-digits long.

    Regular expressions are used to find zoom URLs

    http(s|):\/\/(.*\.)*zoom.us\/j\/\d{9,11}(\?pwd=\w*)?
    Match protocol http:// or https://
    Optionally match Zoom vanity subdomain followed by a dot (e.g firstschool.)
    Match zoom.us/j/
    Match 9-11 digit meeting number
    Optionally match password, ?pwd=PLAINTEXTSTRING
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
        switch -Regex ($Body) {
            # Match zoom strings with/without passwords
            "http(s|):\/\/(.*\.)*zoom.us\/j\/\d{9,11}(\?pwd=\w*)?" {
                return (Select-String -InputObject $Body -Pattern "http(s|):\/\/(.*\.)*zoom.us\/j\/\d{9,11}(\?pwd=\w*)?").Matches.Value
            }
            default {
                break
            }
        }
    }

    end {
    }
}
