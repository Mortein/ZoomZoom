<#
.SYNOPSIS
Join a Zoom Meeting

.PARAMETER MeetingID
ID of meeting to join, full meeting URL, or full meeting URL with password

.PARAMETER Password
Password to join meeting, if not included in MeetingID

.EXAMPLE
Start-ZoomMeeting -MeetingID "123-4444-5555"
Connect to meeting using only meeting ID

.EXAMPLE
Start-ZoomMeeting -MeetingID "123-4444-5555" -Password "422562"
Connect to meeting using meeting ID and password

.EXAMPLE
Start-ZoomMeeting -MeetingID 99922224444
Connect to meeting using only meeting ID

.EXAMPLE
Start-ZoomMeeting -MeetingID "https://zoom.us/j/77712348888"
Connect to meeting using URL

.EXAMPLE
Start-ZoomMeeting -MeetingID "https://zoom.us/j/88833335555?pwd=NE0vT2RkV25SWnBxMC9RRXFUc1F2dz09"
Connect to meeting using URL with password included
#>
function Start-ZoomMeeting {
    param(
        [Parameter(Mandatory = $true, Position = 0)] [string] $MeetingID,
        [String] $Password
    )

    # Extract Password from URL
    if (!$Password) {
        $Parts = $MeetingID -split "pwd="

        if ($Parts.Count -gt 1) {
            # There's a password included, use it
            $Password = $Parts[1]
        }
    }

    # Get the MeetingID, if they've supplied a URL
    $MeetingID = $MeetingID -replace "\?pwd=$Password"
    $MeetingIDParts = $MeetingID -split "/"
    $MeetingID = $MeetingIDParts[$MeetingIDParts.Count - 1]
    $MeetingID = $MeetingID -replace "-"

    Start-Process -FilePath "zoommtg://zoom.us/join?action=join&confno=$MeetingID&pwd=$Password"
}
