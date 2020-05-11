param(
    [Parameter(Mandatory = $true)] [string] $MeetingID,
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
