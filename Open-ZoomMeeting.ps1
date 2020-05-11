param(
    [Parameter(Mandatory = $true)] [string] $MeetingID,
    [String] $Password
)

Start-Process -FilePath "zoommtg://zoom.us/join?action=join&confno=$MeetingID&pwd=$Password"
