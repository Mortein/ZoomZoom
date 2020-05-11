function Add-ZoomMeeting {
    [cmdletbinding(DefaultParameterSetName = "Today")]

    param(
        [Parameter(Mandatory = $true, ParameterSetName = "Tomorrow")] [switch] $Tomorrow,
        [Parameter(Mandatory = $true, ParameterSetName = "Date")] [String] $Date,
        [Parameter(Mandatory = $true)] [string] $Time,
        [Parameter(Mandatory = $true)] [string] $MeetingID,
        [string] $Password
    )
}
