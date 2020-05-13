function Add-ZoomMeeting {
    param(
        [Parameter(Mandatory = $true)] [datetime] $DateTime,
        [Parameter(Mandatory = $true)] [string] $MeetingID,
        [string] $Password
    )

    $Command = "-NoLogo -WindowStyle Hidden -Command ""& { Start-ZoomMeeting -MeetingID \""$($MeetingID)\"" -Password \""$($Password)\"" }"""

    $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $Command
    $Trigger = New-ScheduledTaskTrigger -Once -At $DateTime
    $Settings = New-ScheduledTaskSettingsSet -DeleteExpiredTaskAfter "30:00:00:00" -Hidden

    $Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings
    $Task.Triggers[0].EndBoundary = $DateTime.AddMinutes(60).ToString('s')

    $DateString = Get-Date -Date $DateTime -UFormat "%Y%m%d %H%M"
    Register-ScheduledTask -TaskName "ZoomZoom $DateString" -InputObject $Task
}
