<#
.SYNOPSIS
    Returns meeting objects from Outlook
.DESCRIPTION
    Creates an Outlook COM application and returns all objects with message class IPM.Appointment from the Calendar Folder from the specified date range. If no date range is specified, all meetings from midnight the previous day, to the current time are returned
.EXAMPLE
    Get-OutlookCalendarAppointments
    Returns Outlook calendar appointments for the last 24 hours

    Get-OutlookCalendarAppointments -Start 12/01/2020 -End 21/02/2020
    Returns all non-recurring meetings between 12th January and 21st February 2020
.INPUTS
    DateTime
.OUTPUTS
    System.__ComObject
.NOTES

#>
function Get-OutlookCalendarAppointments {

    [CmdletBinding(DefaultParameterSetName = 'DefaultSearchTime')]

    param (

        # The start date for the search
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'SearchTime')]
        [string]
        $Start,

        # The end date for the search
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'SearchTime')]
        [string]
        $End
    )

    begin {
        #Add the required Assembly to PowerShell before processing input
        Add-Type -assembly "Microsoft.Office.Interop.Outlook"

        #Defines the DefaultFolderID for the Calendar, and creates our new COM Object
        $OutlookFolders = “Microsoft.Office.Interop.Outlook.OlDefaultFolders” -as [type]
        $Outlook = New-Object -ComObject Outlook.Application
        $NameSpace = $Outlook.GetNamespace('MAPI')

        #Empty array to store our results
        $Meetings = @()
        $RecurringMeetings = @()
    }

    process {
        $Folder = $NameSpace.getDefaultFolder($OutlookFolders::olFolderCalendar)
        $RecurringAppointments = $Folder.Items | Where-Object -Property IsRecurring -eq $True
        if ($null -eq $start) {
            $Meetings += $Folder.Items | Where-Object -Property Start -gt ((Get-Date -Hour 00 -Minute 00 -Second 00).AddDays(-1))
        }
        else {
            $Meetings += $Folder.Items | Where-Object -Property Start -gt (Get-Date $Start -Hour 00 -Minute 00 -Second 00) | Where-Object -Property End -lt (Get-Date $End -Hour 00 -Minute 00 -Second 00)
        }
    }

    end {
        if ($Null -eq $RecurringAppointments) {
            return $Meetings
        }
        else {
            Foreach ($Appointment in $RecurringAppointments) {
                Get-RecurringMeetingDates -Appointment $Appointment
            }
            return [PSCustomObject]@{
                Meetings          = $Meetings
                RecurringMeetings = $RecurringMeetings
            }
        }
    }
}
