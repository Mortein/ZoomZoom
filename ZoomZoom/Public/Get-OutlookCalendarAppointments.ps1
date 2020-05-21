<#
.SYNOPSIS
    Returns meeting objects from Outlook
.DESCRIPTION
    Creates an Outlook COM application and returns all objects with message class IPM.Appointment from the Calendar Folder from the specified date range. If no date range is specified, all meetings from midnight the previous day, to the current time are returned
.EXAMPLE
    Get-OutlookCalendarAppointments
    Returns Outlook calendar appointments for the next 48 hours

    Get-OutlookCalendarAppointments -Start 12/01/2020 -End 21/02/2020
    Returns all non-recurring meetings between 12th January and 21st February 2020

.OUTPUTS
    PSCustomObject
.NOTES

#>
function Get-OutlookCalendarAppointments {

    [CmdletBinding(DefaultParameterSetName = 'Default')]

    param (

        # The start date for the search
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'SearchTime')]
        $Start,

        # The end date for the search
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'SearchTime')]
        $End
    )

    begin {
        #Add the required Assembly to PowerShell before processing input
        Add-Type -assembly "Microsoft.Office.Interop.Outlook"

        #Defines the DefaultFolderID for the Calendar, and creates our new COM Object
        $OutlookFolders = 'Microsoft.Office.Interop.Outlook.OlDefaultFolders' -as [type]
        $Outlook = New-Object -ComObject Outlook.Application
        $NameSpace = $Outlook.GetNamespace('MAPI')

        #Empty array to store our results
        $Meetings = @()
        $Results = @()
        if ($PSCmdlet.ParameterSetName -eq 'Default') {
            $Start = Get-Date
            $End = $Start.AddDays(2)
        }
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
        foreach ($Meeting in $Meetings) {
            $Results += [PSCustomObject]@{
                MeetingTime    = $Meeting.Start
                MeetingSubject = $Meeting.Subject
                MeetingBody    = $Meeting.Body
                MeetingURL     = Get-ZoomStringFromBody $Meeting.Body
                Recurring      = $false
            }
        }
        if ($Null -eq $RecurringAppointments) {
            return $Results
        }
        else {
            Foreach ($Appointment in $RecurringAppointments) {
                $AppointmentDates = @()
                $AppointmentDates += (Get-RecurringMeetingDates -Appointment $Appointment -EndDate (Get-Date $End).ToString()).AppointmentDates
                foreach ($Date in $AppointmentDates) {
                    if ($Date -gt (Get-Date $Start)) {
                        $Results += [pscustomobject]@{
                            MeetingTime    = $Date
                            MeetingSubject = $Appointment.Subject
                            MeetingBody    = $Appointment.Body
                            MeetingURL     = Get-ZoomStringFromBody $Appointment.Body
                            Recurring      = $true
                        }
                    }
                }
            }

            $defaultProperties = @('MeetingTime', 'MeetingSubject', 'MeetingURL', 'Recurring')
            $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’, [string[]]$defaultProperties)
            $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
            $Results | Add-Member MemberSet PSStandardMembers $PSStandardMembers

            return $Results
        }
    }
}
