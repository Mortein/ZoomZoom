<#
.SYNOPSIS
    Returns meeting objects from Outlook
.DESCRIPTION
    Creates an Outlook COM application and returns all objects with message class IPM.Appointment from the Calendar Folder from the specified date range
.EXAMPLE
    Get-OutlookCalendarAppointments
    Returns Outlook calendar appointments for the last 24 hours
.INPUTS
    DateTime
.OUTPUTS
    System.__ComObject
.NOTES
    General notes
#>
function Get-OutlookCalendarAppointments {

    [CmdletBinding()]

    param (

        # The start date for the search
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, parameterset = 'SearchTime')]
        [datetime]
        $Start,

        # The end date for the search
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, parameterSet = 'SearchTime')]
        [datetime]
        $End
    )

    begin {
        #Add the required Assembly to PowerShell before processing input
        Add-Type -assembly "Microsoft.Office.Interop.Outlook"

        #Defines the DefaultFolderID for the Calendar, and creates our new COM Object
        $CalendarFolderID = 9
        $Outlook = New-Object -ComObject Outlook.Application
        $NameSpace = $Outlook.GetNamespace('MAPI')

        #Results that will be returned
        $Meetings = @()
    }

    process {
        if ($null -eq $Start) {
            $Start = (Get-Date $Start).ToShortDateString()
        }
        if ($null -eq $End) {
            $End = (Get-Date $End).ToShortDateString()
        }

        $Filter = "[MessageClass]='IPM.Appointment' AND [Start] > '$Start' AND [End] < '$End'"

        #Finds all meetings in the date ranges provided and returns them as
        $Meetings += $NameSpace.GetDefaultFolder($CalendarFolderID).Items.Restrict($Filter)
    }

    end {
        return $Meetings
    }
}
