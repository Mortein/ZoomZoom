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

    [CmdletBinding(DefaultParameterSetName = 'DefaultSearchTime')]

    param (

        # The start date for the search
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'SearchTime')]
        [datetime]
        $Start,

        # The end date for the search
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'SearchTime')]
        [datetime]
        $End
    )

    begin {
        #Add the required Assembly to PowerShell before processing input
        Add-Type -assembly "Microsoft.Office.Interop.Outlook"

        #Defines the DefaultFolderID for the Calendar, and creates our new COM Object
        $OutlookFolders = “Microsoft.Office.Interop.Outlook.OlDefaultFolders” -as [type]
        $Outlook = New-Object -ComObject Outlook.Application
        $NameSpace = $Outlook.GetNamespace('MAPI')

        #Results that will be returned
        $Meetings = @()
    }

    process {
        $Folder = $NameSpace.getDefaultFolder($OutlookFolders::olFolderCalendar)

        if ($null -eq $start) {
            $Meetings += $Folder.Items | Where-Object -Property Start -gt ((Get-Date).AddDays(-1))
        }
        else {
            $Meetings += $Folder.Items | Where-Object -Property Start -gt $Start | Where-Object -Property End -lt $End
        }
    }
    #Process reucrring meetings

    end {
        return $Meetings
    }
}
