<#
CdoRecurTypeDaily/0
CdoRecurTypeWeekly/1 DayOfWeekMask Interval
CdoRecurTypeMonthly/2 DayOfMonth Interval
CdoRecurTypeMonthlyNth/3 DayOfWeekMask Instance Interval
CdoRecurTypeYearly/5 DayOfMonth Interval MonthOfYear
CdoRecurTypeYearlyNth/6
#>

Add-Type -assembly "Microsoft.Office.Interop.Outlook"

#Defines the DefaultFolderID for the Calendar, and creates our new COM Object
$OutlookFolders = “Microsoft.Office.Interop.Outlook.OlDefaultFolders” -as [type]
$Outlook = New-Object -ComObject Outlook.Application
$NameSpace = $Outlook.GetNamespace('MAPI')
$Folder = $NameSpace.getDefaultFolder($OutlookFolders::olFolderCalendar)

$Folder.Items | Where-Object -Property IsRecurring -eq $True | Tee-Object -Variable recur
.GetRecurrencePattern()


Expand-RecurrencePattern

<#
CdoSunday/1

The appointment recurs on Sundays.

CdoMonday/2

The appointment recurs on Mondays.

CdoTuesday/4

The appointment recurs on Tuesdays.

CdoWednesday/8

The appointment recurs on Wednesdays.

CdoThursday/16

The appointment recurs on Thursdays.

CdoFriday/32

The appointment recurs on Fridays.

CdoSaturday/64

The appointment recurs on Saturdays.
#>

