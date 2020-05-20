#RecurrenceType values
<#
CdoRecurTypeDaily/0
CdoRecurTypeWeekly/1 DayOfWeekMask Interval
CdoRecurTypeMonthly/2 DayOfMonth Interval
CdoRecurTypeMonthlyNth/3 DayOfWeekMask Instance Interval
CdoRecurTypeYearly/5 DayOfMonth Interval MonthOfYear
CdoRecurTypeYearlyNth/6
#>

function Get-RecurringMeetingDates {
    param (
        # An Outlook appointment item
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]
        $Appointment,

        # A DateArray from Get-DaysInRange that contains datetime objects
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [array]
        $DateArray
    )

    Switch (($Appointment.GetRecurrencePattern()).DayOfWeekMask) {
        #CdoMonday/2 - The appointment recurs on Mondays.
        2 {
            Return $DateArray | Where-Object -Property DayOfWeek -eq 'Monday'
        }
        #CdoTuesday/4 - The appointment recurs on Tuesdays.
        4 {
            Return $DateArray | Where-Object -Property DayOfWeek -eq 'Tuesday'
        }
        #CdoWednesday/8 - The appointment recurs on Wednesdays.
        8 {
            Return $DateArray | Where-Object -Property DayOfWeek -eq 'Wednesday'
        }
        #CdoThursday/16 - The appointment recurs on Thursdays.
        16 {
            Return $DateArray | Where-Object -Property DayOfWeek -eq 'Thursday'
        }
        #CdoFriday/32 - The appointment recurs on Fridays.
        32 {
            Return $DateArray | Where-Object -Property DayOfWeek -eq 'Friday'
        }
        #Meeting that occur on every week day have the value of 62 (2+4+8+16+32)
        62 {
            $Results += $DateArray | Where-Object -Property DayOfWeek -NotMatch ^[S*]
            return $Results
        }
    }
}
