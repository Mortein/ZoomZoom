<#
.SYNOPSIS
    Returns an appointment object and all dates for that appointment based off its recurrence pattern
.DESCRIPTION
    Takes an Outlook Recurring Appointment object. Calculates all possible occurences of the appointment based off of
    the Start timestampt, and the Recurrence Pattern.
    The Recurrence Pattern defines a RecurrenceType, an integer that specifies if it's a daily, weekly, monthly, nth monthly,
    yearly, or nth yearly meeting.
    Each Typoe can then contain DayofWeek masks or WeekofMonth masks that confirm what day, days, or week they occur on.
.EXAMPLE
    $Appointment | Get-RecurringMeetingDates -EndDate 30/01/2020
    Returns all valid appointment dates for the input recurring appointment object
.INPUTS
    Outlook Recurring Appointment
.OUTPUTS
    pscustomobject
.NOTES

#>
function Get-RecurringMeetingDates {
    param (
        # An Outlook appointment item
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]
        $Appointment,

        # The date to calculate the last occurence of a recurring appointment
        [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
        $EndDate
    )

    $DateArray = Get-DaysinRange -StartDate ($Appointment.Start).ToString() -EndDate (Get-Date $EndDate).ToString()
    $Results = @()

    Switch (($Appointment.GetRecurrencePattern()).DayOfWeekMask) {
        #CdoMonday/2 - The appointment recurs on Mondays.
        2 {
            $DayOfWeekMask = $DateArray | Where-Object -Property DayOfWeek -eq 'Monday'
        }
        #CdoTuesday/4 - The appointment recurs on Tuesdays.
        4 {
            $DayOfWeekMask = $DateArray | Where-Object -Property DayOfWeek -eq 'Tuesday'
        }
        #CdoWednesday/8 - The appointment recurs on Wednesdays.
        8 {
            $DayOfWeekMask = $DateArray | Where-Object -Property DayOfWeek -eq 'Wednesday'
        }
        #CdoThursday/16 - The appointment recurs on Thursdays.
        16 {
            $DayOfWeekMask = $DateArray | Where-Object -Property DayOfWeek -eq 'Thursday'
        }
        #CdoFriday/32 - The appointment recurs on Fridays.
        32 {
            $DayOfWeekMask = $DateArray | Where-Object -Property DayOfWeek -eq 'Friday'
        }
        #Meeting that occur on every week day have the value of 62 (2+4+8+16+32)
        62 {
            $DayOfWeekMask += $DateArray | Where-Object -Property DayOfWeek -NotMatch ^[S*]
        }
    }

    Switch (($Appointment.GetRecurrencePattern()).RecurrenceType) {
        #CdoRecurTypeDaily/0 - Appointments that recur daily
        0 {
            $Results += $DateArray | Where-Object -Property DayOfWeek -NotMatch ^[S*]
            return $Results
        }
        #CdoRecurTypeWeekly/1 - Appointment recurs weekly (DayOfWeekMask,Interval)
        1 {
            $AppointmentDates = @()
            for ($i = 0; $i -lt ($DayOfWeekMask).Count; $i += (($Appointment.GetRecurrencePattern()).Interval)) {
                $AppointmentDates += $DayOfWeekMask[$i]
            }
            $Results += [PSCustomObject]@{
                Appointment      = $Appointment
                AppointmentDates = $AppointmentDates
            }
        }
        #CdoRecurTypeMonthly/2 - Appointment recurs monthly (DayOfMonth Interval)
        2 {

        }
        #CdoRecurTypeMonthlyNth/3 - Appointment recurs every Nth month
        3 {

        }
        #CdoRecurTypeYearly/5 - Appointment recurs every year
        5 {

        }
        #CdoRecurTypeYearlyNth/6 - Appointment recurs every Nth year
        6 {

        }
    }
    return $results
}
